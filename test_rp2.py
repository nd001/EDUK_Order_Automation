import os
import re
from datetime import datetime, timedelta

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright


# ============================================================
# CONFIGURATION
# ============================================================

# Change this depending on where you are running the program:
#
# "work"   = work network
# "remote" = home / remote connection
#
MODE = "remote"

WORK_URL = "https://gar.rpdns.co.uk/companyGAR/rp2.cgi"
REMOTE_URL = "https://garremote.rpdns.co.uk/companyGAR/rp2.cgi"

# Temporary test order.
# We will replace this with the B&Q Excel import next.
TEST_ORDER = "1068814250-A"


# ============================================================
# LOAD REMOTE CREDENTIALS
# ============================================================

load_dotenv()

RP2_REMOTE_USERNAME = os.getenv("RP2_REMOTE_USERNAME")
RP2_REMOTE_PASSWORD = os.getenv("RP2_REMOTE_PASSWORD")


if MODE == "remote":
    RP2_URL = REMOTE_URL

    if not RP2_REMOTE_USERNAME or not RP2_REMOTE_PASSWORD:
        raise RuntimeError(
            "Remote RPii username/password missing from .env file."
        )

elif MODE == "work":
    RP2_URL = WORK_URL

else:
    raise ValueError(
        'MODE must be either "work" or "remote".'
    )


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def extract_web_ref(text):
    """
    Extracts:

        Web Ref: 1068814250-A

    from the larger RPii sale information panel.
    """

    match = re.search(
        r"Web Ref:\s*([A-Za-z0-9\-]+)",
        text
    )

    return match.group(1) if match else None


def extract_customer_name(text):
    """
    RPii customer panel example:

        DIY0205 • DIY • ...

    For B&Q orders this is normally the generic
    RPii account such as DIY, rather than the actual
    delivery customer's name.
    """

    parts = [
        part.strip()
        for part in text.split("•")
    ]

    if len(parts) >= 2:
        return parts[1]

    return None


def parse_delivery_block(text):
    """
    Example SGK:

        DELIVER SK
        BN:09
        070926:AM
        SK

    Example EDUK:

        DELIVER ED
        BN:09
        120926:PM
        ED

    Business rule:

        SGK:
            use the RPii date exactly.

        ED:
            actual customer delivery date is
            ONE DAY BEFORE the RPii date.
    """

    cleaned = " ".join(text.split())

    courier = None

    if "DELIVER SK" in cleaned:
        courier = "SK"

    elif "DELIVER ED" in cleaned:
        courier = "ED"

    date_match = re.search(
        r"(\d{6}):(AM|PM)",
        cleaned
    )

    rp2_date = None
    customer_date = None
    slot = None

    if date_match:
        date_text = date_match.group(1)
        slot = date_match.group(2)

        rp2_date = datetime.strptime(
            date_text,
            "%d%m%y"
        ).date()

        if courier == "ED":
            customer_date = (
                rp2_date - timedelta(days=1)
            )

        else:
            customer_date = rp2_date

    return {
        "courier": courier,
        "rp2_date": rp2_date,
        "customer_date": customer_date,
        "slot": slot,
        "raw": cleaned,
    }


def read_active_product(sales_frame):
    """
    RPii presents sale product information in pairs:

        Cell 0 = SKU / stock information
        Cell 1 = description
        Cell 2 = next SKU
        Cell 3 = next description

    Example cancelled/import-problem line:

        WEBPART ( - : )

    Example active line:

        IB55732W (IB55 732 W UK) (1 -S L:W2)

    We therefore choose the first product line
    containing a positive quantity.
    """

    product_cells = sales_frame.locator(
        "td.SALES-lines-left-wrap"
    )

    product_cell_count = product_cells.count()

    sku = None
    description = None
    quantity = None

    for i in range(
        0,
        product_cell_count - 1,
        2
    ):
        sku_cell = product_cells.nth(i)
        desc_cell = product_cells.nth(i + 1)

        sku_text = " ".join(
            sku_cell.inner_text().split()
        )

        description_text = " ".join(
            desc_cell.inner_text().split()
        )

        # Active product example:
        #
        # (1 -S L:W2)
        #
        active_match = re.search(
            r"\(\s*(\d+)\s+-",
            sku_text
        )

        if not active_match:
            continue

        candidate_quantity = int(
            active_match.group(1)
        )

        if candidate_quantity <= 0:
            continue

        candidate_sku = sku_text.split()[0]

        sku = candidate_sku
        description = description_text
        quantity = candidate_quantity

        break

    return {
        "sku": sku,
        "description": description,
        "quantity": quantity,
    }


# ============================================================
# START PLAYWRIGHT
# ============================================================

with sync_playwright() as p:

    browser = p.chromium.launch(
        headless=False
    )

    # --------------------------------------------------------
    # Create browser context
    # --------------------------------------------------------

    if MODE == "remote":

        context = browser.new_context(
            http_credentials={
                "username": RP2_REMOTE_USERNAME,
                "password": RP2_REMOTE_PASSWORD,
            }
        )

    else:

        context = browser.new_context()

    page = context.new_page()

    # --------------------------------------------------------
    # Open RPii
    # --------------------------------------------------------

    print()
    print("=" * 60)
    print("EDUK RPii ORDER READER")
    print("=" * 60)
    print(f"Mode: {MODE.upper()}")
    print(f"Opening: {RP2_URL}")
    print()

        # --------------------------------------------------------
    # Open RPii and wait for the actual EPoS module
    # --------------------------------------------------------

    print("Opening RPii...")

    max_attempts = 3
    rp2_ready = False

    for attempt in range(1, max_attempts + 1):

        print(
            f"\nRPii load attempt "
            f"{attempt}/{max_attempts}..."
        )

        try:
            page.goto(
                RP2_URL,
                wait_until="domcontentloaded",
                timeout=30000
            )

            print(
                "Outer RPii page loaded. "
                "Waiting for Sales module..."
            )

            # RPii remote can be slow to initialise
            page.wait_for_selector(
                "#SALESFrame",
                state="attached",
                timeout=30000
            )

            sales_frame = page.frame_locator(
                "#SALESFrame"
            )

            sales_enquiry = sales_frame.locator(
                "#orderno"
            )

            # Give remote RPii plenty of time
            sales_enquiry.wait_for(
                state="visible",
                timeout=60000
            )

            rp2_ready = True

            print(
                "✓ RPii Sales module is ready."
            )

            break

        except Exception as error:

            print(
                "RPii did not finish loading."
            )

            print(
                f"Reason: {error}"
            )

            if attempt < max_attempts:

                print(
                    "Waiting 5 seconds before retry..."
                )

                page.wait_for_timeout(5000)

            else:

                print()
                print("=" * 60)
                print(
                    "RPii REMOTE CONNECTION FAILED"
                )
                print("=" * 60)
                print(
                    "The outer RPii screen loaded, "
                    "but the Sales module did not."
                )
                print()
                print(
                    "No order processing has taken place."
                )

    if not rp2_ready:

        input(
            "\nPress ENTER to close..."
        )

        browser.close()

        raise SystemExit

    input(
        "\nRPii is ready. "
        "Press ENTER to open the test order..."
    )

    # --------------------------------------------------------
    # Search for order
    # --------------------------------------------------------

    print(
        f"\nOpening order {TEST_ORDER}..."
    )

    sales_enquiry.fill(TEST_ORDER)
    sales_enquiry.press("Enter")

    # Wait until the sale total has appeared.
    # This is better than relying only on a fixed delay.
    total_sale_field = sales_frame.locator(
        "#qh_ordval"
    )

    total_sale_field.wait_for(
        state="visible",
        timeout=30000
    )

    # Small additional pause because RPii updates
    # several parts of the sale asynchronously.
    page.wait_for_timeout(1000)

    # ========================================================
    # READ WEB REFERENCE
    # ========================================================

    web_panel = sales_frame.locator(
        "td.SALES-panel",
        has_text="Web Ref:"
    )

    web_panel_text = web_panel.inner_text()

    web_ref = extract_web_ref(
        web_panel_text
    )

    # ========================================================
    # READ RPii CUSTOMER ACCOUNT
    # ========================================================

    customer_panel = sales_frame.locator(
        "td.SALES-panel-s",
        has_text="•"
    ).first

    customer_text = customer_panel.inner_text()

    customer_name = extract_customer_name(
        customer_text
    )

    # ========================================================
    # READ ACTIVE PRODUCT
    # ========================================================

    product = read_active_product(
        sales_frame
    )

    sku = product["sku"]
    description = product["description"]
    quantity = product["quantity"]

    # ========================================================
    # READ FINANCIALS
    # ========================================================

    total_sale = sales_frame.locator(
        "#qh_ordval"
    ).input_value()

    total_paid = sales_frame.locator(
        "#qh_totpaid"
    ).input_value()

    balance = sales_frame.locator(
        "#qh_balance"
    ).input_value()

    # ========================================================
    # READ COURIER / DELIVERY INFORMATION
    # ========================================================

    delivery_cells = sales_frame.locator(
        "td.tcb"
    )

    delivery_text = None

    for i in range(
        delivery_cells.count()
    ):
        text = delivery_cells.nth(
            i
        ).inner_text()

        if "DELIVER" in text:
            delivery_text = text
            break

    if delivery_text:

        delivery = parse_delivery_block(
            delivery_text
        )

    else:

        delivery = {
            "courier": None,
            "rp2_date": None,
            "customer_date": None,
            "slot": None,
            "raw": None,
        }

    # ========================================================
    # REPORT
    # ========================================================

    print()
    print()
    print("=" * 60)
    print("RPii ORDER CHECK")
    print("=" * 60)

    print(
        f"Web Ref:       {web_ref}"
    )

    print(
        f"RPii Account:  {customer_name}"
    )

    print(
        f"SKU:           {sku}"
    )

    print(
        f"Description:   {description}"
    )

    if quantity is not None:
        print(
            f"Quantity:      {quantity}"
        )

    print()

    print(
        f"Total Sale:    £{total_sale}"
    )

    print(
        f"Total Paid:    £{total_paid}"
    )

    print(
        f"Balance:       £{balance}"
    )

    print()

    courier = delivery["courier"]

    # --------------------------------------------------------
    # Courier reporting
    # --------------------------------------------------------

    if courier == "SK":

        print(
            "Courier:       SGK Distribution"
        )

    elif courier == "ED":

        print(
            "Courier:       Electrical Discount UK"
        )

    else:

        print(
            f"Courier:       UNKNOWN ({courier})"
        )

    # --------------------------------------------------------
    # RPii scheduled date
    # --------------------------------------------------------

    if delivery["rp2_date"]:

        print(
            "RPii Date:      "
            + delivery["rp2_date"].strftime(
                "%d/%m/%Y"
            )
        )

    # --------------------------------------------------------
    # Customer-facing date
    # --------------------------------------------------------

    if delivery["customer_date"]:

        if courier == "SK":

            print(
                "Collection:     "
                + delivery[
                    "customer_date"
                ].strftime(
                    "%d/%m/%Y"
                )
            )

        elif courier == "ED":

            print(
                "Delivery Date:  "
                + delivery[
                    "customer_date"
                ].strftime(
                    "%d/%m/%Y"
                )
            )

            print(
                "                "
                "(RPii date minus 1 day)"
            )

    # --------------------------------------------------------
    # AM / PM
    # --------------------------------------------------------

    if delivery["slot"]:

        print(
            f"Slot:          "
            f"{delivery['slot']}"
        )

    print("=" * 60)

    # ========================================================
    # BASIC SAFETY CHECK
    # ========================================================

    print()
    print("SAFETY CHECK")
    print("-" * 60)

    if web_ref == TEST_ORDER:

        print(
            "✓ Web Ref matches requested order."
        )

    else:

        print(
            "✗ WARNING: Web Ref does NOT match "
            "requested order."
        )

    if sku:

        print(
            f"✓ Active product found: {sku}"
        )

    else:

        print(
            "✗ WARNING: No active product "
            "could be identified."
        )

    if courier in ("SK", "ED"):

        print(
            "✓ Recognised delivery method."
        )

    else:

        print(
            "⚠ Delivery method not recognised."
        )

    print("-" * 60)

    print(
        "\nREAD-ONLY MODE: "
        "No changes have been made to the sale."
    )

    input(
        "\nPress ENTER to close..."
    )

    browser.close() 