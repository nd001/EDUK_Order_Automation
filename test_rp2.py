import re
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright

RP2_URL = "https://gar.rpdns.co.uk/companyGAR/rp2.cgi"
TEST_ORDER = "1068814250-A"


def extract_web_ref(text):
    match = re.search(r"Web Ref:\s*([A-Za-z0-9\-]+)", text)
    return match.group(1) if match else None


def extract_customer_name(text):
    # Example:
    # LELL2953 • Mark Lelliott • Acton House ...
    parts = [part.strip() for part in text.split("•")]

    if len(parts) >= 2:
        return parts[1]

    return None


def extract_product(text):
    """
    Example RPii text:

    IB55732W (IB55 732 W UK)

    INDESIT IB55732WUK WH 50/50 FRIDGE FREEZER
    """

    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    sku = None
    description = None

    if lines:
        sku = lines[0].split()[0]

    # Find the first useful line after the SKU/model line
    for line in lines[1:]:
        if not line.startswith("("):
            description = line
            break

    return sku, description


def parse_delivery_block(text):
    """
    Example SGK:
        DELIVER SK
        BN:09
        070926:AM
        SK

    Example ED:
        DELIVER ED
        BN:09
        120926:PM
        ED
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
    slot = None
    customer_date = None

    if date_match:
        date_text = date_match.group(1)
        slot = date_match.group(2)

        rp2_date = datetime.strptime(
            date_text,
            "%d%m%y"
        ).date()

        if courier == "ED":
            customer_date = rp2_date - timedelta(days=1)
        else:
            customer_date = rp2_date

    return {
        "courier": courier,
        "rp2_date": rp2_date,
        "customer_date": customer_date,
        "slot": slot,
        "raw": cleaned,
    }


with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()

    page.goto(RP2_URL)

    print("RPii opened.")
    input(
        "Log into RPii and make sure you are on the main Sales screen, "
        "then press ENTER here..."
    )

    sales_frame = page.frame_locator("#SALESFrame")

    # -------------------------------------------------
    # Search for the test order
    # -------------------------------------------------

    sales_enquiry = sales_frame.locator("#orderno")
    sales_enquiry.wait_for(state="visible")

    print(f"\nOpening order {TEST_ORDER}...")

    sales_enquiry.fill(TEST_ORDER)
    sales_enquiry.press("Enter")

    # Give RPii a moment to update the sale screen
    page.wait_for_timeout(1500)

    # -------------------------------------------------
    # Read Web Ref
    # -------------------------------------------------

    web_panel = sales_frame.locator(
        "td.SALES-panel",
        has_text="Web Ref:"
    )

    web_panel_text = web_panel.inner_text()
    web_ref = extract_web_ref(web_panel_text)

    # -------------------------------------------------
    # Read customer
    # -------------------------------------------------

    customer_panel = sales_frame.locator(
        "td.SALES-panel-s",
        has_text="•"
    ).first

    customer_text = customer_panel.inner_text()
    customer_name = extract_customer_name(customer_text)

    # -------------------------------------------------
    # Read SKU
    # -------------------------------------------------
        # -------------------------------------------------
    # Read active product SKU + description
    # -------------------------------------------------

    sku = None
    description = None

    product_cells = sales_frame.locator(
        "td.SALES-lines-left-wrap"
    )

    product_cell_count = product_cells.count()

    # RPii presents product information in pairs:
    #
    # Cell 0 = SKU / stock information
    # Cell 1 = description
    # Cell 2 = next SKU / stock information
    # Cell 3 = next description
    #
    # Example active line:
    # IB55732W (IB55 732 W UK ) (1 -S L:W2)

    for i in range(0, product_cell_count - 1, 2):

        sku_cell = product_cells.nth(i)
        desc_cell = product_cells.nth(i + 1)

        sku_text = " ".join(
            sku_cell.inner_text().split()
        )

        description_text = " ".join(
            desc_cell.inner_text().split()
        )

        # Active RPii product lines contain a quantity
        # immediately after "(":
        #
        # (1 -S L:W2)
        #
        # Cancelled/problem lines such as WEBPART contain:
        #
        # ( - : )

        active_match = re.search(
            r"\(\s*(\d+)\s+-",
            sku_text
        )

        if not active_match:
            continue

        quantity = int(active_match.group(1))

        if quantity <= 0:
            continue

        # SKU is the first value in the cell
        sku = sku_text.split()[0]

        description = description_text

        break
    # -------------------------------------------------
    # Read financials
    # -------------------------------------------------

    total_sale = sales_frame.locator(
        "#qh_ordval"
    ).input_value()

    total_paid = sales_frame.locator(
        "#qh_totpaid"
    ).input_value()

    balance = sales_frame.locator(
        "#qh_balance"
    ).input_value()

    # -------------------------------------------------
    # Read courier / delivery block
    # -------------------------------------------------

    delivery_cells = sales_frame.locator("td.tcb")

    delivery_text = None

    for i in range(delivery_cells.count()):
        text = delivery_cells.nth(i).inner_text()

        if "DELIVER" in text:
            delivery_text = text
            break

    if delivery_text:
        delivery = parse_delivery_block(delivery_text)
    else:
        delivery = {
            "courier": None,
            "rp2_date": None,
            "customer_date": None,
            "slot": None,
            "raw": None,
        }

    # -------------------------------------------------
    # Report
    # -------------------------------------------------

    print("\n")
    print("=" * 60)
    print("RPii ORDER CHECK")
    print("=" * 60)

    print(f"Web Ref:       {web_ref}")
    print(f"Customer:      {customer_name}")
    print(f"SKU:           {sku}")
    print(f"Description:   {description}")
    print()
    
    print(f"Total Sale:    £{total_sale}")
    print(f"Total Paid:    £{total_paid}")
    print(f"Balance:       £{balance}")
    print()

    courier = delivery["courier"]

    if courier == "SK":
        print("Courier:       SGK Distribution")

    elif courier == "ED":
        print("Courier:       Electrical Discount UK")

    else:
        print(f"Courier:       UNKNOWN ({courier})")

    if delivery["rp2_date"]:
        print(
            "RPii Date:      "
            + delivery["rp2_date"].strftime("%d/%m/%Y")
        )

    if delivery["customer_date"]:
        if courier == "SK":
            print(
                "Collection:     "
                + delivery["customer_date"].strftime("%d/%m/%Y")
            )

        elif courier == "ED":
            print(
                "Delivery Date:  "
                + delivery["customer_date"].strftime("%d/%m/%Y")
            )

    if delivery["slot"]:
        print(f"Slot:          {delivery['slot']}")

    print("=" * 60)

    print("\nNo changes have been made to the sale.")

    input("\nPress ENTER to close...")
    browser.close()