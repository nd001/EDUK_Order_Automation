import os
import re
import time

from playwright.sync_api import sync_playwright

from config import (
    MODE,
    ORDERS_FILE,
    INVOICE_FOLDER,
    RP2_URL,
    RP2_REMOTE_USERNAME,
    RP2_REMOTE_PASSWORD,
    MIRAKL_BASE_URL,
    MIRAKL_API_KEY,
    DIY_DELIVERY_TOPIC_CODE,
)

from marketplaces.diy import (
    DIYMarketplaceAdapter,
    read_mirakl_order,
    read_mirakl_threads,
    build_mirakl_dry_run_payload,
    send_mirakl_delivery_message,
    build_mirakl_shipping_dry_run,
    ship_mirakl_order,
)

from messaging import (
    build_sgk_message,
    build_ed_message,
    build_sgk_sms,
    build_ed_sms,
)

from rp2 import (
    read_payment_summary,
    extract_web_ref,
    extract_customer_name,
    extract_delivery_customer,
    extract_uk_postcode,
    parse_delivery_block,
    read_active_product,
)

from validation import (
    normalise_phone,
    normalise_postcode,
    normalise_sku,
    normalise_price,
    build_three_way_safety_result,
    can_override_sku_only,
)

from datetime import datetime

# Temporary test order.
# We will replace this with the B&Q Excel import next.
TEST_ORDER = None

# Live Mirakl writes are enabled only after all safety gates pass.
# The operator must still type the exact SEND <order_id> confirmation.
LIVE_MIRAKL_SEND_ENABLED = True

SMS_TEST_MODE = False
SMS_TEST_MOBILE = "my mobile number"



# ============================================================
# HELPER FUNCTIONS
# ============================================================

def get_queue_status(mirakl_state):
    """
    Convert Mirakl order state into a simple
    operator-facing queue status.
    """

    if mirakl_state in (
        "SHIPPED",
        "WAITING_DEBIT",
        "RECEIVED",
        "CLOSED",
    ):
        return "✓ COMPLETE"

    if mirakl_state == "SHIPPING":
        return "→ READY"

    if mirakl_state:
        return f"⚠ {mirakl_state}"

    return "⚠ UNKNOWN"



def get_customer_surname(full_name):
    """
    Extract a safe surname for use in an invoice filename.

    Example:
        "mark booth" -> "Booth"
        "Jackie Jeeves" -> "Jeeves"
    """

    if not full_name:
        return "Unknown"

    parts = full_name.strip().split()

    if not parts:
        return "Unknown"

    surname = parts[-1]

    # Remove characters Windows does not allow in filenames
    surname = re.sub(
        r'[<>:"/\\|?*]',
        '',
        surname
    )

    return surname.title()


def complete_mirakl_shipping(
    browser,
    marketplace_order,
    mirakl_order,
):
    """
    Complete the controlled Mirakl shipping flow for one order.

    Safety design:
    - Build/display a dry run first.
    - Require exact SHIP <order_id> confirmation.
    - Perform exactly one shipment PUT request.
    - Never automatically retry a shipment write.
    - Read the order back and verify the post-shipment state.
    """

    shipping_dry_run = build_mirakl_shipping_dry_run(
        mirakl_order
    )

    print()
    print("=" * 60)
    print("MIRAKL SHIPPING - DRY RUN")
    print("=" * 60)

    print(
        f"Order:          "
        f"{shipping_dry_run['order_id']}"
    )

    print(
        f"Current State:  "
        f"{shipping_dry_run['current_state']}"
    )

    print(
        f"Target State:   "
        f"{shipping_dry_run['target_state']}"
    )

    print(
        f"Endpoint:       "
        f"{shipping_dry_run['endpoint']}"
    )

    print(
        f"Tracking Number: "
        f"{shipping_dry_run['tracking_number']}"
    )

    print(
        f"Carrier Code:    "
        f"{shipping_dry_run['carrier_code']}"
    )

    print(
        f"Carrier:         "
        f"{shipping_dry_run['carrier_label']}"
    )

    print()
    print("NO SHIPPING API WRITE HAS BEEN MADE.")

    # --------------------------------------------------------
    # Shipping dry-run safety checks
    # --------------------------------------------------------

    shipping_failures = []

    if (
        shipping_dry_run["order_id"]
        != marketplace_order.order_id
    ):
        shipping_failures.append(
            "Shipping order number mismatch"
        )

    if (
        shipping_dry_run["current_state"]
        != "SHIPPING"
    ):
        shipping_failures.append(
            "Mirakl order is not in SHIPPING state"
        )

    if (
        shipping_dry_run["target_state"]
        != "SHIPPED"
    ):
        shipping_failures.append(
            "Unexpected target shipping state"
        )

    if shipping_failures:

        print()
        print("✗ SHIPPING SAFETY CHECK FAILED")

        for failure in shipping_failures:
            print(
                f"✗ {failure}"
            )

        print()
        print(
            "NO SHIPPING API WRITE HAS BEEN MADE."
        )

        input(
            "\nPress ENTER to close..."
        )

        browser.close()
        raise SystemExit

    # ========================================================
    # FINAL MIRAKL SHIPPING GATE
    # ========================================================

    print()
    print("=" * 60)
    print("FINAL MIRAKL SHIPPING CONFIRMATION")
    print("=" * 60)

    print(
        f"Order:          "
        f"{marketplace_order.order_id}"
    )

    print(
        f"Current State:  "
        f"{shipping_dry_run['current_state']}"
    )

    print(
        f"Target State:   "
        f"{shipping_dry_run['target_state']}"
    )

    expected_confirmation = (
        f"SHIP {marketplace_order.order_id}"
    )

    print()
    print("=" * 60)
    print(
        "WARNING: THIS WILL MARK THE ORDER AS SHIPPED"
    )
    print("=" * 60)

    print()
    print(
        "To authorise this exact order, type:"
    )

    print()
    print(expected_confirmation)

    confirmation = input(
        "\nFinal shipping confirmation: "
    ).strip()

    if confirmation != expected_confirmation:

        print()
        print("✗ SHIPPING NOT AUTHORISED")

        print()
        print(
            "The confirmation did not exactly match:"
        )

        print(expected_confirmation)

        print()
        print(
            "NO SHIPPING API WRITE HAS BEEN MADE."
        )

        input(
            "\nPress ENTER to close..."
        )

        browser.close()
        raise SystemExit

    print()
    print("✓ SHIPPING AUTHORISATION ACCEPTED")

    # ========================================================
    # LIVE MIRAKL SHIP
    # ========================================================

    print()
    print("=" * 60)
    print("LIVE MIRAKL SHIPPING")
    print("=" * 60)

    ship_result = ship_mirakl_order(
        order_id=marketplace_order.order_id,
        carrier_code=shipping_dry_run[
            "carrier_code"
        ],
        carrier_name=shipping_dry_run[
            "carrier_label"
        ],
        tracking_number=shipping_dry_run[
            "tracking_number"
        ],
    )

    print()
    print(
        "✓ MIRAKL SHIP REQUEST COMPLETED"
    )

    print(
        f"Tracking update HTTP: "
        f"{ship_result['tracking_status_code']}"
    )

    print(
        f"Ship HTTP:            "
        f"{ship_result['ship_status_code']}"
    )

    # ========================================================
    # POST-SHIP MIRAKL VERIFICATION
    # ========================================================

    print()
    print("=" * 60)
    print("POST-SHIP MIRAKL VERIFICATION")
    print("=" * 60)

    print()
    print(
        "Reading the order back from Mirakl..."
    )

    time.sleep(1)

    verified_order = read_mirakl_order(
        marketplace_order.order_id
    )

    verified_state = verified_order.get(
        "order_state"
    )

    verified_carrier = verified_order.get(
        "shipping_carrier_code"
    )

    verified_tracking = verified_order.get(
        "shipping_tracking"
    )

    expected_tracking = (
        marketplace_order.order_id
    )

    expected_carrier = "DIR"

    print()
    print(
        f"Mirakl State:     {verified_state}"
    )

    print(
        f"Carrier Code:     {verified_carrier}"
    )

    print(
        f"Tracking Number:  {verified_tracking}"
    )

    print()
    print("-" * 60)

    verification_passed = True

    valid_post_ship_states = {
        "SHIPPED",
        "WAITING_DEBIT",
    }

    # --------------------------------------------------------
    # State verification
    # --------------------------------------------------------

    if verified_state in valid_post_ship_states:

        print(
            f"✓ Order has advanced to valid "
            f"post-shipment state: {verified_state}"
        )

    else:

        print(
            f"✗ Unexpected post-shipment state: "
            f"{verified_state}"
        )

        verification_passed = False

    # --------------------------------------------------------
    # Carrier verification
    # --------------------------------------------------------

    if verified_carrier == expected_carrier:

        print(
            "✓ Carrier code is DIR"
        )

    else:

        print(
            f"✗ Carrier code mismatch "
            f"(expected {expected_carrier})"
        )

        verification_passed = False

    # --------------------------------------------------------
    # Tracking verification
    # --------------------------------------------------------

    if verified_tracking == expected_tracking:

        print(
            "✓ Tracking number matches order number"
        )

    else:

        print(
            "✗ Tracking number mismatch "
            f"(expected {expected_tracking})"
        )

        verification_passed = False

    print("-" * 60)

    if not verification_passed:

        print()
        print("=" * 60)
        print("⚠ POST-SHIP VERIFICATION FAILED")
        print("=" * 60)

        print()
        print(
            "A Mirakl write may already have occurred."
        )

        print(
            "DO NOT automatically retry it."
        )

        print()
        print(
            "Check this order manually in Mirakl."
        )

        input(
            "\nPress ENTER to close..."
        )

        browser.close()
        raise SystemExit

    print()
    print(
        "✓ POST-SHIP VERIFICATION PASSED"
    )

    print()
    print("=" * 60)
    print("✓ MIRAKL SHIPMENT VERIFIED")
    print("=" * 60)

    print()
    print(
        f"Order {marketplace_order.order_id} "
        "is complete."
    )

    input(
        "\nPress ENTER to close..."
    )

    browser.close()
    raise SystemExit

marketplace = DIYMarketplaceAdapter(
    ORDERS_FILE
)

marketplace_orders = marketplace.read_orders()

print()
print("=" * 100)
print("DIY ORDER QUEUE")
print("=" * 100)

print(
    f"{'':>3}"
    f"{'ORDER':<15} "
    f"{'SKU':<18} "
    f"{'PRICE':>9}   "
    f"{'MIRAKL STATE':<15} "
    f"{'SHIP BY':<12} "
    f"STATUS"
)

print("-" * 100)

for index, order in enumerate(
    marketplace_orders,
    start=1
):

    try:
        mirakl_queue_order = read_mirakl_order(
            order.order_id
        )

        mirakl_state = mirakl_queue_order[
            "order_state"
        ]

        # ----------------------------------------------------
        # Shipping deadline
        # ----------------------------------------------------

        shipping_deadline_raw = (
            mirakl_queue_order.get(
                "shipping_deadline"
            )
        )

        if shipping_deadline_raw:

            try:
                shipping_deadline = (
                    datetime.fromisoformat(
                        shipping_deadline_raw.replace(
                            "Z",
                            "+00:00"
                        )
                    ).strftime(
                        "%d/%m/%Y"
                    )
                )

            except ValueError:
                shipping_deadline = "INVALID DATE"

        else:
            shipping_deadline = "-"

        # ----------------------------------------------------
        # Basic queue status
        # ----------------------------------------------------

        queue_status = get_queue_status(
            mirakl_state
        )

        # ----------------------------------------------------
        # Check for an existing customer conversation
        # ----------------------------------------------------

        if mirakl_state == "SHIPPING":

            mirakl_threads = read_mirakl_threads(
                order.order_id
            )

            queue_messages = mirakl_threads[
                "messages"
            ]

            if queue_messages:

                latest_queue_message = (
                    queue_messages[-1]
                )

                latest_topic = (
                    latest_queue_message[
                        "topic_code"
                    ]
                )

                # Topic 44 is our normal delivery message.
                # Any other conversation requires operator
                # review before treating the order as ready.
                if str(latest_topic) != "44":
                    queue_status = (
                        "⚠ CUSTOMER REVIEW"
                    )

    except Exception as error:
        mirakl_state = "ERROR"
        shipping_deadline = "-"
        queue_status = "⚠ LOOKUP FAILED"

    print(
        f"{index:>2}. "
        f"{order.order_id:<15} "
        f"{order.sku:<18} "
        f"£{order.price:>8.2f}   "
        f"{mirakl_state:<15} "
        f"{shipping_deadline:<12} "
        f"{queue_status}"
    )

print("=" * 100)

while True:

    selection = input(
        "\nSelect order number to process: "
    ).strip()

    try:
        selection_number = int(selection)

    except ValueError:

        print(
            "Please enter the number shown "
            "beside the order."
        )
        continue

    if not (
        1
        <= selection_number
        <= len(marketplace_orders)
    ):

        print(
            "That order number is not in the list."
        )
        continue

    break

marketplace_order = marketplace_orders[
    selection_number - 1
]

print()
print("=" * 60)
print("SELECTED MARKETPLACE ORDER")
print("=" * 60)

print(
    f"Marketplace:  {marketplace_order.marketplace}"
)

print(
    f"Order ID:     {marketplace_order.order_id}"
)

print(
    f"SKU:          {marketplace_order.sku}"
)

print(
    f"Price:        £{marketplace_order.price:.2f}"
)

print("=" * 60)

input(
    "\nCheck the selected order above. "
    "Press ENTER to continue..."
)


print()
print("MARKETPLACE ADAPTER")
print("=" * 60)
print(f"Marketplace: {marketplace.name}")
print(f"Order type:  {type(marketplace_order).__name__}")
print("=" * 60)

print("\n MARKETPLACE ORDER DEBUG:")
print(marketplace_order)

TEST_ORDER = marketplace_order.order_id

print()
print("=" * 60)
print("B&Q EXCEL ORDER")
print("=" * 60)
print(
    f"Order ID:      {marketplace_order.order_id}"
)
print(
    f"SKU:           {marketplace_order.sku}"
)
print(
    f"Price:         £{marketplace_order.price:.2f}"
)
print("=" * 60)

input(
    "\nCheck the Excel order above. "
    "Press ENTER to continue..."
)


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
    # CUSTOMER PANEL DIAGNOSTIC
    # ========================================================

    print()
    print("CUSTOMER PANEL DIAGNOSTIC")
    print("=" * 70)

    customer_panels = sales_frame.locator(
        "td.SALES-panel-s"
    )

    print(
        f"SALES-panel-s elements found: "
        f"{customer_panels.count()}"
    )

    for i in range(customer_panels.count()):

        panel = customer_panels.nth(i)

        print()
        print(f"PANEL {i}")
        print("TEXT:")
        print(repr(panel.inner_text()))
        print("-" * 50)

    print("=" * 70)

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
    # SAFETY CHECK 1 - ORDER NUMBER
    # ========================================================

    print()
    print("=" * 60)
    print("SAFETY CHECK 1 - ORDER NUMBER")
    print("=" * 60)

    print(
        f"Marketplace Order:   {marketplace_order.order_id}"
    )

    print(
        f"RPii Web Ref:  {web_ref}"
    )

    if web_ref == marketplace_order.order_id:

        print()
        print("✓ ORDER NUMBER MATCH")

    else:

        print()
        print("✗ ORDER NUMBER MISMATCH")
        print()
        print(
            "STOPPING - no further processing will take place."
        )

        input(
            "\nPress ENTER to close..."
        )

        browser.close()

        raise SystemExit

    

    # ========================================================
    # READ RPii CUSTOMER ACCOUNT
    # ========================================================

    customer_panel = None

    for i in range(customer_panels.count()):
        panel = customer_panels.nth(i)
        text = panel.inner_text()

        if "T:" in text:
            customer_panel = panel
            break

    if customer_panel is None:
        raise RuntimeError(
            "Could not find RPii delivery customer panel."
        )

    customer_text = customer_panel.inner_text()

    customer_name = extract_customer_name(
        customer_text
    )

    # ========================================================
    # READ DELIVERY CUSTOMER DETAILS
    # ========================================================

    delivery_customer = extract_delivery_customer(
        customer_text
    )

    delivery_postcode = extract_uk_postcode(
       delivery_customer["address"]
    )

    # ========================================================
    # SAFETY CHECK 4 - POSTCODE
    # ========================================================

    print()
    print("=" * 60)
    print("SAFETY CHECK 4 - POSTCODE")
    print("=" * 60)

    excel_postcode = (
        marketplace_order.postcode
        .replace(" ", "")
        .upper()
    )

    rp2_postcode = (
        delivery_postcode
        .replace(" ", "")
        .upper()
        if delivery_postcode
        else None
    )

    print(
        f"Marketplace Postcode: {marketplace_order.postcode}"
    )

    print(
        f"RPii Postcode:  {delivery_postcode}"
    )

    if rp2_postcode == excel_postcode:

        print()
        print("✓ POSTCODE MATCH")

    else:

        print()
        print("✗ POSTCODE MISMATCH")
        print()
        print(
            "STOPPING - no further processing will take place."
        )

        input(
            "\nPress ENTER to close..."
        )

        browser.close()

        raise SystemExit
    
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
    # SAFETY CHECK 2 - SKU
    # ========================================================
    
    print()
    print("=" * 60)
    print("SAFETY CHECK 2 - SKU")
    print("=" * 60)

    excel_sku = str(
        marketplace_order.sku
    ).strip().upper()

    rp2_sku = str(
        sku
    ).strip().upper()

    print(
        f"Marketplace SKU:     {excel_sku}"
    )

    print(
        f"RPii SKU:      {rp2_sku}"
    )

    if excel_sku == rp2_sku:

        print()
        print("✓ SKU MATCH")

        rp2_sku_mismatch = False

    else:

        print()
        print("⚠ RPii SKU MISMATCH")
        print()
        print(
            "RPii does not exactly match the "
            "marketplace SKU."
        )
        print(
            "Processing will continue to the "
            "read-only Mirakl validation."
        )
        print()
        print(
            "NO SKU OVERRIDE HAS BEEN AUTHORISED."
        )

        rp2_sku_mismatch = True

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
    # READ PAYMENT SUMMARY
    # ========================================================

    payment_summary = read_payment_summary(
        sales_frame
    )

    # ========================================================
    # SAFETY CHECK 3 - ORDER VALUE
    # ========================================================

    print()
    print("=" * 60)
    print("SAFETY CHECK 3 - ORDER VALUE")
    print("=" * 60)

    excel_amount = round(
        float(marketplace_order.price),
        2
    )

    rp2_amount = round(
        float(total_sale),
        2
    )

    print(
        f"Marketplace Amount:  £{excel_amount:.2f}"
    )

    print(
        f"RPii Total:    £{rp2_amount:.2f}"
    )

    if excel_amount == rp2_amount:

        print()
        print("✓ ORDER VALUE MATCH")

    else:

        print()
        print("✗ ORDER VALUE MISMATCH")
        print()
        print(
            "STOPPING - no further processing will take place."
        )

        input(
            "\nPress ENTER to close..."
        )

        browser.close()

        raise SystemExit

    # ========================================================
    # READ COURIER / DELIVERY INFORMATION
    # ========================================================

    delivery_cells = sales_frame.locator(
        "td.tcb"
    )

    print()
    print("=" * 70)
    print("DELIVERY CELL DIAGNOSTIC")
    print("=" * 70)

    for i in range(delivery_cells.count()):

        cell_text = (
            delivery_cells.nth(i)
            .inner_text()
            .strip()
        )

        if cell_text:

            print(
                f"CELL {i}: {repr(cell_text)}"
            )

    print("=" * 70)

    delivery_text = None

    for i in range(
        delivery_cells.count()
    ):
        text = delivery_cells.nth(
            i
        ).inner_text()

        if (
            "DELIVER" in text
            or "ON MANFST" in text
            or "DELVRD" in text
        ):
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
    # Payment reporting
    # --------------------------------------------------------
    
    print()
    print("PAYMENT SUMMARY")
    print("-" * 60)

    if payment_summary["processed"]:

        print("Status:        PAYMENT PRESENT")

        for detail in payment_summary["details"]:
            print(
                f"Details:       {detail}"
            )

    else:

        print("Status:        PAYMENT REQUIRED")
        print("Details:       None")

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

    # ========================================================
    # SAFETY CHECK 5 - TELEPHONE
    # ========================================================

    print()
    print("=" * 60)
    print("SAFETY CHECK 5 - TELEPHONE")
    print("=" * 60)

    excel_phone_1 = normalise_phone(
        marketplace_order.phone_1
    )

    excel_phone_2 = normalise_phone(
        marketplace_order.phone_2
    )

    rp2_phone = normalise_phone(
        delivery_customer["telephone"]
    )

    print(
        f"Marketplace Phone 1: {marketplace_order.phone_1}"
    )

    print(
        f"Marketplace Phone 2: {marketplace_order.phone_2}"
    )

    print(
        f"RPii Phone:    {delivery_customer['telephone']}"
    )

    valid_excel_phones = {
        phone
        for phone in (
            excel_phone_1,
            excel_phone_2,
        )
        if phone
    }

    if rp2_phone in valid_excel_phones:

        print()
        print("✓ TELEPHONE MATCH")

    else:

        print()
        print("✗ TELEPHONE MISMATCH")
        print()
        print(
            "STOPPING - no further processing will take place."
        )

        input(
            "\nPress ENTER to close..."
        )

        browser.close()

        raise SystemExit

    # ========================================================
    # READ MIRAKL ORDER
    # ========================================================

    print()
    print("=" * 60)
    print("MIRAKL READ-ONLY VALIDATION")
    print("=" * 60)

    mirakl_order = read_mirakl_order(
        marketplace_order.order_id
    )

    shipping_recipient_name = (
        mirakl_order.get("customer_name")
        or delivery_customer["name"]
    )

    print(
        f"Mirakl Order:  "
        f"{mirakl_order['order_id']}"
    )

    print(
        f"Customer:      "
        f"{mirakl_order['customer_name']}"
    )

    print(
        f"SKU:           "
        f"{mirakl_order['sku']}"
    )

    print(
        f"Price:         "
        f"£{mirakl_order['price']:.2f}"
    )

    print(
        f"Postcode:      "
        f"{mirakl_order['postcode']}"
    )

    print(
        f"Telephone:     "
        f"{mirakl_order['phone_1']}"
    )

    # ========================================================
    # SAFETY GATE - EXCEL / RPii / MIRAKL
    # ========================================================

    print()
    print("=" * 60)
    print("THREE-WAY SAFETY GATE")
    print("=" * 60)

    # --------------------------------------------------------
    # Build RPii data for validation
    # --------------------------------------------------------

    rp2_validation_data = {
        "order_id": web_ref,
        "sku": sku,
        "price": total_sale,
        "postcode": delivery_postcode,
        "telephone": delivery_customer["telephone"],
    }

    # --------------------------------------------------------
    # Run three-way validation
    # --------------------------------------------------------

    safety_result = build_three_way_safety_result(
        marketplace_order,
        rp2_validation_data,
        mirakl_order,
    )

    safety_failures = safety_result["failures"]
    safety_values = safety_result["values"]

    # --------------------------------------------------------
    # DISPLAY RESULTS
    # --------------------------------------------------------

    if "Order number" in safety_failures:
        print("✗ Order number mismatch")
    else:
        print("✓ Order number matches all 3 systems")

    if "SKU" in safety_failures:
        print("✗ SKU mismatch")
    else:
        print("✓ SKU matches all 3 systems")

    if "Price" in safety_failures:
        print("✗ Price mismatch")
    else:
        print("✓ Price matches all 3 systems")

    if "Postcode" in safety_failures:
        print("✗ Postcode mismatch")
    else:
        print("✓ Postcode matches all 3 systems")

    if "Telephone" in safety_failures:
        print("✗ Telephone mismatch")
    else:
        print("✓ Telephone matches all 3 systems")

    print("=" * 60)

    # ========================================================
    # CONTROLLED MANUAL OVERRIDE - SKU ONLY
    # ========================================================

    override_used = False
    override_reason = None

    excel_sku = safety_values[
        "excel_sku"
    ]

    rp2_sku = safety_values[
        "rp2_sku"
    ]

    mirakl_sku = safety_values[
        "mirakl_sku"
    ]

    # --------------------------------------------------------
    # Override is available ONLY when:
    #
    # - SKU is the only failed safety check
    # - Marketplace and Mirakl agree exactly
    # - RPii contains a different SKU representation
    # --------------------------------------------------------

    rp2_only_sku_difference = (
        set(safety_failures) == {"SKU"}
        and excel_sku == mirakl_sku
        and rp2_sku != excel_sku
    )

    if rp2_only_sku_difference:

        print()
        print("=" * 60)
        print("RPii SKU DIFFERENCE - MANUAL OVERRIDE AVAILABLE")
        print("=" * 60)

        print(
            f"Marketplace SKU: {excel_sku}"
        )

        print(
            f"RPii SKU:        {rp2_sku}"
        )

        print(
            f"Mirakl SKU:      {mirakl_sku}"
        )

        print()
        print(
            "✓ Marketplace and Mirakl agree."
        )

        print(
            "⚠ RPii contains a different SKU representation."
        )

        expected_override = (
            f"OVERRIDE SKU {marketplace_order.order_id}"
        )

        print()
        print(
            "To approve this RPii SKU difference "
            "for this exact order, type:"
        )

        print()
        print(expected_override)

        confirmation = input(
            "\nSKU override confirmation: "
        ).strip()

        if confirmation == expected_override:

            override_used = True
            override_reason = "RPii SKU difference"

            # Deliberately remove ONLY the SKU failure.
            safety_failures.remove(
                "SKU"
            )

            print()
            print(
                "⚠ SKU OVERRIDE ACCEPTED"
            )

            print()
            print(
                "Marketplace and Mirakl SKU will be "
                "treated as authoritative for this order."
            )

            print(
                f"Approved SKU: {excel_sku}"
            )

            print(
                f"RPii SKU:     {rp2_sku}"
            )

        else:

            print()
            print(
                "✗ SKU OVERRIDE NOT AUTHORISED"
            )

            print()
            print(
                "The confirmation did not exactly match:"
            )

            print(expected_override)

    
    # ========================================================
    # HARD STOP
    # ========================================================

    if safety_failures:

        print()
        print(
            "✗ THREE-WAY SAFETY GATE FAILED"
        )

        print()

        print(
            "Problems found: "
            + ", ".join(
                safety_failures
            )
        )

        print()
        print(
            "STOPPING - no invoice/message "
            "processing will continue."
        )

        input(
            "\nPress ENTER to close..."
        )

        browser.close()
        raise SystemExit

    # ========================================================
    # SAFETY GATE RESULT
    # ========================================================

    print()

    if override_used:

        print(
            "⚠ THREE-WAY SAFETY GATE PASSED "
            "WITH MANUAL SKU OVERRIDE"
        )

        print()

        print(
            "Marketplace SKU: "
            f"{safety_values['excel_sku']}"
        )

        print(
            "Mirakl SKU:      "
            f"{safety_values['mirakl_sku']}"
        )

        print(
            "RPii SKU:        "
            f"{safety_values['rp2_sku']}"
        )

        print()

        print(
            "✓ Marketplace and Mirakl agree."
        )

        print(
            "⚠ RPii SKU difference was manually approved."
        )

        print()

        print(
            "All other safety checks agree."
        )

    else:

        print(
            "✓ THREE-WAY SAFETY GATE PASSED"
        )

        print()

        print(
            "Marketplace, RPii and Mirakl agree."
        )

    # ========================================================
    # CUSTOMER-FACING SKU
    # ========================================================

    # Customer communications should use the marketplace SKU.
    #
    # This is especially important where RPii uses an internal
    # or extended stock code, e.g.:
    #
    # Marketplace / Mirakl: EY48SB8
    # RPii:                 EY48SB8-80(EY48SB8-80)

    customer_sku = str(
        marketplace_order.sku
    ).strip()

    customer_message_already_sent = False

    # ========================================================
    # MIRAKL CUSTOMER COMMUNICATION SAFETY CHECK
    # ========================================================

    print()
    print("=" * 60)
    print("MIRAKL COMMUNICATION SAFETY CHECK")
    print("=" * 60)

    mirakl_threads = read_mirakl_threads(
        marketplace_order.order_id
    )

    existing_messages = mirakl_threads[
        "messages"
    ]

    if not existing_messages:

        print("No existing customer messages found.")
        print()
        print("✓ CLEAR TO CONTINUE")

    else:

        first_message = existing_messages[0]
        latest_message = existing_messages[-1]

        print(
            f"Existing messages: {len(existing_messages)}"
        )

        print(
            f"Conversation started by: "
            f"{first_message['sender_type']}"
        )

        print(
            f"Latest message from:      "
            f"{latest_message['sender_type']}"
        )

        print(
            f"Latest sender:            "
            f"{latest_message['sender_name']}"
        )

        print(
            f"Latest message date:      "
            f"{latest_message['date']}"
        )

        print(
            f"Topic code:               "
            f"{latest_message['topic_code']}"
        )

        print()
        print("LATEST MESSAGE")
        print("-" * 60)

        print(
            latest_message["body"]
        )

        print("-" * 60)

        # ----------------------------------------------------
        # CUSTOMER MESSAGE = HARD PAUSE
        # ----------------------------------------------------

        if latest_message["sender_type"] in (
            "CUSTOMER",
            "CUSTOMER_USER",
        ):

            print()
            print(
                "⚠ CUSTOMER MESSAGE REQUIRES REVIEW"
            )

            print()
            print(
                "The customer is currently the latest "
                "person to have messaged."
            )

            print(
                "Automatic invoice/message processing "
                "has been paused."
            )

            print()
            print(
                "No customer message has been sent."
            )

            input(
                "\nReview the customer's message. "
                "Press ENTER to close..."
            )

            browser.close()
            raise SystemExit

        elif latest_message["sender_type"] in (
            "SHOP",
            "SHOP_USER",
        ):

            print()

            if (
                str(latest_message["topic_code"])
                == str(DIY_DELIVERY_TOPIC_CODE)
            ):

                print(
                    "⚠ EXISTING DELIVERY MESSAGE FOUND"
                )

                print()
                print(
                    "The latest Mirakl message was sent "
                    "by the shop using delivery topic 44."
                )

                print()
                print(
                    "✓ DUPLICATE MESSAGE PROTECTION ACTIVE"
                )

                print(
                    "The customer-message stage will be "
                    "skipped for this order."
                )

                print()
                print(
                    f"Latest sender: "
                    f"{latest_message['sender_name']}"
                )

                print(
                    f"Latest date:   "
                    f"{latest_message['date']}"
                )

                customer_message_already_sent = True

            else:

                print(
                    "✓ Latest message was sent by the shop."
                )

                print(
                    "Existing conversation uses a different "
                    "topic - continuing normal route."
                )

        else:

            print()
            print(
                "⚠ UNKNOWN MIRAKL SENDER TYPE"
            )

            print(
                f"Sender type: "
                f"{latest_message['sender_type']}"
            )

            print()
            print(
                "STOPPING - sender could not be "
                "safely classified."
            )

            print(
                "No customer message has been sent."
            )

            input(
                "\nPress ENTER to close..."
            )

            browser.close()
            raise SystemExit

    # ========================================================
    # DETERMINE ORDER ROUTE
    # ========================================================

    if payment_summary["processed"]:
        order_route = "READY FOR INVOICE"
    else:
        order_route = "WAITING FOR PAYMENT"

    # ========================================================
    # PAYMENT ROUTE SAFETY STOP
    # ========================================================

    if order_route == "WAITING FOR PAYMENT":
        print()
        print("ORDER ROUTE")
        print("=" * 60)
        print("Order is waiting for payment.")
        print()
        print(
            "STOPPING - no invoice has been generated "
            "and no customer message has been prepared."
        )
        print()
        print("NOTHING HAS BEEN SENT.")

        input(
            "\nPress ENTER to close..."
        )

        browser.close()
        raise SystemExit

    # ========================================================
    # EXISTING DELIVERY MESSAGE - CONTINUE TO SHIPPING
    # ========================================================

    if customer_message_already_sent:

        print()
        print("=" * 60)
        print("CUSTOMER MESSAGE ALREADY COMPLETE")
        print("=" * 60)

        print(
            f"Order:       "
            f"{marketplace_order.order_id}"
        )

        print(
            "Message:     Existing delivery "
            "message found in Mirakl"
        )

        print()
        print(
            "No invoice/customer message will "
            "be generated or sent again."
        )

        complete_mirakl_shipping(
            browser=browser,
            marketplace_order=marketplace_order,
            mirakl_order=mirakl_order,
        )

    # ========================================================
    # NORMAL ROUTE - OPEN A4 INVOICE
    # ========================================================

    if order_route == "READY FOR INVOICE":

        print()
        print("NORMAL ROUTE")
        print("=" * 60)
        print("Order is ready for invoice.")

        # ----------------------------------------------------
        # Find RPii Print button
        # ----------------------------------------------------

        print_button = sales_frame.locator(
            'td.tcbw[onclick^="doPrint("]'
        )

        if print_button.count() != 1:

            print(
                f"✗ Expected 1 Print button, "
                f"found {print_button.count()}."
            )

            input(
                "\nPress ENTER to close..."
            )

            browser.close()
            raise SystemExit

        # ----------------------------------------------------
        # Open Print options
        # ----------------------------------------------------

        print("Opening RPii Print panel...")

        print_button.evaluate(
            "element => element.click()"
        )

        page.wait_for_timeout(1000)

        print("✓ Print panel opened.")

        # ====================================================
        # RPii SMS CONTROLS - READ ONLY
        # ====================================================

        sms_mobile = sales_frame.locator(
            "#pr_cus_mobile"
        )

        sms_text = sales_frame.locator(
            "#del_cus_text"
        )

        sms_send_button = sales_frame.locator(
            'img[onclick^="sendText("]'
        )

        print()
        print("=" * 60)
        print("RPii SMS CONTROLS - READ ONLY")
        print("=" * 60)

        print(
            f"Mobile field found: "
            f"{sms_mobile.count() == 1}"
        )

        print(
            f"Message field found: "
            f"{sms_text.count() == 1}"
        )

        print(
            f"SMS button found:    "
            f"{sms_send_button.count() == 1}"
        )

        print()
        print("NO SMS HAS BEEN SENT.")

        # ----------------------------------------------------
        # END SMS BLOCK
        # ----------------------------------------------------
        


        # ----------------------------------------------------
        # Find A4 Invoice
        # ----------------------------------------------------

        a4_invoice = sales_frame.locator(
            'tr:has-text("A4 Invoice") '
            'img[onclick^="doInvoiceP("]'
        )

        if a4_invoice.count() != 1:

            print(
                f"✗ Expected 1 A4 Invoice option, "
                f"found {a4_invoice.count()}."
            )

            input(
                "\nPress ENTER to close..."
            )

            browser.close()
            raise SystemExit

        # Show us the RPii action for diagnostic purposes
        invoice_action = a4_invoice.get_attribute(
            "onclick"
        )

        print(
            f"A4 Invoice action: {invoice_action}"
        )

        # ----------------------------------------------------
        # Open A4 Invoice
        # ----------------------------------------------------

        print("Opening A4 Invoice...")

        a4_invoice.evaluate(
            "element => element.click()"
        )

        page.wait_for_timeout(1500)

        # ========================================================
        # SAVE GENERATED INVOICE PDF
        # ========================================================

        print()
        print("SAVING INVOICE")
        print("=" * 60)

        # Find RPii's PRINTFrame
        print_frame = None

        for frame in page.frames:
            if frame.name == "PRINTFrame":
                print_frame = frame
                break

        if print_frame is None:
            print("✗ PRINTFrame could not be found.")

            input(
                "\nPress ENTER to close..."
            )

            browser.close()
            raise SystemExit

        invoice_url = print_frame.url

        print(
            f"Invoice URL: {invoice_url}"
        )

        # Safety check - make sure this really is a PDF
        if not invoice_url.lower().endswith(".pdf"):
            print(
                "✗ PRINTFrame does not contain a PDF."
            )

            input(
                "\nPress ENTER to close..."
            )

            browser.close()
            raise SystemExit

        # Make sure invoice folder exists
        os.makedirs(
            INVOICE_FOLDER,
            exist_ok=True
        )

        # Create invoice filename
        customer_surname = get_customer_surname(
            delivery_customer["name"]
        )

        invoice_filename = (
            f"{marketplace_order.order_id}-"
            f"{customer_surname}-invoice.pdf"
        )

        invoice_path = os.path.join(
            INVOICE_FOLDER,
            invoice_filename
        )

        print(
            f"Saving as: {invoice_filename}"
        )

        # Download using the authenticated
        # Playwright browser context
        response = context.request.get(
            invoice_url
        )

        if not response.ok:

            print(
                f"✗ Invoice download failed. "
                f"HTTP status: {response.status}"
            )

            input(
                "\nPress ENTER to close..."
            )

            browser.close()
            raise SystemExit

        # Save PDF
        with open(
            invoice_path,
            "wb"
        ) as file:

            file.write(
                response.body()
            )

        print()
        print("✓ INVOICE SAVED SUCCESSFULLY")
        print()
        print(
            f"Saved to: {invoice_path}"
        )

        # input(
        #     "\nCheck the saved PDF file. "
        #     "Press ENTER to close..."
        # )

        # browser.close()
        # raise SystemExit

        # ====================================================
        # RETURN FROM INVOICE TO EPoS
        # ====================================================

        print()
        print("Returning from invoice to EPoS...")

        epos_tab = page.locator(
            "#SALEStd"
        )

        if epos_tab.count() != 1:

            print()
            print(
                "✗ Could not uniquely identify the EPoS tab."
            )

            input(
                "\nPress ENTER to close..."
            )

            browser.close()
            raise SystemExit

        epos_tab.evaluate(
            "element => element.click()"
        )

        page.wait_for_timeout(1000)

        # ----------------------------------------------------
        # Re-acquire SALESFrame.
        # EPoS reloads the Sales module.
        # ----------------------------------------------------

        print(
            "Re-acquiring RPii SALESFrame..."
        )

        sales_frame = None

        for frame in page.frames:

            if frame.name == "SALESFrame":
                sales_frame = frame
                break

        if sales_frame is None:

            print()
            print(
                "✗ SALESFrame could not be found "
                "after returning to EPoS."
            )

            input(
                "\nPress ENTER to close..."
            )

            browser.close()
            raise SystemExit

        print("✓ SALESFrame re-acquired.")

        # ----------------------------------------------------
        # Confirm that the current sale has returned.
        # ----------------------------------------------------

        current_sale_total = sales_frame.locator(
            "#qh_ordval"
        )

        current_sale_total.wait_for(
            state="visible",
            timeout=10000,
        )

        print("✓ Returned to EPoS sale.")

       
    # ========================================================
    # BUILD CUSTOMER MESSAGE PREVIEW
    # ========================================================

    print()
    print("CUSTOMER MESSAGE PREVIEW")
    print("=" * 60)

    courier = delivery["courier"]

    # --------------------------------------------------------
    # Customer-facing delivery details
    # --------------------------------------------------------

    customer_delivery_details = dict(
        delivery_customer
    )

    customer_delivery_details["name"] = (
        shipping_recipient_name
    )

    if courier == "SK":

        customer_date_text = (
            delivery["customer_date"]
            .strftime("%d/%m/%Y")
        )

        customer_message = build_sgk_message(
            customer_delivery_details,
            customer_sku,
            description,
            customer_date_text,
        )

    elif courier == "ED":

        customer_date_text = (
            delivery["customer_date"]
            .strftime("%d/%m/%Y")
        )

        customer_message = build_ed_message(
            customer_delivery_details,
            customer_sku,
            description,
            customer_date_text,
        )

    else:

        print(
            "✗ Cannot build message because "
            "courier is not recognised."
        )

        input(
            "\nPress ENTER to close..."
        )

        browser.close()
        raise SystemExit

    # ========================================================
    # BUILD CUSTOMER SMS PREVIEW
    # ========================================================

    if courier == "SK":

        customer_sms = build_sgk_sms()

    elif courier == "ED":

        customer_sms = build_ed_sms(
            sku=customer_sku,
            delivery_date=customer_date_text,
        )

    else:

        raise RuntimeError(
            f"Cannot build SMS for "
            f"unknown courier: {courier}"
        )

    print()
    print("=" * 60)
    print("CUSTOMER SMS PREVIEW")
    print("=" * 60)

    print(
        f"Order:       "
        f"{marketplace_order.order_id}"
    )

    print(
        f"Mobile:      "
        f"{delivery_customer['telephone']}"
    )

    print(
        f"Courier:     {courier}"
    )

    print()
    print("-" * 60)
    print(customer_sms)
    print("-" * 60)

    print()
    print(
        "SMS PREVIEW ONLY - "
        "NO SMS HAS BEEN SENT."
    )

    print()
    print(
        f"Order:       {marketplace_order.order_id}"
    )

    print(
        f"Customer:    {shipping_recipient_name}"
    )

    print(
        f"Courier:     {courier}"
    )

    print(
        f"Attachment:  {invoice_filename}"
    )

    print(
        "Mirakl Topic: Information about delivery "
        "(incl. tracking)"
    )

    print(
        "Topic Code:   44"
    )

    print()
    print("-" * 60)
    print()
    print(customer_message)
    print()

    # ========================================================
    # POPULATE RPii SMS FIELDS - DO NOT SEND
    # ========================================================

    print()
    print("Preparing RPii SMS fields...")

    # --------------------------------------------------------
    # Re-acquire the current RPii Print button.
    # Returning from the invoice can refresh the Sales module.
    # --------------------------------------------------------

    current_print_button = sales_frame.locator(
        'td.tcbw[onclick^="doPrint("]'
    )

    if current_print_button.count() != 1:

        print()
        print(
            f"✗ Expected 1 current Print button, "
            f"found {current_print_button.count()}."
        )

        input(
            "\nPress ENTER to close..."
        )

        browser.close()
        raise SystemExit

    # --------------------------------------------------------
    # Re-acquire the SMS controls from the current Sales frame.
    # --------------------------------------------------------

    sms_mobile = sales_frame.locator(
        "#pr_cus_mobile"
    )

    sms_text = sales_frame.locator(
        "#del_cus_text"
    )

    sms_send_button = sales_frame.locator(
        'img[onclick^="sendText("]'
    )

    if (
        not sms_mobile.is_visible()
        or not sms_text.is_visible()
    ):

        print(
            "SMS controls are not currently visible. "
            "Opening the current RPii Print panel..."
        )

        current_print_button.evaluate(
            "element => element.click()"
        )

        page.wait_for_timeout(1000)

    # --------------------------------------------------------
    # Safety check after opening Print panel.
    # --------------------------------------------------------

    print()
    print("SMS CONTROL STATE AFTER OPENING PRINT")
    print("-" * 60)

    print(
        f"Mobile field count:   {sms_mobile.count()}"
    )

    print(
        f"Mobile field visible: {sms_mobile.is_visible()}"
    )

    print(
        f"Message field count:  {sms_text.count()}"
    )

    print(
        f"Message visible:      {sms_text.is_visible()}"
    )

    print(
        f"Send button count:    {sms_send_button.count()}"
    )

    print(
        f"Send button visible:  {sms_send_button.is_visible()}"
    )

    print("-" * 60)

    if (
        not sms_mobile.is_visible()
        or not sms_text.is_visible()
        or not sms_send_button.is_visible()
    ):

        print()
        print(
            "✗ RPii SMS controls did not become visible."
        )

        print()
        print(
            "NO SMS HAS BEEN SENT."
        )

        input(
            "\nPress ENTER to close..."
        )

        browser.close()
        raise SystemExit

    # --------------------------------------------------------
    # Require the fields to be visible before entering data.
    # --------------------------------------------------------

    sms_mobile.wait_for(
        state="visible",
        timeout=5000,
    )

    sms_text.wait_for(
        state="visible",
        timeout=5000,
    )

    sms_send_button.wait_for(
        state="visible",
        timeout=5000,
    )

    # --------------------------------------------------------
    # Select SMS target number
    # --------------------------------------------------------

    if SMS_TEST_MODE:
        sms_target_mobile = SMS_TEST_MOBILE
    else:
        sms_target_mobile = delivery_customer["telephone"]

    # --------------------------------------------------------
    # Populate fields only.
    # NO SMS button click occurs here.
    # --------------------------------------------------------

    sms_mobile.fill(
        sms_target_mobile
    )

    sms_text.fill(
        customer_sms
    )

    entered_mobile = sms_mobile.input_value()
    entered_sms = sms_text.input_value()

    print()
    print("=" * 60)
    print("RPii SMS FIELD VERIFICATION")
    print("=" * 60)

    print(
        f"Expected mobile: {sms_target_mobile}"
    )

    print(
        f"RPii mobile:     {entered_mobile}"
    )

    print()

    print("Expected SMS:")
    print("-" * 60)
    print(customer_sms)

    print()
    print("RPii SMS field:")
    print("-" * 60)
    print(entered_sms)

    print()

    if (
        entered_mobile
        == sms_target_mobile
        and entered_sms
        == customer_sms
    ):

        print(
            "✓ RPii SMS fields match the expected values"
        )

    else:

        print(
            "✗ RPii SMS field verification failed"
        )

        input(
            "\nPress ENTER to close..."
        )

        browser.close()
        raise SystemExit

    print()
    print(
        "SMS FIELDS POPULATED ONLY - "
        "THE SMS SEND BUTTON HAS NOT BEEN CLICKED."
    )

    # ========================================================
    # FINAL RPii SMS SEND GATE
    # ========================================================

    print()
    print("=" * 60)
    print("FINAL RPii SMS CONFIRMATION")
    print("=" * 60)

    if SMS_TEST_MODE:
        print()
        print("⚠ SMS TEST MODE ACTIVE")
        print("Customer mobile will NOT be used.")
        print(
            f"Test mobile: {sms_target_mobile}"
        )

    print()
    print(
        f"Order:       {marketplace_order.order_id}"
    )

    print(
        f"Mobile:      {sms_target_mobile}"
    )

    print()
    print("SMS:")
    print("-" * 60)
    print(customer_sms)
    print("-" * 60)

    expected_sms_confirmation = (
        f"TEXT {marketplace_order.order_id}"
    )

    print()
    print(
        "To authorise this exact SMS, type:"
    )

    print()
    print(expected_sms_confirmation)
    print()

    sms_confirmation = input(
        "Final SMS confirmation: "
    ).strip()

    if sms_confirmation != expected_sms_confirmation:

        print()
        print("✗ SMS NOT AUTHORISED")

        print()
        print(
            "The confirmation did not exactly match:"
        )

        print(expected_sms_confirmation)

        print()
        print("NO SMS HAS BEEN SENT.")

        input(
            "\nPress ENTER to close..."
        )

        browser.close()
        raise SystemExit

    print()
    print("✓ SMS AUTHORISATION ACCEPTED")

    # ========================================================
    # LIVE RPii SMS SEND
    # ========================================================

    print()
    print("=" * 60)
    print("LIVE RPii SMS SEND")
    print("=" * 60)

    expected_alert_text = (
        f"SMS sent to {sms_target_mobile}"
    )

    print()
    print(
        f"Sending SMS to {sms_target_mobile}..."
    )

    # --------------------------------------------------------
    # IMPORTANT: Click the SMS button ONCE ONLY
    # --------------------------------------------------------

    sms_send_button.evaluate(
        "element => element.click()"
    )

    # --------------------------------------------------------
    # Verify RPii acknowledgement.
    #
    # RPii creates the acknowledgement dynamically inside
    # .alertDiv. Search every frame because the alert is
    # generated inside SALESFrame.
    # --------------------------------------------------------

    print()
    print(
        "Waiting for RPii SMS acknowledgement..."
    )

    acknowledgement_found = False
    acknowledgement_text = ""
    acknowledgement_frame = ""

    acknowledgement_found = False
    acknowledgement_text = ""
    acknowledgement_frame = ""

    # --------------------------------------------------------
    # RPii creates the acknowledgement dynamically.
    # Search all frames for the actual alertDiv text.
    # --------------------------------------------------------

    for attempt in range(20):

        for frame in page.frames:

            try:
                alert_divs = frame.locator(
                    ".alertDiv"
                )

                for i in range(
                    alert_divs.count()
                ):

                    text = (
                        alert_divs.nth(i)
                        .text_content()
                        or ""
                    ).strip()

                    if expected_alert_text in text:

                        acknowledgement_found = True
                        acknowledgement_text = text
                        acknowledgement_frame = (
                            frame.name
                            or "(unnamed)"
                        )

                        break

                if acknowledgement_found:
                    break

            except Exception:
                continue

        if acknowledgement_found:
            break

        page.wait_for_timeout(500)

    print()

    if acknowledgement_found:

        print(
            f"RPii acknowledgement frame: "
            f"{acknowledgement_frame}"
        )

        print(
            f"RPii acknowledgement: "
            f"{acknowledgement_text}"
        )

        print()
        print(
            "✓ RPii SMS SEND ACKNOWLEDGED"
        )

        print(
            f"✓ RPii reported SMS sent to "
            f"{sms_target_mobile}"
        )

    else:

        print(
            "✗ RPii SMS VERIFICATION FAILED"
        )

        print()
        print(
            f"Expected acknowledgement: "
            f"{expected_alert_text}"
        )

        print()
        print(
            "The SMS send action has already "
            "been attempted."
        )

        print(
            "DO NOT automatically retry it."
        )

        input(
            "\nPress ENTER to close..."
        )

        browser.close()
        raise SystemExit

   

    
    # ========================================================
    # MIRAKL SEND PAYLOAD - DRY RUN ONLY
    # ========================================================

    dry_run_payload = build_mirakl_dry_run_payload(
        order_id=marketplace_order.order_id,
        customer_name=shipping_recipient_name,
        message_body=customer_message,
        invoice_path=invoice_path,
        existing_threads=mirakl_threads["threads"],
    )

    print()
    print("=" * 60)
    print("MIRAKL SEND PAYLOAD - DRY RUN")
    print("=" * 60)

    print(
        f"Mode:          "
        f"{dry_run_payload['mode']}"
    )

    print(
        f"Order:         "
        f"{dry_run_payload['order_id']}"
    )

    print(
        f"Customer:      "
        f"{dry_run_payload['customer']}"
    )

    print(
        f"Topic:         "
        f"{dry_run_payload['topic']['label']}"
    )

    print(
        f"Topic Code:    "
        f"{dry_run_payload['topic']['value']}"
    )

    print(
        f"Existing Thread: "
        f"{dry_run_payload['existing_thread_id']}"
    )

    print(
        f"Attachment:    "
        f"{dry_run_payload['attachment']['filename']}"
    )

    print(
        f"File exists:   "
        f"{dry_run_payload['attachment']['exists']}"
    )

    print()
    print("MESSAGE BODY")
    print("-" * 60)

    print(
        dry_run_payload["message_body"]
    )

    print("-" * 60)

    
    # --------------------------------------------------------
    # DRY-RUN SAFETY VALIDATION
    # --------------------------------------------------------

    dry_run_failures = []

    if dry_run_payload["mode"] != "DRY_RUN":
        dry_run_failures.append(
            "Payload is not marked DRY_RUN"
        )

    if (
        dry_run_payload["order_id"]
        != marketplace_order.order_id
    ):
        dry_run_failures.append(
            "Payload order number mismatch"
        )

    if (
        dry_run_payload["topic"]["value"]
        != DIY_DELIVERY_TOPIC_CODE
    ):
        dry_run_failures.append(
            "Incorrect Mirakl topic code"
        )

    if not dry_run_payload[
        "message_body"
    ].strip():
        dry_run_failures.append(
            "Customer message is blank"
        )

    if not dry_run_payload[
        "attachment"
    ]["exists"]:
        dry_run_failures.append(
            "Invoice attachment does not exist"
        )

    print()
    print("=" * 60)
    print("DRY-RUN PAYLOAD SAFETY CHECK")
    print("=" * 60)

    if dry_run_failures:

        for failure in dry_run_failures:
            print(
                f"✗ {failure}"
            )

        print()
        print(
            "✗ DRY-RUN PAYLOAD FAILED"
        )

        print(
            "A Mirakl send operation would "
            "NOT be permitted."
        )

        print()
        print("STOPPING - live send gate will not be reached.")

        input(
            "\nPress ENTER to close..."
        )

        browser.close()
        raise SystemExit

    else:

        print(
            "✓ Order number present and correct"
        )
        print(
            "✓ Delivery topic is 44"
        )
        print(
            "✓ Message body is present"
        )
        print(
            "✓ Invoice attachment exists"
        )

        print()
        print(
            "✓ DRY-RUN PAYLOAD PASSED"
        )

    print()
    print(
        "DRY RUN ONLY - NO API WRITE REQUEST "
        "HAS BEEN MADE."
    )

    print("-" * 60)

    print()
    print(
        "NO MIRAKL MESSAGE HAS BEEN SENT."
    )

     # ========================================================
    # FINAL CUSTOMER COMMUNICATION REVIEW
    # ========================================================

    if existing_messages:

        print()
        print("=" * 60)
        print("LATEST MIRAKL MESSAGE - REVIEW BEFORE SENDING")
        print("=" * 60)

        print(
            f"Latest sender: "
            f"{latest_message['sender_name']}"
        )

        print(
            f"Latest date:   "
            f"{latest_message['date']}"
        )

        print(
            f"Topic code:    "
            f"{latest_message['topic_code']}"
        )

        print()
        print("LATEST MESSAGE")
        print("-" * 60)

        print(
            latest_message["body"]
        )

        print("-" * 60)

    else:

        print()
        print("=" * 60)
        print("LATEST MIRAKL MESSAGE - REVIEW BEFORE SENDING")
        print("=" * 60)

        print("No existing customer messages.")

        print("-" * 60)

    # ========================================================
    # FINAL CUSTOMER SEND GATE
    # ========================================================

    print()
    print("=" * 60)
    print("WARNING: THIS IS THE FINAL CUSTOMER SEND GATE")
    print("=" * 60)

    print()
    print(
        "The next stage will send this message "
        "to the customer through Mirakl."
    )

    print()
    print(
        "To authorise this exact order, type:"
    )

    expected_confirmation = (
        f"SEND {marketplace_order.order_id}"
    )

    print()
    print(expected_confirmation)
    print()

    confirmation = input(
        "Final confirmation: "
    ).strip()

    if confirmation != expected_confirmation:

        print()
        print("✗ LIVE SEND NOT AUTHORISED")

        print()
        print(
            "The confirmation did not exactly match:"
        )

        print(expected_confirmation)

        print()
        print(
            "No Mirakl message has been sent."
        )

        input(
            "\nPress ENTER to close..."
        )

        browser.close()
        raise SystemExit

    print()
    print("✓ LIVE SEND AUTHORISATION ACCEPTED")

    print()
    print(
        f"Order {marketplace_order.order_id} "
        "has passed the final operator gate."
    )

    # --------------------------------------------------------
    # LIVE SEND
    # --------------------------------------------------------

    print()
    print("=" * 60)
    print("LIVE MIRAKL SEND")
    print("=" * 60)

    send_result = send_mirakl_delivery_message(
        order_id=marketplace_order.order_id,
        message_body=customer_message,
        invoice_path=invoice_path,
    )

    print()
    print("✓ MIRAKL SEND REQUEST COMPLETED")

    print(
        f"HTTP Status: "
        f"{send_result['status_code']}"
    )

    # ========================================================
    # POST-SEND MIRAKL VERIFICATION
    # ========================================================

    print()
    print("=" * 60)
    print("POST-SEND MIRAKL VERIFICATION")
    print("=" * 60)

    print()
    print(
        "Waiting briefly for Mirakl to register "
        "the new message..."
    )

    time.sleep(2)

    print(
        "Reading the Mirakl conversation back "
        "to verify the customer message..."
    )

    verification_threads = read_mirakl_threads(
        marketplace_order.order_id
    )

    verification_messages = verification_threads[
        "messages"
    ]

    if not verification_messages:

        print()
        print("✗ POST-SEND VERIFICATION FAILED")
        print()
        print(
            "Mirakl returned no messages after the send."
        )

        print()
        print(
            "IMPORTANT: The API send request was made, "
            "so do NOT automatically retry."
        )

        input(
            "\nCheck the order manually in Mirakl. "
            "Press ENTER to close..."
        )

        browser.close()
        raise SystemExit

    latest_message = verification_messages[-1]

    verification_failures = []

    # --------------------------------------------------------
    # Sender
    # --------------------------------------------------------

    if latest_message["sender_type"] not in (
        "SHOP",
        "SHOP_USER",
    ):
        verification_failures.append(
            "Latest message was not sent by the shop"
        )

    # --------------------------------------------------------
    # Topic
    # --------------------------------------------------------

    if (
        str(latest_message["topic_code"])
        != str(DIY_DELIVERY_TOPIC_CODE)
    ):
        verification_failures.append(
            "Latest message does not use delivery topic 44"
        )

    # --------------------------------------------------------
    # Message body
    # --------------------------------------------------------

    expected_body = customer_message.strip()

    actual_body = (
        latest_message["body"] or ""
    ).strip()

    if actual_body != expected_body:
        verification_failures.append(
            "Latest Mirakl message body does not match "
            "the message that was sent"
        )

    # --------------------------------------------------------
    # RESULT
    # --------------------------------------------------------

    if verification_failures:

        print()
        print("✗ POST-SEND VERIFICATION FAILED")
        print()

        for failure in verification_failures:
            print(
                f"✗ {failure}"
            )

        print()
        print(
            "IMPORTANT: The API send request was already made."
        )

        print(
            "Do NOT automatically retry this order."
        )

        print()
        print(
            "Check the Mirakl conversation manually."
        )

        input(
            "\nPress ENTER to close..."
        )

        browser.close()
        raise SystemExit

    print()
    print("✓ POST-SEND VERIFICATION PASSED")
    print()
    print(
        f"Order:      {marketplace_order.order_id}"
    )
    print(
        f"Sender:     {latest_message['sender_name']}"
    )
    print(
        f"Topic Code: {latest_message['topic_code']}"
    )
    print(
        f"Date:       {latest_message['date']}"
    )

    print()
    print("=" * 60)
    print("✓ CUSTOMER MESSAGE VERIFIED IN MIRAKL")
    print("=" * 60)

    print()
    print(
        "The Mirakl customer message was sent and "
        "read back successfully."
    )

    input(
        "\nCustomer communication is complete. "
        "Press ENTER to continue to Mirakl shipping..."
    )

    complete_mirakl_shipping(
        browser=browser,
        marketplace_order=marketplace_order,
        mirakl_order=mirakl_order,
    )
