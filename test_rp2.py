import os
import re

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
    read_first_bq_order,
    read_mirakl_order,
    read_mirakl_threads,
    build_mirakl_dry_run_payload,
)

from messaging import (
    build_sgk_message,
    build_ed_message,
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

# Temporary test order.
# We will replace this with the B&Q Excel import next.
TEST_ORDER = None



# ============================================================
# HELPER FUNCTIONS
# ============================================================


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


excel_order = read_first_bq_order(
        ORDERS_FILE
    )

print("\nEXCEL DEBUG:")
print(excel_order)

TEST_ORDER = excel_order.order_id

print()
print("=" * 60)
print("B&Q EXCEL ORDER")
print("=" * 60)
print(
    f"Order ID:      {excel_order.order_id}"
)
print(
    f"SKU:           {excel_order.sku}"
)
print(
    f"Price:         £{excel_order.price:.2f}"
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
        f"Excel Order:   {excel_order.order_id}"
    )

    print(
        f"RPii Web Ref:  {web_ref}"
    )

    if web_ref == excel_order.order_id:

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
        excel_order.postcode
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
        f"Excel Postcode: {excel_order.postcode}"
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
        excel_order.sku
    ).strip().upper()

    rp2_sku = str(
        sku
    ).strip().upper()

    print(
        f"Excel SKU:     {excel_sku}"
    )

    print(
        f"RPii SKU:      {rp2_sku}"
    )

    if excel_sku == rp2_sku:

        print()
        print("✓ SKU MATCH")

    else:

        print()
        print("✗ SKU MISMATCH")
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
        float(excel_order.price),
        2
    )

    rp2_amount = round(
        float(total_sale),
        2
    )

    print(
        f"Excel Amount:  £{excel_amount:.2f}"
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
        excel_order.phone_1
    )

    excel_phone_2 = normalise_phone(
        excel_order.phone_2
    )

    rp2_phone = normalise_phone(
        delivery_customer["telephone"]
    )

    print(
        f"Excel Phone 1: {excel_order.phone_1}"
    )

    print(
        f"Excel Phone 2: {excel_order.phone_2}"
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
        excel_order.order_id
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
        excel_order,
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

    if can_override_sku_only(
        safety_result
    ):

        excel_sku = safety_values[
            "excel_sku"
        ]

        rp2_sku = safety_values[
            "rp2_sku"
        ]

        mirakl_sku = safety_values[
            "mirakl_sku"
        ]

        print()
        print("=" * 60)
        print("MANUAL OVERRIDE AVAILABLE")
        print("=" * 60)

        print(
            "Excel SKU:     ",
            excel_sku
        )

        print(
            "RPii SKU:      ",
            rp2_sku
        )

        print(
            "Mirakl SKU:    ",
            mirakl_sku
        )

        print()
        print(
            "Excel and RPii agree, but Mirakl "
            "contains a different SKU."
        )

        print()
        print(
            "Type OVERRIDE to continue with "
            "this order."
        )

        confirmation = input(
            "\nOverride SKU mismatch? "
        ).strip().upper()

        if confirmation == "OVERRIDE":

            override_used = True

            # We deliberately remove ONLY the SKU
            # failure after explicit user approval.
            safety_failures.remove(
                "SKU"
            )

            print()
            print(
                "⚠ SKU OVERRIDE ACCEPTED"
            )

            print(
                "Continuing using the "
                "Excel/RPii SKU."
            )

        else:

            print()
            print(
                "Override not accepted."
            )

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
            "Using Excel/RPii SKU: "
            f"{safety_values['excel_sku']}"
        )

        print(
            "Ignoring Mirakl SKU:  "
            f"{safety_values['mirakl_sku']}"
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
            "Excel, RPii and Mirakl agree."
        )

    # ========================================================
    # MIRAKL CUSTOMER COMMUNICATION SAFETY CHECK
    # ========================================================

    print()
    print("=" * 60)
    print("MIRAKL COMMUNICATION SAFETY CHECK")
    print("=" * 60)

    mirakl_threads = read_mirakl_threads(
        excel_order.order_id
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

            if str(latest_message["topic_code"]) == DIY_DELIVERY_TOPIC_CODE:

                print(
                    "⚠ EXISTING DELIVERY MESSAGE FOUND"
                )

                print()
                print(
                    "The latest Mirakl message was sent "
                    "by the shop using delivery topic 44."
                )

                print(
                    "Automatic customer-message processing "
                    "has been paused to avoid a duplicate."
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

                print()
                print(
                    "No new customer message has been sent."
                )

                input(
                    "\nReview the existing delivery message. "
                    "Press ENTER to close..."
                )

                browser.close()
                raise SystemExit

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
            f"{excel_order.order_id}-"
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

    # ========================================================
    # BUILD CUSTOMER MESSAGE PREVIEW
    # ========================================================

    print()
    print("CUSTOMER MESSAGE PREVIEW")
    print("=" * 60)

    courier = delivery["courier"]

    if courier == "SK":

        customer_date_text = (
            delivery["customer_date"]
            .strftime("%d/%m/%Y")
        )

        customer_message = build_sgk_message(
            delivery_customer,
            sku,
            description,
            customer_date_text,
        )

    elif courier == "ED":

        customer_date_text = (
            delivery["customer_date"]
            .strftime("%d/%m/%Y")
        )

        customer_message = build_ed_message(
            delivery_customer,
            sku,
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

    print()
    print(
        f"Order:       {excel_order.order_id}"
    )

    print(
        f"Customer:    {delivery_customer['name']}"
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
    # MIRAKL SEND PAYLOAD - DRY RUN ONLY
    # ========================================================

    dry_run_payload = build_mirakl_dry_run_payload(
        order_id=excel_order.order_id,
        customer_name=customer_name,
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
        != excel_order.order_id
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
    print("NOTHING HAS BEEN SENT.")

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

    print()
    print("DELIVERY CUSTOMER")
    print("-" * 60)

    print(
        f"Name:          {delivery_customer['name']}"
    )

    print(
        f"Address:       {delivery_customer['address']}"
    )

    print(
       f"Postcode:      {delivery_postcode}"
    )

    print(
        f"Telephone:     {delivery_customer['telephone']}"
    )

    print()
    print("ORDER ROUTE")
    print("-" * 60)
    print(f"Route:         {order_route}")

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
        "\nSAFE TEST MODE: "
        "No Mirakl customer message has been sent."
    )

    input(
        "\nPress ENTER to close..."
    )

    browser.close() 