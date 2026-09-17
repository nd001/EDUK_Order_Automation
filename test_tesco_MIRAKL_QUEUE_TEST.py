import json
import os
import re
import time
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

from config import (
    MODE,
    RP2_URL,
    INVOICE_FOLDER,
    RP2_REMOTE_USERNAME,
    RP2_REMOTE_PASSWORD,
)

from marketplaces.tesco import (
    read_tesco_orders,
    read_tesco_shipping_orders,
    build_tesco_rp2_search_ref,
    read_tesco_mirakl_order,
    read_tesco_mirakl_threads,
    build_tesco_mirakl_dry_run_payload,
    send_tesco_mirakl_delivery_message,
    TESCO_DELIVERY_TOPIC_TYPE,
    TESCO_DELIVERY_SUBJECT,
)

from rp2 import (
    extract_web_ref,
    extract_rp2_order_number,
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
)

from messaging import (
    build_sgk_message,
    build_ed_message,
)


load_dotenv()

RPII_USERNAME = os.getenv("RPII_USERNAME")
RPII_PASSWORD = os.getenv("RPII_PASSWORD")


TESCO_ORDER_FILE = "Tescoorders.xlsx"

TESCO_RP2_CACHE_FILE = (
    Path(__file__).resolve().parent
    / "tesco_rp2_tracking_cache.json"
)



def write_tesco_rp2_tracking_cache(
    order_id,
    rp2_order_number,
    rp2_courier,
):
    """
    Persist only the RPii/Stream details required by the later
    Tesco tracking-completion script.

    This is called only after the Tesco Mirakl delivery message
    has been successfully read back and verified.
    """

    order_id = str(order_id or "").strip()
    rp2_order_number = str(rp2_order_number or "").strip()
    rp2_courier = str(rp2_courier or "").strip().upper()

    if not order_id:
        raise ValueError("Cannot cache a blank Tesco order ID.")

    if not rp2_order_number:
        raise ValueError(
            "Cannot cache a blank RPii / Stream order number."
        )

    if rp2_courier == "SK":
        stream_courier = "SGK"
    elif rp2_courier == "ED":
        stream_courier = "ED"
    else:
        raise ValueError(
            f"Cannot cache unrecognised RPii courier: "
            f"{rp2_courier!r}"
        )

    cache = {}

    if TESCO_RP2_CACHE_FILE.exists():
        try:
            with TESCO_RP2_CACHE_FILE.open(
                "r",
                encoding="utf-8",
            ) as file:
                loaded = json.load(file)
        except json.JSONDecodeError as exc:
            raise RuntimeError(
                "Existing Tesco RPii tracking cache is not "
                "valid JSON. It has NOT been overwritten."
            ) from exc

        if not isinstance(loaded, dict):
            raise RuntimeError(
                "Existing Tesco RPii tracking cache does not "
                "contain a JSON object. It has NOT been overwritten."
            )

        cache = loaded

    cache[order_id] = {
        "rp2_order_number": rp2_order_number,
        "courier": stream_courier,
        "cached_at": datetime.now().isoformat(
            timespec="seconds"
        ),
        "message_verified": True,
        "tracking_completed": False,
    }

    temp_file = TESCO_RP2_CACHE_FILE.with_suffix(
        ".json.tmp"
    )

    with temp_file.open(
        "w",
        encoding="utf-8",
    ) as file:
        json.dump(
            cache,
            file,
            indent=2,
            sort_keys=True,
        )
        file.write("\n")

    temp_file.replace(TESCO_RP2_CACHE_FILE)

def read_tesco_processing_cache():
    """
    Read the Tesco RPii tracking cache for queue/status purposes only.

    READ ONLY:
    - Does not modify the cache.
    - Does not contact RPii.
    - Does not contact Go2Stream.
    - Does not perform any marketplace write.
    """

    if not TESCO_RP2_CACHE_FILE.exists():
        return {}

    try:
        with TESCO_RP2_CACHE_FILE.open(
            "r",
            encoding="utf-8",
        ) as file:
            cache = json.load(file)

    except (json.JSONDecodeError, OSError) as exc:
        print()
        print("⚠ TESCO PROCESSING CACHE COULD NOT BE READ")
        print(str(exc))
        print()
        print(
            "For safety, no orders will be treated as "
            "previously processed from the local cache."
        )
        return {}

    if not isinstance(cache, dict):
        print()
        print("⚠ TESCO PROCESSING CACHE HAS INVALID FORMAT")
        print(
            "For safety, no orders will be treated as "
            "previously processed from the local cache."
        )
        return {}

    return cache


def is_tesco_message_verified(order_id, cache):
    """
    Return True only when this exact Tesco order has a cache
    entry explicitly marked message_verified=True.
    """

    order_id = str(order_id or "").strip()

    entry = cache.get(order_id)

    if not isinstance(entry, dict):
        return False

    return entry.get("message_verified") is True


def get_customer_surname(full_name):
    """
    Extract a safe surname for use in an invoice filename.
    """

    if not full_name:
        return "Unknown"

    parts = full_name.strip().split()

    if not parts:
        return "Unknown"

    surname = parts[-1]

    surname = re.sub(
        r'[<>:"/\\|?*]',
        '',
        surname
    )

    return surname.title()



def format_tesco_delivery_deadline(mirakl_order):
    """
    Return Tesco Mirakl's latest customer delivery date as DD/MM/YYYY.

    IMPORTANT:
    This uses the customer delivery window (delivery_date.latest),
    NOT Mirakl's shipping_deadline.
    """

    delivery_date = mirakl_order.get("delivery_date")

    if isinstance(delivery_date, dict):
        delivery_deadline_raw = delivery_date.get("latest")
    else:
        delivery_deadline_raw = None

    # Compatibility fallback in case the Tesco adapter is later changed
    # to expose the same flattened field used by the B&Q adapter.
    if not delivery_deadline_raw:
        delivery_deadline_raw = mirakl_order.get(
            "delivery_date_latest"
        )

    if not delivery_deadline_raw:
        raise RuntimeError(
            "Tesco Mirakl did not return delivery_date.latest.\n"
            "The RPii Delivery Notes have NOT been changed."
        )

    deadline_value = str(delivery_deadline_raw).strip()

    if deadline_value.endswith("Z"):
        deadline_value = deadline_value[:-1] + "+00:00"

    try:
        return datetime.fromisoformat(
            deadline_value
        ).strftime("%d/%m/%Y")
    except ValueError as exc:
        raise RuntimeError(
            "Tesco Mirakl delivery_date.latest could not be "
            "converted to DD/MM/YYYY.\n"
            f"Raw value: {delivery_deadline_raw!r}\n"
            "The RPii Delivery Notes have NOT been changed."
        ) from exc


def add_tesco_rp2_delivery_deadline_note(
    sales_frame,
    mirakl_order,
):
    """
    Add the Tesco Mirakl delivery deadline to RPii Delivery Notes.

    Safety behaviour:
    - Preserve existing Delivery Notes.
    - Do not add the same deadline note twice.
    - Replace only an older automated deadline line.
    - Respect RPii's 1000-character Delivery Notes limit.
    - Click Save Notes once only.
    - Re-open notes and verify the saved text.
    """

    delivery_deadline = format_tesco_delivery_deadline(
        mirakl_order
    )

    deadline_note = (
        "**** IMPORTANT - DELIVERY DEADLINE - "
        f"{delivery_deadline} ****"
    )

    print()
    print("=" * 70)
    print("TESCO RPii DELIVERY DEADLINE NOTE")
    print("=" * 70)
    print(
        f"Mirakl delivery deadline: {delivery_deadline}"
    )
    print(
        f"Required RPii note:       {deadline_note}"
    )

    add_notes_button = sales_frame.locator(
        'img[title="Add notes to sale"]'
    )

    if add_notes_button.count() != 1:
        raise RuntimeError(
            "SAFETY STOP\n"
            "Could not uniquely identify the RPii "
            "'Add notes to sale' button."
        )

    print()
    print("Opening RPii Delivery Notes...")

    add_notes_button.evaluate(
        "element => element.click()"
    )

    delivery_note_field = sales_frame.locator(
        "#tDelNote"
    )

    delivery_note_field.wait_for(
        state="visible",
        timeout=10000,
    )

    existing_note = (
        delivery_note_field.input_value()
        or ""
    ).strip()

    if deadline_note in existing_note:
        print()
        print(
            "✓ Correct delivery deadline note already exists."
        )
        print("No RPii note write is required.")

        try:
            delivery_note_field.evaluate(
                "element => closeAlert()"
            )
        except Exception:
            pass

        return deadline_note

    deadline_pattern = re.compile(
        r"^\*\*\*\* IMPORTANT - DELIVERY DEADLINE - "
        r"\d{2}/\d{2}/\d{4} \*\*\*\*$",
        re.MULTILINE,
    )

    if deadline_pattern.search(existing_note):
        new_note = deadline_pattern.sub(
            deadline_note,
            existing_note,
            count=1,
        )
        print()
        print(
            "Existing automated delivery deadline found."
        )
        print(
            "It will be replaced; all other Delivery Notes "
            "will be preserved."
        )
    elif existing_note:
        new_note = (
            existing_note.rstrip()
            + "\n"
            + deadline_note
        )
    else:
        new_note = deadline_note

    if len(new_note) > 1000:
        raise RuntimeError(
            "SAFETY STOP\n"
            "Adding the delivery deadline would exceed "
            "RPii's 1000-character Delivery Notes limit.\n"
            "No RPii note has been saved."
        )

    print()
    print("Existing Delivery Notes:")
    print("-" * 70)
    print(existing_note or "(blank)")
    print("-" * 70)
    print()
    print("New Delivery Notes:")
    print("-" * 70)
    print(new_note)
    print("-" * 70)

    delivery_note_field.fill(new_note)

    entered_note = (
        delivery_note_field.input_value()
        or ""
    )

    if entered_note != new_note:
        raise RuntimeError(
            "SAFETY STOP\n"
            "RPii Delivery Notes field did not contain "
            "the expected text before saving.\n"
            "Save Notes has NOT been clicked."
        )

    save_notes_button = sales_frame.locator(
        'img[onclick*="saveCustNote("]'
        '[onclick*="closeAlert()"]'
    )

    save_button_count = save_notes_button.count()

    if save_button_count != 1:
        raise RuntimeError(
            "SAFETY STOP\n"
            "Could not uniquely identify the RPii "
            "Save Notes button.\n"
            f"Matching controls found: {save_button_count}\n"
            "The note has NOT been saved."
        )

    print()
    print("Saving RPii Delivery Notes...")
    print("One Save Notes click only.")

    save_notes_button.evaluate(
        "element => element.click()"
    )

    try:
        delivery_note_field.wait_for(
            state="hidden",
            timeout=10000,
        )
    except Exception:
        pass

    print("✓ Save Notes action submitted.")
    print("Waiting for RPii to finish saving...")
    time.sleep(1.0)
    print("Re-opening Delivery Notes to verify...")

    add_notes_button = sales_frame.locator(
        'img[title="Add notes to sale"]'
    )

    add_notes_button.evaluate(
        "element => element.click()"
    )

    verification_field = sales_frame.locator(
        "#tDelNote"
    )

    verification_field.wait_for(
        state="visible",
        timeout=10000,
    )

    verification_deadline = time.monotonic() + 10.0
    saved_note = ""

    while time.monotonic() < verification_deadline:
        saved_note = (
            verification_field.input_value()
            or ""
        )

        if saved_note == new_note:
            break

        time.sleep(0.5)

    if saved_note != new_note:
        print()
        print("=" * 70)
        print("RPii DELIVERY NOTES READ-BACK DIAGNOSTIC")
        print("=" * 70)
        print("Expected saved note:")
        print(repr(new_note))
        print()
        print("RPii read-back:")
        print(repr(saved_note))
        print()
        print(f"Expected length: {len(new_note)}")
        print(f"Read-back length: {len(saved_note)}")
        print("=" * 70)

        raise RuntimeError(
            "RPii Delivery Notes save could not be exactly "
            "verified after waiting for the saved value.\n"
            "The Save Notes action may already have occurred.\n"
            "DO NOT automatically retry the note write."
        )

    print()
    print("✓ RPii DELIVERY DEADLINE NOTE VERIFIED")
    print(f"  {deadline_note}")

    try:
        verification_field.evaluate(
            "element => closeAlert()"
        )
    except Exception as exc:
        raise RuntimeError(
            "The deadline note was verified as saved, but "
            "the RPii notes window could not be closed "
            "cleanly after verification.\n"
            f"Reason: {exc}"
        ) from exc

    return deadline_note

def build_tesco_sgk_sms():
    """Build the Tesco customer SMS for SGK deliveries."""

    return (
        "Electrical Discount UK: "
        "Thanks for your Tesco order. "
        "Your delivery will be made by SGK. "
        "SGK will contact you directly about delivery. "
        "We have emailed further information. "
        "If you have any issues, call 01282 443850."
    )


def build_tesco_ed_sms(description, delivery_date):
    """Build the Tesco customer SMS for EDUK deliveries."""

    product_text = str(description or "").strip()

    if not product_text:
        product_text = "item"

    return (
        "Electrical Discount UK: "
        "Thanks for your Tesco order. "
        f"Your {product_text} is due {delivery_date}, 7am-7pm. "
        "We will email an estimated delivery window "
        "the day before. "
        "Queries: 01282 443850."
    )


def ensure_rpii_login(page):
    """
    Automatically log into the RPii application if the
    RPii sign-in form is displayed.

    The remote HTTP authentication is handled separately
    by RP2_REMOTE_USERNAME / RP2_REMOTE_PASSWORD.
    """

    login_box = page.locator('input[name="login"]')

    if login_box.count() > 0 and login_box.is_visible():
        print("RPii login page detected.")

        if not RPII_USERNAME or not RPII_PASSWORD:
            raise RuntimeError(
                "RPII_USERNAME / RPII_PASSWORD are missing from .env"
            )

        print("Logging into RPii...")

        login_box.fill(RPII_USERNAME)
        page.locator('input[name="password"]').fill(
            RPII_PASSWORD
        )
        page.locator('button[name="doLogin"]').click()

        page.wait_for_load_state(
            "domcontentloaded"
        )

        print("✓ RPii login submitted.")

    else:
        print("✓ RPii login not required.")


# ============================================================
# TESCO QUEUE
# ============================================================

def format_tesco_ship_by(shipping_deadline):
    """Format Mirakl's ISO shipping deadline as DD/MM/YYYY."""

    if not shipping_deadline:
        return "-"

    try:
        deadline_text = str(shipping_deadline).strip()

        # Tesco currently returns timestamps such as:
        # 2026-09-11T22:59:59.999Z
        if deadline_text.endswith("Z"):
            deadline_text = deadline_text[:-1] + "+00:00"

        deadline = datetime.fromisoformat(deadline_text)

        return deadline.strftime("%d/%m/%Y")

    except (TypeError, ValueError):
        # Do not block the queue merely because a deadline has an
        # unexpected format. Show a clear review marker instead.
        return "REVIEW"


def get_tesco_queue_mirakl_status(order):
    """
    Read the Tesco Mirakl order for queue display only.

    This is READ ONLY. It performs no Mirakl write action.
    """

    try:
        mirakl_order = read_tesco_mirakl_order(
            order.order_id
        )

    except Exception:
        return {
            "state": "ERROR",
            "ship_by": "-",
            "status": "⚠ REVIEW",
        }

    if not mirakl_order:
        return {
            "state": "NOT FOUND",
            "ship_by": "-",
            "status": "⚠ REVIEW",
        }

    state = str(
        mirakl_order.get(
            "order_state",
            ""
        )
    ).strip().upper()

    shipping_deadline = mirakl_order.get(
        "shipping_deadline",
        ""
    )

    ship_by = format_tesco_ship_by(
        shipping_deadline
    )

    if state == "SHIPPED":
        status = "✓ COMPLETE"

    elif state == "SHIPPING":
        status = "READY"

    else:
        status = "⚠ REVIEW"

    return {
        "state": state or "-",
        "ship_by": ship_by,
        "status": status,
    }


def print_tesco_queue(orders):

    processing_cache = read_tesco_processing_cache()
    print()
    print("=" * 112)
    print("TESCO ORDER QUEUE")
    print("=" * 112)

    print(
        f"{'':>3} "
        f"{'ORDER':<20} "
        f"{'SKU':<20} "
        f"{'PRICE':>10}   "
        f"{'MIRAKL STATE':<15} "
        f"{'SHIP BY':<12} "
        f"{'STATUS':<12}"
    )

    print("-" * 112)

    for index, order in enumerate(
        orders,
        start=1,
    ):

        mirakl_status = get_tesco_queue_mirakl_status(
            order
        )

        message_verified = is_tesco_message_verified(
        order.order_id,
        processing_cache,
        )

        mirakl_state = str(
            mirakl_status.get("state") or ""
        ).strip().upper()

        if mirakl_state == "SHIPPED":
            queue_status = "✓ COMPLETE"

        elif (
            mirakl_state == "SHIPPING"
            and message_verified
        ):
            queue_status = "WAITING FOR TRACKING"

        elif mirakl_state == "SHIPPING":
            queue_status = "NEW ORDER"

        else:
            queue_status = mirakl_status["status"]

        print(
            f"{index:>2}. "
            f"{order.order_id:<20} "
            f"{order.sku:<20} "
            f"£{order.price:>8.2f}   "
            f"{mirakl_status['state']:<15} "
            f"{mirakl_status['ship_by']:<12} "
            f"{queue_status:<22}"
        )

    print("=" * 112)



def select_tesco_order(orders):

    while True:

        selection = input(
            "\nSelect order number to inspect: "
        ).strip()

        try:
            selection_number = int(
                selection
            )

        except ValueError:

            print(
                "Please enter a number from the queue."
            )

            continue

        if not (
            1 <= selection_number <= len(orders)
        ):

            print(
                f"Please enter a number between "
                f"1 and {len(orders)}."
            )

            continue

        return orders[
            selection_number - 1
        ]


def show_selected_order(order):

    rp2_search_ref = build_tesco_rp2_search_ref(
        order.order_id
    )

    print()
    print("=" * 60)
    print("SELECTED TESCO ORDER")
    print("=" * 60)

    print(
        f"Marketplace:    {order.marketplace}"
    )

    print(
        f"Tesco Order ID: {order.order_id}"
    )

    print(
        f"RPii Search:    {rp2_search_ref}"
    )

    print(
        f"SKU:            {order.sku}"
    )

    print(
        f"Price:          £{order.price:.2f}"
    )

    print(
        f"Postcode:       {order.postcode}"
    )

    print(
        f"Phone 1:        {order.phone_1}"
    )

    print(
        f"Phone 2:        {order.phone_2}"
    )

    print("=" * 60)


# ============================================================
# READ-ONLY RPii TEST
# ============================================================

def inspect_tesco_order_in_rp2(order):

    rp2_search_ref = build_tesco_rp2_search_ref(
        order.order_id
    )

    print()
    print("=" * 60)
    print("TESCO RPii READ-ONLY TEST")
    print("=" * 60)

    print(
        f"Tesco Order: {order.order_id}"
    )

    print(
        f"RPii Search: {rp2_search_ref}"
    )

    print()
    print(
        "NO RPii WRITE ACTIONS WILL BE PERFORMED."
    )

    with sync_playwright() as playwright:

        browser = playwright.chromium.launch(
            headless=False
        )

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

        # ----------------------------------------------------
        # OPEN RPii
        # ----------------------------------------------------

        print()
        print("=" * 60)
        print("EDUK RPii ORDER READER")
        print("=" * 60)
        print(f"Mode: {MODE.upper()}")
        print(f"Opening: {RP2_URL}")
        print()

        print("Opening RPii...")

        max_attempts = 3
        rp2_ready = False

        for attempt in range(
            1,
            max_attempts + 1
        ):

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

                ensure_rpii_login(page)

                print(
                    "Outer RPii page loaded. "
                    "Waiting for Sales module..."
                )

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

                    page.wait_for_timeout(
                        5000
                    )

        if not rp2_ready:

            print()
            print(
                "✗ RPii could not be opened."
            )

            input(
                "\nPress ENTER to close..."
            )

            browser.close()

            return

        # ----------------------------------------------------
        # OPERATOR GATE
        # ----------------------------------------------------

        print()
        print(
            "RPii is ready."
        )

        print(
            f"The READ-ONLY search will use: "
            f"{rp2_search_ref}"
        )

        input(
            "\nPress ENTER to search RPii..."
        )

        # ----------------------------------------------------
        # SEARCH
        # ----------------------------------------------------

        print()
        print(
            f"Opening Tesco order "
            f"{rp2_search_ref}..."
        )

        sales_enquiry.fill(
            rp2_search_ref
        )

        sales_enquiry.press(
            "Enter"
        )

        total_sale_field = sales_frame.locator(
            "#qh_ordval"
        )

        total_sale_field.wait_for(
            state="visible",
            timeout=30000
        )

        page.wait_for_timeout(
            1000
        )

        # ----------------------------------------------------
        # CUSTOMER PANEL DIAGNOSTIC
        # ----------------------------------------------------

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

        for i in range(
            customer_panels.count()
        ):

            panel = customer_panels.nth(
                i
            )

            print()
            print(
                f"PANEL {i}"
            )

            print("TEXT:")

            print(
                repr(
                    panel.inner_text()
                )
            )

            print(
                "-" * 50
            )

        print("=" * 70)

                # ----------------------------------------------------
        # RPii INVOICE ACCOUNT
        # ----------------------------------------------------

        rp2_account = ""

        if customer_panels.count() > 0:

            account_panel_text = (
                customer_panels.nth(0)
                .inner_text()
                .strip()
            )

            if account_panel_text:

                rp2_account = (
                    account_panel_text
                    .split("•", 1)[0]
                    .strip()
                )


        # ----------------------------------------------------
        # WEB REF
        # ----------------------------------------------------

        web_panel = sales_frame.locator(
            "td.SALES-panel",
            has_text="Web Ref:"
        )

        web_panel_text = (
            web_panel.inner_text()
        )

        web_ref = extract_web_ref(
            web_panel_text
        )

        # ----------------------------------------------------
        # PRODUCT
        # ----------------------------------------------------

        product = read_active_product(
            sales_frame
        )

        rp2_sku = product[
            "sku"
        ]

        description = product[
            "description"
        ]

        quantity = product[
            "quantity"
        ]

        # ----------------------------------------------------
        # FINANCIAL
        # ----------------------------------------------------

        total_sale = (
            total_sale_field.input_value()
        )

        # ----------------------------------------------------
        # RPii CUSTOMER ACCOUNT
        # ----------------------------------------------------

        customer_panel = None

        for i in range(
            customer_panels.count()
        ):

            panel = customer_panels.nth(
                i
            )

            text = panel.inner_text()

            if "T:" in text:
                customer_panel = panel
                break

        if customer_panel is None:

            print()
            print(
                "✗ Could not find RPii delivery "
                "customer panel."
            )

            input(
                "\nPress ENTER to close..."
            )

            browser.close()
            return

        customer_text = (
            customer_panel.inner_text()
        )

        customer_name = extract_customer_name(
            customer_text
        )

        delivery_customer = extract_delivery_customer(
            customer_text
        )

        delivery_postcode = extract_uk_postcode(
            delivery_customer["address"]
        )

        # ----------------------------------------------------
        # TESCO MIRAKL - READ ONLY
        # ----------------------------------------------------

        print()
        print("=" * 60)
        print("TESCO MIRAKL READ-ONLY LOOKUP")
        print("=" * 60)

        mirakl_order = read_tesco_mirakl_order(
            order.order_id
        )

        if mirakl_order is None:

            print(
                "✗ Tesco Mirakl order was not found."
            )

            input(
                "\nPress ENTER to close..."
            )

            browser.close()
            return

        mirakl_customer = mirakl_order.get(
            "customer",
            {}
        )

        mirakl_shipping_address = (
            mirakl_customer.get(
                "shipping_address",
                {}
            )
        )

        mirakl_order_lines = mirakl_order.get(
            "order_lines",
            []
        )

        if not mirakl_order_lines:

            print(
                "✗ Tesco Mirakl order has no order lines."
            )

            input(
                "\nPress ENTER to close..."
            )

            browser.close()
            return

        mirakl_first_line = (
            mirakl_order_lines[0]
        )

        mirakl_order_id = mirakl_order.get(
            "order_id",
            ""
        )

        mirakl_state = mirakl_order.get(
            "order_state",
            ""
        )

        mirakl_sku = mirakl_first_line.get(
            "offer_sku",
            ""
        )

        mirakl_price = mirakl_order.get(
            "total_price",
            0
        )

        mirakl_postcode = (
            mirakl_shipping_address.get(
                "zip_code",
                ""
            )
        )

        mirakl_phone_1 = (
            mirakl_shipping_address.get(
                "phone",
                ""
            )
        )

        mirakl_phone_2 = (
            mirakl_shipping_address.get(
                "phone_secondary",
                ""
            )
        )

        mirakl_shipping_deadline = (
            mirakl_order.get(
                "shipping_deadline",
                ""
            )
        )

        print(
            f"Order ID:          "
            f"{mirakl_order_id}"
        )

        print(
            f"State:             "
            f"{mirakl_state}"
        )

        print(
            f"SKU:               "
            f"{mirakl_sku}"
        )

        print(
            f"Price:             "
            f"£{mirakl_price}"
        )

        print(
            f"Postcode:          "
            f"{mirakl_postcode}"
        )

        print(
            f"Phone 1:           "
            f"{mirakl_phone_1}"
        )

        print(
            f"Phone 2:           "
            f"{mirakl_phone_2}"
        )

        print(
            f"Shipping deadline: "
            f"{mirakl_shipping_deadline}"
        )

        print("=" * 60)

        # ----------------------------------------------------
        # DISPLAY OBSERVED RPii DATA
        # ----------------------------------------------------

        print()
        print("=" * 60)
        print("TESCO RPii OBSERVED DATA")
        print("=" * 60)

        print(
            f"Tesco Full ID:  {order.order_id}"
        )

        print(
            f"RPii Search:    {rp2_search_ref}"
        )

        print(
            f"RPii Web Ref:   {web_ref}"
        )

        print()

        print(
            f"RPii Customer:  {customer_name}"
        )

        print(
            f"RPii SKU:       {rp2_sku}"
        )

        print(
            f"Description:    {description}"
        )

        print(
            f"Quantity:       {quantity}"
        )

        print(
            f"RPii Total:     £{total_sale}"
        )

        print(
            f"RPii Postcode:  {delivery_postcode}"
        )

        print(
            f"RPii Telephone: "
            f"{delivery_customer.get('telephone', '')}"
        )

        print("=" * 60)


        # ====================================================
        # TESCO THREE-WAY READ-ONLY SAFETY CHECKS
        # Spreadsheet ↔ RPii ↔ Mirakl
        # ====================================================

        print()
        print("=" * 60)
        print("TESCO THREE-WAY SAFETY CHECKS - READ ONLY")
        print("=" * 60)

        safety_failures = []

        # ----------------------------------------------------
        # RPii INVOICE ACCOUNT
        # ----------------------------------------------------

        print()
        print("RPii INVOICE ACCOUNT")
        print("-" * 60)

        expected_rp2_account = "TESC2386"

        print(
            f"Expected:    {expected_rp2_account}"
        )

        print(
            f"RPii:        {rp2_account}"
        )

        if rp2_account == expected_rp2_account:

            print(
                "✓ TESCO INVOICE ACCOUNT MATCH"
            )

        else:

            print(
                "✗ WRONG RPii INVOICE ACCOUNT"
            )

            safety_failures.append(
                "RPii invoice account"
            )

        # ----------------------------------------------------
        # ORDER NUMBER
        # ----------------------------------------------------

        print()
        print("ORDER NUMBER")
        print("-" * 60)

        spreadsheet_order_id = str(
            order.order_id
        ).strip()

        rp2_order_id = str(
            web_ref
        ).strip()

        mirakl_order_id_check = str(
            mirakl_order_id
        ).strip()

        print(
            f"Spreadsheet: {spreadsheet_order_id}"
        )

        print(
            f"RPii:        {rp2_order_id}"
        )

        print(
            f"Mirakl:      {mirakl_order_id_check}"
        )

        if (
            spreadsheet_order_id
            == rp2_order_id
            == mirakl_order_id_check
        ):

            print(
                "✓ THREE-WAY ORDER NUMBER MATCH"
            )

        else:

            print(
                "✗ ORDER NUMBER MISMATCH"
            )

            safety_failures.append(
                "Order number"
            )

        # ----------------------------------------------------
        # SKU
        # ----------------------------------------------------

        print()
        print("SKU")
        print("-" * 60)

        spreadsheet_sku = normalise_sku(
            order.sku
        )

        rp2_normalised_sku = normalise_sku(
            rp2_sku
        )

        mirakl_normalised_sku = normalise_sku(
            mirakl_sku
        )

        print(
            f"Spreadsheet: {order.sku}"
        )

        print(
            f"RPii:        {rp2_sku}"
        )

        print(
            f"Mirakl:      {mirakl_sku}"
        )

        marketplace_skus_match = (
            spreadsheet_sku
            == mirakl_normalised_sku
        )

        rp2_sku_compatible = (
            rp2_normalised_sku
            == spreadsheet_sku
            or rp2_normalised_sku.startswith(
                spreadsheet_sku
            )
        )

        if (
            marketplace_skus_match
            and rp2_sku_compatible
        ):

            if (
                rp2_normalised_sku
                == spreadsheet_sku
            ):

                print(
                    "✓ THREE-WAY EXACT SKU MATCH"
                )

            else:

                print(
                    "✓ THREE-WAY COMPATIBLE SKU MATCH"
                )

                print(
                    "Tesco/Mirakl SKU is present "
                    "at the start of the RPii SKU."
                )

        else:

            print(
                "✗ SKU MISMATCH"
            )

            if not marketplace_skus_match:

                print(
                    "Tesco spreadsheet SKU does not "
                    "match Tesco Mirakl."
                )

            if not rp2_sku_compatible:

                print(
                    "RPii SKU is not compatible with "
                    "the Tesco SKU."
                )

            safety_failures.append(
                "SKU"
            )

        # ----------------------------------------------------
        # PRICE
        # ----------------------------------------------------

        print()
        print("PRICE")
        print("-" * 60)

        spreadsheet_price = normalise_price(
            order.price
        )

        rp2_price = normalise_price(
            total_sale
        )

        mirakl_price_check = normalise_price(
            mirakl_price
        )

        print(
            f"Spreadsheet: £{spreadsheet_price:.2f}"
        )

        print(
            f"RPii:        £{rp2_price:.2f}"
        )

        print(
            f"Mirakl:      £{mirakl_price_check:.2f}"
        )

        if (
            spreadsheet_price
            == rp2_price
            == mirakl_price_check
        ):

            print(
                "✓ THREE-WAY PRICE MATCH"
            )

        else:

            print(
                "✗ PRICE MISMATCH"
            )

            safety_failures.append(
                "Price"
            )

        # ----------------------------------------------------
        # POSTCODE
        # ----------------------------------------------------

        print()
        print("POSTCODE")
        print("-" * 60)

        spreadsheet_postcode = normalise_postcode(
            order.postcode
        )

        rp2_normalised_postcode = normalise_postcode(
            delivery_postcode
        )

        mirakl_normalised_postcode = normalise_postcode(
            mirakl_postcode
        )

        print(
            f"Spreadsheet: {order.postcode}"
        )

        print(
            f"RPii:        {delivery_postcode}"
        )

        print(
            f"Mirakl:      {mirakl_postcode}"
        )

        if (
            spreadsheet_postcode
            == rp2_normalised_postcode
            == mirakl_normalised_postcode
        ):

            print(
                "✓ THREE-WAY POSTCODE MATCH"
            )

        else:

            print(
                "✗ POSTCODE MISMATCH"
            )

            safety_failures.append(
                "Postcode"
            )

        # ----------------------------------------------------
        # TELEPHONE
        # ----------------------------------------------------

        print()
        print("TELEPHONE")
        print("-" * 60)

        spreadsheet_phones = {
            normalise_phone(
                order.phone_1
            ),
            normalise_phone(
                order.phone_2
            ),
        }

        spreadsheet_phones.discard("")

        rp2_phone = normalise_phone(
            delivery_customer.get(
                "telephone",
                ""
            )
        )

        mirakl_phones = {
            normalise_phone(
                mirakl_phone_1
            ),
            normalise_phone(
                mirakl_phone_2
            ),
        }

        mirakl_phones.discard("")

        print(
            f"Spreadsheet Phone 1: "
            f"{order.phone_1}"
        )

        print(
            f"Spreadsheet Phone 2: "
            f"{order.phone_2}"
        )

        print(
            f"RPii:                "
            f"{delivery_customer.get('telephone', '')}"
        )

        print(
            f"Mirakl Phone 1:       "
            f"{mirakl_phone_1}"
        )

        print(
            f"Mirakl Phone 2:       "
            f"{mirakl_phone_2}"
        )

        common_phones = (
            spreadsheet_phones
            & {rp2_phone}
            & mirakl_phones
        )

        if common_phones:

            print(
                "✓ THREE-WAY TELEPHONE MATCH"
            )

        else:

            print(
                "✗ TELEPHONE MISMATCH"
            )

            safety_failures.append(
                "Telephone"
            )

        # ----------------------------------------------------
        # SUMMARY
        # ----------------------------------------------------

        print()
        print("=" * 60)
        print("TESCO THREE-WAY SAFETY RESULT")
        print("=" * 60)

        if safety_failures:

            print(
                "⚠ REVIEW REQUIRED"
            )

            print(
                "Issues found: "
                + ", ".join(
                    safety_failures
                )
            )

            print()
            print(
                "NO WRITE ACTIONS MAY PROCEED."
            )

            print("=" * 60)

            input(
                "\nPress ENTER to close..."
            )

            browser.close()
            return

        print(
            "✓ ALL TESCO THREE-WAY CHECKS PASSED"
        )

        print("=" * 60)

        # ====================================================
        # READ COURIER / DELIVERY INFORMATION
        # ====================================================

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

        courier = delivery["courier"]

        print()
        print("=" * 60)
        print("TESCO DELIVERY ROUTE")
        print("=" * 60)
        print(
            f"Courier:       {courier}"
        )
        print(
            f"Delivery date: {delivery['customer_date']}"
        )

        if courier not in ("SK", "ED"):

            print()
            print(
                "✗ Tesco delivery courier is not recognised."
            )
            print(
                "No invoice or SMS action will proceed."
            )

            input(
                "\nPress ENTER to close..."
            )

            browser.close()
            return

        if (
            courier == "ED"
            and delivery["customer_date"] is None
        ):

            print()
            print(
                "✗ Tesco EDUK delivery date could not be read."
            )
            print(
                "No invoice or SMS action will proceed."
            )

            input(
                "\nPress ENTER to close..."
            )

            browser.close()
            return

        print(
            "✓ Tesco delivery route recognised."
        )
        print("=" * 60)

        # ====================================================
        # ADD TESCO DELIVERY DEADLINE TO RPii DELIVERY NOTES
        # ====================================================
        #
        # This occurs only after the existing three-way safety
        # gate has passed and the delivery route is recognised.
        # It uses Mirakl delivery_date.latest, NOT shipping_deadline.
        # Any uncertainty stops processing before invoice / SMS.
        # ====================================================

        try:
            add_tesco_rp2_delivery_deadline_note(
                sales_frame=sales_frame,
                mirakl_order=mirakl_order,
            )
        except Exception as exc:
            print()
            print("=" * 60)
            print("✗ TESCO RPii DELIVERY DEADLINE NOTE FAILED")
            print("=" * 60)
            print(str(exc))
            print()
            print(
                "For safety, Tesco processing will stop here."
            )
            print(
                "No invoice / SMS / Mirakl customer message "
                "will be created or sent."
            )
            print("=" * 60)
            input("\nPress ENTER to close...")
            browser.close()
            return

        # ====================================================
        # TESCO INVOICE GENERATION
        # ====================================================

        print()
        print("=" * 60)
        print("TESCO INVOICE GENERATION")
        print("=" * 60)

        print(
            f"Order:    {order.order_id}"
        )

        print(
            f"Customer: {delivery_customer['name']}"
        )

        print()
        print(
            "All three-way safety checks have passed."
        )

        input(
            "\nPress ENTER to generate the A4 invoice..."
        )

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
            return

        # ----------------------------------------------------
        # Open Print options
        # ----------------------------------------------------

        print()
        print("Opening RPii Print panel...")

        print_button.evaluate(
            "element => element.click()"
        )

        page.wait_for_timeout(
            1000
        )

        print(
            "✓ Print panel opened."
        )

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
            return

        invoice_action = (
            a4_invoice.get_attribute(
                "onclick"
            )
        )

        print(
            f"A4 Invoice action: "
            f"{invoice_action}"
        )

        # ----------------------------------------------------
        # Open A4 Invoice
        # ----------------------------------------------------

        print(
            "Opening A4 Invoice..."
        )

        a4_invoice.evaluate(
            "element => element.click()"
        )

        page.wait_for_timeout(
            1500
        )

        # ====================================================
        # SAVE GENERATED INVOICE PDF
        # ====================================================

        print()
        print("SAVING INVOICE")
        print("=" * 60)

        print_frame = None

        for frame in page.frames:

            if frame.name == "PRINTFrame":

                print_frame = frame
                break

        if print_frame is None:

            print(
                "✗ PRINTFrame could not be found."
            )

            input(
                "\nPress ENTER to close..."
            )

            browser.close()
            return

        # ----------------------------------------------------
        # WAIT FOR PRINTFrame TO NAVIGATE TO THE PDF
        # ----------------------------------------------------
        #
        # RPii can create PRINTFrame immediately at blank.htm,
        # then navigate that frame to the generated invoice PDF
        # a moment later.  Do not treat blank.htm as a failure
        # until we have allowed the frame time to finish loading.
        #
        print(
            "Waiting for PRINTFrame to load the invoice PDF..."
        )

        invoice_url = ""

        for wait_attempt in range(1, 21):

            # Re-acquire PRINTFrame on every pass in case RPii
            # has replaced/reloaded it during invoice generation.
            print_frame = None

            for frame in page.frames:

                if frame.name == "PRINTFrame":

                    print_frame = frame
                    break

            if print_frame is not None:

                invoice_url = print_frame.url

                if invoice_url.lower().endswith(
                    ".pdf"
                ):

                    break

            page.wait_for_timeout(
                500
            )

        print(
            f"Invoice URL: {invoice_url}"
        )

        if not invoice_url.lower().endswith(
            ".pdf"
        ):

            print()
            print(
                "✗ PRINTFrame did not navigate to a PDF "
                "within 10 seconds."
            )

            print(
                "No invoice has been downloaded and no "
                "customer communication has been sent."
            )

            input(
                "\nPress ENTER to close..."
            )

            browser.close()
            return

        os.makedirs(
            INVOICE_FOLDER,
            exist_ok=True
        )

        customer_surname = (
            get_customer_surname(
                delivery_customer["name"]
            )
        )

        invoice_filename = (
            f"{order.order_id}-"
            f"{customer_surname}-invoice.pdf"
        )

        invoice_path = os.path.join(
            INVOICE_FOLDER,
            invoice_filename
        )

        print(
            f"Saving as: {invoice_filename}"
        )

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
            return

        with open(
            invoice_path,
            "wb"
        ) as file:

            file.write(
                response.body()
            )

        print()
        print(
            "✓ INVOICE SAVED SUCCESSFULLY"
        )

        print()

        print(
            f"Saved to: {invoice_path}"
        )

        # ====================================================
        # RETURN FROM INVOICE TO EPoS
        # ====================================================

        print()
        print(
            "Returning from invoice to EPoS..."
        )

        epos_tab = page.locator(
            "#SALEStd"
        )

        if epos_tab.count() != 1:

            print()
            print(
                "✗ Could not uniquely identify "
                "the EPoS tab."
            )

            input(
                "\nPress ENTER to close..."
            )

            browser.close()
            return

        epos_tab.evaluate(
            "element => element.click()"
        )

        page.wait_for_timeout(
            1000
        )

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
            return

        print(
            "✓ SALESFrame re-acquired."
        )

        current_sale_total = (
            sales_frame.locator(
                "#qh_ordval"
            )
        )

        current_sale_total.wait_for(
            state="visible",
            timeout=10000,
        )

        print(
            "✓ Returned to EPoS sale."
        )

        # ====================================================
        # BUILD TESCO SMS PREVIEW
        # ====================================================

        if courier == "SK":

            customer_sms = build_tesco_sgk_sms()

        elif courier == "ED":

            customer_date_text = (
                delivery["customer_date"]
                .strftime("%d/%m/%Y")
            )

            customer_sms = build_tesco_ed_sms(
                description=description,
                delivery_date=customer_date_text,
            )

        else:

            raise RuntimeError(
                f"Cannot build Tesco SMS for "
                f"unknown courier: {courier}"
            )

        sms_target_mobile = (
            delivery_customer.get(
                "telephone",
                ""
            )
        )

        print()
        print("=" * 60)
        print("TESCO CUSTOMER SMS PREVIEW")
        print("=" * 60)

        print(
            f"Order:   {order.order_id}"
        )

        print(
            f"Mobile:  {sms_target_mobile}"
        )

        print(
            f"Courier: {courier}"
        )

        print()
        print("-" * 60)
        print(customer_sms)
        print("-" * 60)

        print()
        print(
            "SMS PREVIEW ONLY - NO SMS HAS BEEN SENT."
        )

        input(
            "\nPress ENTER to populate the RPii SMS fields..."
        )

        # ====================================================
        # POPULATE, VERIFY AND GUARDED-SEND RPii SMS
        # ====================================================

        print()
        print(
            "Preparing RPii SMS fields..."
        )

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
            return

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

            page.wait_for_timeout(
                1000
            )

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
            return

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

        # Populate and verify the fields before the guarded live-send gate.

        sms_mobile.fill(
            sms_target_mobile
        )

        sms_text.fill(
            customer_sms
        )

        entered_mobile = (
            sms_mobile.input_value()
        )

        entered_sms = (
            sms_text.input_value()
        )

        print()
        print("=" * 60)
        print("TESCO RPii SMS FIELD VERIFICATION")
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
            entered_mobile == sms_target_mobile
            and entered_sms == customer_sms
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
            return

        # ====================================================
        # FINAL RPii SMS SEND GATE
        # ====================================================

        print()
        print("=" * 60)
        print("FINAL RPii SMS CONFIRMATION")
        print("=" * 60)

        print()
        print(f"Order:       {order.order_id}")
        print(f"Mobile:      {sms_target_mobile}")

        print()
        print("SMS:")
        print("-" * 60)
        print(customer_sms)
        print("-" * 60)

        expected_sms_confirmation = (
            f"TEXT {order.order_id}"
        )

        print()
        print("To authorise this exact SMS, type:")
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

            input("\nPress ENTER to close...")
            browser.close()
            return

        print()
        print("✓ SMS AUTHORISATION ACCEPTED")

        # ====================================================
        # LIVE RPii SMS SEND
        # ====================================================

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

        # IMPORTANT: click once only. Never automatically retry.
        sms_send_button.evaluate(
            "element => element.click()"
        )

        print()
        print(
            "Waiting for RPii SMS acknowledgement..."
        )

        acknowledgement_found = False
        acknowledgement_text = ""
        acknowledgement_frame = ""

        for attempt in range(20):

            for frame in page.frames:

                try:
                    alert_divs = frame.locator(
                        ".alertDiv"
                    )

                    for i in range(
                        alert_divs.count()
                    ):

                        alert_text = (
                            alert_divs.nth(i)
                            .text_content()
                            or ""
                        ).strip()

                        if expected_alert_text in alert_text:

                            acknowledgement_found = True
                            acknowledgement_text = alert_text
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
            print("✓ RPii SMS SEND ACKNOWLEDGED")
            print(
                f"✓ RPii reported SMS sent to "
                f"{sms_target_mobile}"
            )

        else:

            print("✗ RPii SMS AUTOMATIC VERIFICATION FAILED")
            print()
            print(
                f"Expected acknowledgement: "
                f"{expected_alert_text}"
            )
            print()
            print(
                "The SMS send action has already been attempted."
            )
            print(
                "DO NOT send the SMS again."
            )

            print()
            print("=" * 60)
            print("MANUAL SMS VERIFICATION")
            print("=" * 60)

            print()
            print(
                "If you have personally confirmed that the SMS "
                "was successfully sent, type:"
            )
            print()
            print("SMS VERIFIED")
            print()

            manual_verification = input(
                "Confirmation: "
            ).strip()

            if manual_verification != "SMS VERIFIED":

                print()
                print("✗ SMS NOT VERIFIED")
                print()
                print(
                    "The Tesco process will stop here."
                )

                input("\nPress ENTER to close...")
                browser.close()
                return

            print()
            print(
                "✓ SMS MANUALLY VERIFIED"
            )
            print(
                "Continuing Tesco processing..."
            )

        # ====================================================
        # TESCO MIRAKL CUSTOMER MESSAGE PREVIEW
        # ====================================================

        customer_delivery_details = dict(
            delivery_customer
        )
        customer_delivery_details["postcode"] = (
            delivery_postcode
        )

        customer_date_text = delivery["customer_date"].strftime(
            "%d/%m/%Y"
        )

        if courier == "SK":
            customer_message = build_sgk_message(
                customer_delivery_details,
                order.sku,
                product["description"],
                customer_date_text,
            )
        elif courier == "ED":
            customer_message = build_ed_message(
                customer_delivery_details,
                order.sku,
                product["description"],
                customer_date_text,
            )
        else:
            raise RuntimeError(
                f"Cannot build Tesco Mirakl message for "
                f"unknown courier: {courier}"
            )

        print()
        print("=" * 60)
        print("TESCO MIRAKL CUSTOMER MESSAGE PREVIEW")
        print("=" * 60)
        print(f"Order:      {order.order_id}")
        print(f"Customer:   {delivery_customer['name']}")
        print(f"Courier:    {courier}")
        print(f"Attachment: {invoice_filename}")
        print()
        print("-" * 60)
        print(customer_message)
        print("-" * 60)
        print()
        print("PREVIEW ONLY - NOTHING HAS BEEN SENT TO MIRAKL.")

        # ====================================================
        # TESCO MIRAKL COMMUNICATION SAFETY CHECK - READ ONLY
        # ====================================================

        print()
        print("=" * 60)
        print("TESCO MIRAKL COMMUNICATION SAFETY CHECK")
        print("=" * 60)
        print("Reading existing Tesco Mirakl conversation...")

        try:
            mirakl_threads = read_tesco_mirakl_threads(
                order.order_id
            )
        except Exception as exc:
            print()
            print("✗ TESCO MIRAKL THREAD LOOKUP FAILED")
            print(str(exc))
            print()
            print("STOPPING - no Mirakl write has been attempted.")
            input("\nPress ENTER to close...")
            browser.close()
            return

        existing_threads = mirakl_threads["threads"]
        existing_messages = mirakl_threads["messages"]

        if existing_threads or existing_messages:
            print()
            print("⚠ EXISTING MIRAKL COMMUNICATION DETECTED")
            print()
            print(f"Existing threads:  {len(existing_threads)}")
            print(f"Existing messages: {len(existing_messages)}")

            if existing_messages:
                latest_message = existing_messages[-1]

                print(
                    f"Latest sender type: "
                    f"{latest_message['sender_type']}"
                )
                print(
                    f"Latest sender:      "
                    f"{latest_message['sender_name']}"
                )
                print(
                    f"Latest date:        "
                    f"{latest_message['date']}"
                )
                print(
                    f"Latest topic type:  "
                    f"{latest_message['topic_type']}"
                )
                print(
                    f"Latest topic value: "
                    f"{latest_message['topic_value']}"
                )
                print()
                print("LATEST MESSAGE")
                print("-" * 60)
                print(latest_message["body"])
                print("-" * 60)

            print()
            print("✓ DUPLICATE MESSAGE PROTECTION ACTIVE")
            print(
                "This guarded first-live-send test will NOT create "
                "another Tesco customer thread."
            )
            print()
            print("NO MIRAKL WRITE ACTION HAS BEEN PERFORMED.")

            input("\nPress ENTER to close...")
            browser.close()
            return

        print()
        print("No existing Tesco Mirakl threads/messages found.")
        print("✓ CLEAR FOR GUARDED LIVE-SEND PREPARATION")

        # ====================================================
        # TESCO MIRAKL DRY-RUN PAYLOAD
        # ====================================================

        dry_run_payload = build_tesco_mirakl_dry_run_payload(
            order_id=order.order_id,
            customer_name=delivery_customer["name"],
            message_body=customer_message,
            invoice_path=invoice_path,
            existing_threads=existing_threads,
        )

        dry_run_failures = []

        if dry_run_payload["mode"] != "DRY_RUN":
            dry_run_failures.append(
                "Payload is not marked DRY_RUN"
            )

        if dry_run_payload["order_id"] != order.order_id:
            dry_run_failures.append(
                "Payload order number mismatch"
            )

        if (
            dry_run_payload["topic"]["type"]
            != TESCO_DELIVERY_TOPIC_TYPE
        ):
            dry_run_failures.append(
                "Incorrect Tesco Mirakl topic type"
            )

        if (
            dry_run_payload["topic"]["value"]
            != TESCO_DELIVERY_SUBJECT
        ):
            dry_run_failures.append(
                "Incorrect Tesco Mirakl subject"
            )

        if not dry_run_payload["message_body"].strip():
            dry_run_failures.append(
                "Customer message is blank"
            )

        if not dry_run_payload["attachment"]["exists"]:
            dry_run_failures.append(
                "Invoice attachment does not exist"
            )

        # The first live Tesco message should only be sent while
        # the Mirakl order is still in the SHIPPING state.
        if mirakl_state != "SHIPPING":
            dry_run_failures.append(
                f"Mirakl order state is {mirakl_state}, not SHIPPING"
            )

        print()
        print("=" * 60)
        print("TESCO MIRAKL SEND PAYLOAD - DRY RUN")
        print("=" * 60)
        print(f"Mode:            {dry_run_payload['mode']}")
        print(f"Order:           {dry_run_payload['order_id']}")
        print(f"Customer:        {dry_run_payload['customer']}")
        print(f"Mirakl state:    {mirakl_state}")
        print(f"Topic type:      {dry_run_payload['topic']['type']}")
        print(f"Subject:         {dry_run_payload['topic']['value']}")
        print(f"Existing thread: {dry_run_payload['existing_thread_id']}")
        print(f"Attachment:      {dry_run_payload['attachment']['filename']}")
        print(f"File exists:     {dry_run_payload['attachment']['exists']}")
        print()
        print("MESSAGE BODY")
        print("-" * 60)
        print(customer_message)
        print("-" * 60)

        if dry_run_failures:
            print()
            for failure in dry_run_failures:
                print(f"✗ {failure}")
            print()
            print("✗ TESCO MIRAKL DRY-RUN PAYLOAD FAILED")
            print("NO LIVE MIRAKL SEND WILL BE PERMITTED.")
            input("\nPress ENTER to close...")
            browser.close()
            return

        print()
        print("✓ Order number present and correct")
        print("✓ Mirakl order state is SHIPPING")
        print("✓ FREE_TEXT topic is correct")
        print("✓ Fixed Tesco subject is correct")
        print("✓ Message body is present")
        print("✓ Invoice attachment exists")
        print("✓ No existing Tesco Mirakl communication exists")
        print("✓ TESCO MIRAKL DRY-RUN PAYLOAD PASSED")

        # ====================================================
        # FINAL GUARDED LIVE-SEND GATE
        # ====================================================

        print()
        print("=" * 60)
        print("FINAL TESCO MIRAKL SEND CONFIRMATION")
        print("=" * 60)
        print(f"Order:       {order.order_id}")
        print(f"Customer:    {delivery_customer['name']}")
        print(f"Topic type:  {TESCO_DELIVERY_TOPIC_TYPE}")
        print(f"Subject:     {TESCO_DELIVERY_SUBJECT}")
        print(f"Attachment:  {invoice_filename}")
        print()
        print("MESSAGE")
        print("-" * 60)
        print(customer_message)
        print("-" * 60)
        print()
        print("WARNING: THE NEXT AUTHORISED ACTION SENDS A REAL")
        print("TESCO MIRAKL MESSAGE TO THIS CUSTOMER.")
        print()
        print("There will be exactly ONE POST attempt.")
        print("The script will NOT automatically retry it.")

        expected_confirmation = f"SEND {order.order_id}"

        print()
        print("To authorise this exact message, type:")
        print()
        print(expected_confirmation)
        print()

        confirmation = input(
            "Final Mirakl confirmation: "
        ).strip()

        if confirmation != expected_confirmation:
            print()
            print("✗ LIVE TESCO MIRAKL SEND NOT AUTHORISED")
            print()
            print("The confirmation did not exactly match:")
            print(expected_confirmation)
            print()
            print("NO MIRAKL MESSAGE HAS BEEN SENT.")
            input("\nPress ENTER to close...")
            browser.close()
            return

        print()
        print("✓ LIVE SEND AUTHORISATION ACCEPTED")
        print()
        print("=" * 60)
        print("LIVE TESCO MIRAKL SEND - ONE ATTEMPT ONLY")
        print("=" * 60)

        try:
            send_result = send_tesco_mirakl_delivery_message(
                order_id=order.order_id,
                message_body=customer_message,
                invoice_path=invoice_path,
            )
        except Exception as exc:
            print()
            print("⚠ TESCO MIRAKL POST DID NOT COMPLETE SUCCESSFULLY")
            print(str(exc))
            print()
            print("A POST attempt has already been made.")
            print("DO NOT automatically retry it.")
            print(
                "Check the Tesco Mirakl order manually before "
                "attempting anything else."
            )
            input("\nPress ENTER to close...")
            browser.close()
            return

        print()
        print("✓ TESCO MIRAKL POST RETURNED SUCCESS")
        print(f"HTTP status: {send_result['status_code']}")
        print(f"Endpoint:    {send_result['endpoint']}")

        # ====================================================
        # POST-SEND READ-BACK VERIFICATION
        # ====================================================

        print()
        print("=" * 60)
        print("TESCO MIRAKL POST-SEND VERIFICATION")
        print("=" * 60)
        print("Waiting briefly for Mirakl to register the thread...")

        time.sleep(2)

        try:
            verified_threads = read_tesco_mirakl_threads(
                order.order_id
            )
        except Exception as exc:
            print()
            print("⚠ POST-SEND READ-BACK FAILED")
            print(str(exc))
            print()
            print("The message POST has already succeeded.")
            print("DO NOT resend it automatically.")
            print("Check the order manually in Tesco Mirakl.")
            input("\nPress ENTER to close...")
            browser.close()
            return

        verification_failures = []
        matching_message = None

        for message in verified_threads["messages"]:
            if (
                message.get("topic_type")
                == TESCO_DELIVERY_TOPIC_TYPE
                and message.get("topic_value")
                == TESCO_DELIVERY_SUBJECT
                and message.get("sender_type")
                in ("SHOP", "SHOP_USER")
            ):
                matching_message = message

        if matching_message is None:
            verification_failures.append(
                "New Tesco FREE_TEXT delivery message not found"
            )
        else:
            expected_body = "\n".join(
                line.rstrip()
                for line in customer_message.strip().splitlines()
            ).strip()

            actual_body = "\n".join(
                line.rstrip()
                for line in matching_message["body"].strip().splitlines()
            ).strip()

            if actual_body != expected_body:
                verification_failures.append(
                    "Read-back message body does not exactly match"
                )

            attachment_names = {
                attachment.get("name")
                for attachment in matching_message.get(
                    "attachments",
                    [],
                )
            }

            if invoice_filename not in attachment_names:
                verification_failures.append(
                    "Invoice attachment filename not found on read-back"
                )

        print()
        print(f"Threads found:  {len(verified_threads['threads'])}")
        print(f"Messages found: {len(verified_threads['messages'])}")

        if matching_message:
            print(
                f"Verified topic:  "
                f"{matching_message['topic_type']} / "
                f"{matching_message['topic_value']}"
            )
            print(
                f"Verified sender: "
                f"{matching_message['sender_type']}"
            )
            print(
                "Attachments:     "
                + ", ".join(
                    attachment.get("name", "")
                    for attachment in matching_message.get(
                        "attachments",
                        [],
                    )
                )
            )

        print()

        if verification_failures:
            print("⚠ POST-SEND VERIFICATION FAILED")
            for failure in verification_failures:
                print(f"✗ {failure}")
            print()
            print("The customer message may already have been sent.")
            print("DO NOT automatically retry it.")
            print("Check the Tesco Mirakl conversation manually.")
        else:
            print("✓ FREE_TEXT subject verified")
            print("✓ Customer message body verified")
            print("✓ Invoice attachment verified")
            print()
            print("✓ TESCO MIRAKL MESSAGE SEND VERIFIED")

            # ------------------------------------------------
            # CACHE RPii / STREAM DETAILS FOR LATER TRACKING
            # ------------------------------------------------
            #
            # Important: this happens only after the customer
            # delivery message and invoice attachment have been
            # read back from Tesco Mirakl and verified.
            #
            # Cache failure must never cause a customer message
            # to be resent. RPii remains the tracking-script
            # fallback if this auxiliary cache cannot be written.
            #

            try:
                rp2_stream_order_number = (
                    extract_rp2_order_number(
                        web_panel_text
                    )
                )

                if not rp2_stream_order_number:
                    raise RuntimeError(
                        "RPii / Stream order number could not "
                        "be extracted from the verified sale."
                    )

                write_tesco_rp2_tracking_cache(
                    order_id=order.order_id,
                    rp2_order_number=rp2_stream_order_number,
                    rp2_courier=courier,
                )

                print()
                print("✓ RPii tracking details cached")
                print(
                    f"  Tesco order:       {order.order_id}"
                )
                print(
                    f"  RPii/Stream order: "
                    f"{rp2_stream_order_number}"
                )
                print(
                    f"  Stream courier:    "
                    f"{'SGK' if courier == 'SK' else 'ED'}"
                )
                print(
                    f"  Cache file:        "
                    f"{TESCO_RP2_CACHE_FILE.name}"
                )

            except Exception as exc:
                print()
                print(
                    "⚠ CUSTOMER MESSAGE IS VERIFIED, "
                    "BUT RPii CACHE WRITE FAILED"
                )
                print(str(exc))
                print(
                    "No customer communication will be retried. "
                    "The tracking script can fall back to RPii."
                )

        print()
        print("=" * 60)
        print("TESCO GUARDED MESSAGE TEST COMPLETE")
        print("=" * 60)
        print()
        print("NO STREAM ACTION HAS BEEN PERFORMED.")
        print("TESCO SHIPPING/TRACKING HAS NOT BEEN CHANGED.")

        input("\nPress ENTER to close...")
        browser.close()



# ============================================================
# MAIN
# ============================================================

def main():

    orders = read_tesco_shipping_orders()

    if not orders:
        print()
        print("No Tesco orders are currently awaiting shipment.")
        print("Nothing has been changed.")
        return

    print_tesco_queue(
        orders
    )

    selected_order = select_tesco_order(
        orders
    )

    show_selected_order(
        selected_order
    )

    # ========================================================
    # PREVIOUSLY-PROCESSED ORDER SAFETY GUARD
    # ========================================================
    #
    # This guard runs BEFORE RPii processing begins.
    #
    # A cache entry is trusted only when this exact Tesco
    # order is explicitly marked message_verified=True.
    #
    # No RPii, Go2Stream or marketplace write occurs here.
    # ========================================================

    processing_cache = read_tesco_processing_cache()

    if is_tesco_message_verified(
        selected_order.order_id,
        processing_cache,
    ):
        print()
        print("=" * 60)
        print("TESCO ORDER ALREADY PROCESSED")
        print("=" * 60)
        print()
        print(
            f"Order: {selected_order.order_id}"
        )
        print()
        print(
            "✓ Customer communication previously verified"
        )
        print(
            "⏳ Waiting for tracking"
        )
        print()
        print(
            "The invoice / SMS / Tesco message process "
            "will NOT run again."
        )
        print()
        print(
            "Use the Tesco tracking script to complete "
            "this order when tracking becomes available."
        )
        print()
        print("=" * 60)

        input("\nPress ENTER to close...")
        return

    # ========================================================
    # PRE-RPii MIRAKL COMMUNICATION SAFETY CHECK - READ ONLY
    # ========================================================
    #
    # This closes the cache-write-failure gap. If customer
    # communication already exists in Tesco Mirakl but the
    # local cache entry is missing, do NOT enter RPii processing.
    #
    # This performs one READ-ONLY Mirakl thread lookup for the
    # selected uncached order. Any lookup failure fails closed.
    # ========================================================

    print()
    print("=" * 60)
    print("TESCO PRE-RPii COMMUNICATION SAFETY CHECK")
    print("=" * 60)
    print("Reading existing Tesco Mirakl conversation...")

    try:
        mirakl_threads = read_tesco_mirakl_threads(
            selected_order.order_id
        )
    except Exception as exc:
        print()
        print("✗ TESCO MIRAKL THREAD LOOKUP FAILED")
        print(str(exc))
        print()
        print("⚠ REVIEW REQUIRED")
        print()
        print(
            "For safety, RPii processing will NOT begin because "
            "existing customer communication could not be checked."
        )
        print()
        print("NO RPii / SMS / Mirakl write action has been attempted.")
        print("=" * 60)
        input("\nPress ENTER to close...")
        return

    existing_threads = mirakl_threads.get("threads", [])
    existing_messages = mirakl_threads.get("messages", [])

    if existing_threads or existing_messages:
        print()
        print("⚠ EXISTING MIRAKL COMMUNICATION DETECTED")
        print()
        print(f"Order:             {selected_order.order_id}")
        print(f"Existing threads:  {len(existing_threads)}")
        print(f"Existing messages: {len(existing_messages)}")
        print()
        print("⚠ REVIEW REQUIRED")
        print()
        print(
            "This order is not in the verified local cache, but "
            "Tesco Mirakl already contains customer communication."
        )
        print(
            "The invoice / SMS / Tesco message process will NOT run."
        )
        print()
        print("Check this order manually before taking any further action.")
        print("=" * 60)
        input("\nPress ENTER to close...")
        return

    print()
    print("✓ No existing Tesco Mirakl communication found")
    print("✓ CLEAR TO ENTER EXISTING RPii SAFETY PROCESS")
    print("=" * 60)

    inspect_tesco_order_in_rp2(
        selected_order
    )


if __name__ == "__main__":
    main()