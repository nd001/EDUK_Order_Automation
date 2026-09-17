import json
import os
import sys
from datetime import datetime
from pathlib import Path

import requests
from dotenv import load_dotenv
from playwright.sync_api import (
    TimeoutError as PlaywrightTimeoutError,
    sync_playwright,
)

from marketplaces.tesco import (
    read_tesco_orders,
    build_tesco_rp2_search_ref,
)
from config import (
    MODE,
    RP2_URL,
    RP2_REMOTE_USERNAME,
    RP2_REMOTE_PASSWORD,
)
from rp2 import (
    extract_web_ref,
    extract_rp2_order_number,
    parse_delivery_block,
)
from stream_tracking_lookup_MULTI import (
    ensure_order_search_page,
    login_to_stream_if_required,
    read_tracking_from_order_page,
    run_stream_search,
    show_result_summary,
)


# ============================================================
# CONFIG
# ============================================================

load_dotenv()

RPII_USERNAME = os.getenv("RPII_USERNAME")
RPII_PASSWORD = os.getenv("RPII_PASSWORD")

TESCO_MIRAKL_URL = os.getenv(
    "TESCO_MIRAKL_URL",
    "https://tescouk-prod.mirakl.net",
).rstrip("/")

TESCO_MIRAKL_API_KEY = os.getenv(
    "TESCO_MIRAKL_API_KEY",
    "",
).strip()

TESCO_ORDER_FILE = "Tescoorders.xlsx"

TESCO_RP2_CACHE_FILE = (
    Path(__file__).resolve().parent
    / "tesco_rp2_tracking_cache.json"
)

GO2STREAM_CARRIER_CODE = "0000"
GO2STREAM_CARRIER_NAME = "Go2Stream"

GO2STREAM_TRACKING_URL = (
    "https://www.go2stream.net/stream/live/track/view/"
    "Tracking.php?trid={tracking_number}"
)

STREAM_LOGIN_URLS = {
    "SGK": (
        "https://www.go2stream.net/stream/live/cons/view/OrderView.php"
    ),
    "ED": (
        "https://live.go2stream.net/stream/core/view/Login.php"
    ),
}

def get_stream_login_url(courier):
    courier = str(courier or "").strip().upper()
    if courier not in STREAM_LOGIN_URLS:
        raise ValueError("Courier must be SGK or ED.")
    return STREAM_LOGIN_URLS[courier]


# ============================================================
# MIRAKL READ
# ============================================================

def get_tesco_order(order_id):
    """
    Read one Tesco Mirakl order.

    GET only.
    """

    if not TESCO_MIRAKL_API_KEY:
        raise RuntimeError(
            "TESCO_MIRAKL_API_KEY is missing from .env"
        )

    url = f"{TESCO_MIRAKL_URL}/api/orders"

    response = requests.get(
        url,
        headers={
            "Authorization": TESCO_MIRAKL_API_KEY,
            "Accept": "application/json",
        },
        params={
            "order_ids": order_id,
        },
        timeout=30,
    )

    if response.status_code != 200:
        raise RuntimeError(
            "Tesco Mirakl GET failed.\n"
            f"HTTP status: {response.status_code}\n"
            f"Response: {response.text}"
        )

    payload = response.json()
    orders = payload.get("orders", [])

    if not orders:
        raise RuntimeError(
            f"No Tesco Mirakl order found for {order_id}"
        )

    if len(orders) != 1:
        raise RuntimeError(
            f"Expected exactly 1 Tesco order, found {len(orders)}"
        )

    return orders[0]


# ============================================================
# TESCO SHIPPING QUEUE
# ============================================================

def format_ship_by(shipping_deadline):
    if not shipping_deadline:
        return "-"

    try:
        value = str(shipping_deadline).strip()

        if value.endswith("Z"):
            value = value[:-1] + "+00:00"

        return datetime.fromisoformat(
            value
        ).strftime("%d/%m/%Y")

    except (TypeError, ValueError):
        return "REVIEW"


def load_tracking_queue():
    """
    Read the current Tesco spreadsheet and enrich each order
    with its live Tesco Mirakl state.

    READ ONLY.
    """

    spreadsheet_orders = read_tesco_orders(
        TESCO_ORDER_FILE
    )

    queue = []

    for spreadsheet_order in spreadsheet_orders:
        try:
            mirakl_order = get_tesco_order(
                spreadsheet_order.order_id
            )

            state = str(
                mirakl_order.get("order_state") or ""
            ).strip().upper()

            ship_by = format_ship_by(
                mirakl_order.get("shipping_deadline")
            )

            if state == "SHIPPING":
                status = "NEEDS SHIPPING"
            elif state == "SHIPPED":
                status = "COMPLETE"
            else:
                status = "REVIEW"

            queue.append({
                "spreadsheet_order": spreadsheet_order,
                "mirakl_order": mirakl_order,
                "state": state or "-",
                "ship_by": ship_by,
                "status": status,
            })

        except Exception as exc:
            queue.append({
                "spreadsheet_order": spreadsheet_order,
                "mirakl_order": None,
                "state": "ERROR",
                "ship_by": "-",
                "status": "REVIEW",
                "error": str(exc),
            })

    return queue


def print_tracking_queue(queue):
    print()
    print("=" * 112)
    print("TESCO TRACKING / SHIPPING QUEUE")
    print("=" * 112)

    print(
        f"{'':>3} "
        f"{'ORDER':<20} "
        f"{'SKU':<20} "
        f"{'PRICE':>10}   "
        f"{'MIRAKL STATE':<15} "
        f"{'SHIP BY':<12} "
        f"{'STATUS':<16}"
    )

    print("-" * 112)

    for index, item in enumerate(
        queue,
        start=1,
    ):
        order = item["spreadsheet_order"]

        print(
            f"{index:>2}. "
            f"{order.order_id:<20} "
            f"{order.sku:<20} "
            f"£{order.price:>8.2f}   "
            f"{item['state']:<15} "
            f"{item['ship_by']:<12} "
            f"{item['status']:<16}"
        )

    print("=" * 112)
    print()
    print(
        "Only orders marked NEEDS SHIPPING can be selected."
    )


def select_tracking_queue_order(queue):
    selectable = {
        index: item
        for index, item in enumerate(
            queue,
            start=1,
        )
        if item["state"] == "SHIPPING"
    }

    if not selectable:
        print()
        print("✓ No Tesco orders currently need shipping.")
        return None

    while True:
        selection = input(
            "\nSelect order number to ship "
            "(or Q to quit): "
        ).strip()

        if selection.upper() == "Q":
            return None

        try:
            selection_number = int(selection)
        except ValueError:
            print(
                "Please enter a queue number "
                "or Q to quit."
            )
            continue

        if selection_number not in selectable:
            print(
                "That order is not marked NEEDS SHIPPING. "
                "Please select one of the SHIPPING orders."
            )
            continue

        return selectable[selection_number]


# ============================================================
# RPii / STREAM CACHE
# ============================================================

def get_cached_rp2_stream_details(order_id):
    """
    Return validated cached RPii/Stream details for a Tesco
    order, or None when no usable cache entry exists.

    Cache problems never block the RPii fallback.
    """

    order_id = str(order_id or "").strip()

    if not TESCO_RP2_CACHE_FILE.exists():
        print("RPii tracking cache not found.")
        return None

    try:
        with TESCO_RP2_CACHE_FILE.open(
            "r",
            encoding="utf-8",
        ) as file:
            cache = json.load(file)
    except Exception as exc:
        print(
            "⚠ Could not read RPii tracking cache. "
            "RPii fallback will be used."
        )
        print(f"  Reason: {exc}")
        return None

    if not isinstance(cache, dict):
        print(
            "⚠ RPii tracking cache has an unexpected format. "
            "RPii fallback will be used."
        )
        return None

    entry = cache.get(order_id)

    if not isinstance(entry, dict):
        print(
            f"No cache entry found for Tesco order "
            f"{order_id}."
        )
        return None

    rp2_order_number = str(
        entry.get("rp2_order_number") or ""
    ).strip()

    courier = str(
        entry.get("courier") or ""
    ).strip().upper()

    message_verified = (
        entry.get("message_verified") is True
    )

    if (
        not rp2_order_number
        or courier not in {"SGK", "ED"}
        or not message_verified
    ):
        print(
            "⚠ Tesco cache entry is incomplete or unverified. "
            "RPii fallback will be used."
        )
        return None

    print()
    print("✓ RPii / Stream cache hit")
    print(f"  Tesco order:       {order_id}")
    print(f"  RPii/Stream order: {rp2_order_number}")
    print(f"  Stream courier:    {courier}")
    print(
        f"  Cached at:         "
        f"{entry.get('cached_at') or '-'}"
    )

    return {
        "rp2_order_number": rp2_order_number,
        "courier": courier,
    }


def mark_cached_tracking_completed(
    order_id,
    tracking_number,
):
    """
    Mark an existing cache entry complete after Tesco Mirakl
    has been read back and verified as SHIPPED.

    This is auxiliary bookkeeping only. Failure here must not
    invalidate a shipment that has already been verified.
    """

    if not TESCO_RP2_CACHE_FILE.exists():
        return

    try:
        with TESCO_RP2_CACHE_FILE.open(
            "r",
            encoding="utf-8",
        ) as file:
            cache = json.load(file)

        if not isinstance(cache, dict):
            return

        entry = cache.get(str(order_id))

        if not isinstance(entry, dict):
            return

        entry["tracking_completed"] = True
        entry["tracking_completed_at"] = (
            datetime.now().isoformat(
                timespec="seconds"
            )
        )
        entry["tracking_number"] = str(
            tracking_number or ""
        ).strip()

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

        print(
            "✓ RPii tracking cache marked complete."
        )

    except Exception as exc:
        print(
            "⚠ Tesco shipping is verified, but the RPii "
            "tracking cache could not be marked complete."
        )
        print(f"  Reason: {exc}")


# ============================================================
# RPii LOOKUP - READ ONLY
# ============================================================

def ensure_rpii_login(page):
    """
    Automatically log into RPii if its application login screen
    is displayed.
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
        page.locator('input[name="password"]').fill(RPII_PASSWORD)
        page.locator('button[name="doLogin"]').click()

        page.wait_for_load_state("domcontentloaded")

        print("✓ RPii login submitted.")

    else:
        print("✓ RPii login not required.")


def get_rp2_stream_details(order_id):
    """
    READ ONLY.

    Open the selected Tesco sale in RPii, verify the exact Tesco
    Web Ref, then read the internal RPii / Stream order number
    and delivery route.
    """

    rp2_search_ref = build_tesco_rp2_search_ref(order_id)

    print()
    print("=" * 86)
    print("RPii STREAM DETAILS LOOKUP - READ ONLY")
    print("=" * 86)
    print(f"Tesco order: {order_id}")
    print(f"RPii search: {rp2_search_ref}")
    print()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)

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

        try:
            print("Opening RPii...")

            page.goto(
                RP2_URL,
                wait_until="domcontentloaded",
                timeout=30000,
            )

            ensure_rpii_login(page)

            page.wait_for_selector(
                "#SALESFrame",
                state="attached",
                timeout=30000,
            )

            sales_frame = page.frame_locator("#SALESFrame")
            sales_enquiry = sales_frame.locator("#orderno")

            sales_enquiry.wait_for(
                state="visible",
                timeout=60000,
            )

            print(f"Searching RPii for {rp2_search_ref}...")

            sales_enquiry.fill(rp2_search_ref)
            sales_enquiry.press("Enter")

            total_sale_field = sales_frame.locator("#qh_ordval")
            total_sale_field.wait_for(
                state="visible",
                timeout=30000,
            )

            page.wait_for_timeout(1000)

            web_panel = sales_frame.locator(
                "td.SALES-panel",
                has_text="Web Ref:",
            )

            if web_panel.count() != 1:
                raise RuntimeError(
                    "SAFETY STOP\n"
                    "Could not uniquely identify the RPii Web Ref sale header."
                )

            web_panel_text = web_panel.inner_text()
            web_ref = extract_web_ref(web_panel_text)

            if str(web_ref).strip() != str(order_id).strip():
                raise RuntimeError(
                    "SAFETY STOP\n"
                    "RPii Web Ref does not exactly match the selected Tesco order.\n"
                    f"Tesco: {order_id}\n"
                    f"RPii:  {web_ref}"
                )

            print("✓ Exact Tesco Web Ref verified.")

            rp2_order_number = extract_rp2_order_number(
                web_panel_text
            )

            if not rp2_order_number:
                raise RuntimeError(
                    "SAFETY STOP\n"
                    "RPii internal order number could not be extracted "
                    "from the verified sale header."
                )

            delivery_cells = sales_frame.locator("td.tcb")
            delivery_text = None

            for i in range(delivery_cells.count()):
                cell_text = delivery_cells.nth(i).inner_text()

                if (
                    "DELIVER" in cell_text
                    or "ON MANFST" in cell_text
                    or "DELVRD" in cell_text
                ):
                    delivery_text = cell_text
                    break

            if not delivery_text:
                raise RuntimeError(
                    "SAFETY STOP\n"
                    "RPii delivery information could not be found "
                    "on the verified Tesco sale."
                )

            delivery = parse_delivery_block(delivery_text)

            rp2_courier = str(
                delivery.get("courier") or ""
            ).strip().upper()

            if rp2_courier == "SK":
                stream_courier = "SGK"
            elif rp2_courier == "ED":
                stream_courier = "ED"
            else:
                raise RuntimeError(
                    "SAFETY STOP\n"
                    "RPii delivery route is not recognised.\n"
                    f"Courier: {rp2_courier!r}"
                )

            print()
            print("✓ RPii Stream details verified")
            print(f"  RPii order number: {rp2_order_number}")
            print(f"  RPii courier:      {rp2_courier}")
            print(f"  Stream courier:    {stream_courier}")
            print()

            return {
                "rp2_order_number": str(rp2_order_number).strip(),
                "courier": stream_courier,
            }

        finally:
            browser.close()


# ============================================================
# STREAM TRACKING LOOKUP
# ============================================================

def get_mirakl_postcode(order):
    customer = order.get("customer") or {}
    shipping_address = (
        customer.get("shipping_address")
        or {}
    )

    return str(
        shipping_address.get("zip_code")
        or ""
    ).strip()


def fetch_stream_tracking(
    stream_order_number,
    postcode,
    courier,
):
    """
    READ ONLY.

    Log into Stream, locate the exact Stream order using
    RPii/Stream order number + postcode, and read the visible
    tracking number if it is available.

    Returns:
        tracking number string when available
        None when the Stream order/tracking is not yet available
    """

    print()
    print("=" * 86)
    print("STREAM TRACKING LOOKUP - READ ONLY")
    print("=" * 86)
    print(f"Courier:             {courier}")
    print(f"RPii / Stream order: {stream_order_number}")
    print(f"Postcode:            {postcode}")
    print()

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False
        )

        context = browser.new_context()
        page = context.new_page()

        try:
            print("Opening Stream...")

            stream_login_url = get_stream_login_url(
                courier
            )

            print(
                f"Using Stream ({courier}) URL: "
                f"{stream_login_url}"
            )

            page.goto(
                stream_login_url,
                wait_until="domcontentloaded",
            )

            login_to_stream_if_required(
                page,
                courier=courier,
            )

            page.wait_for_timeout(1000)

            ensure_order_search_page(
                page,
                courier=courier,
            )

            match, status_label = run_stream_search(
                page,
                stream_order_number,
                postcode,
                courier=courier,
            )

            if match is None:
                print()
                print("=" * 86)
                print("STREAM ORDER NOT FOUND")
                print("=" * 86)
                print(
                    "No exact Stream order matched the "
                    "RPii/Stream order number and postcode."
                )
                print(
                    "No Tesco/Mirakl changes have been made."
                )
                return None

            show_result_summary(match)

            print()
            print("Opening matched Stream order...")

            edit_button = match.locator(
                ".editOrderIcon:visible, "
                ".editOrderIconMobile:visible"
            )

            if edit_button.count() == 0:
                raise RuntimeError(
                    "Matched Stream order was found, but no "
                    "visible Edit this order button was available."
                )

            edit_button.first.click()
            page.wait_for_timeout(1000)

            try:
                tracking_number = (
                    read_tracking_from_order_page(page)
                )
            except PlaywrightTimeoutError:
                print()
                print("=" * 86)
                print("TRACKING PENDING")
                print("=" * 86)
                print(
                    "The Stream order exists, but no tracking "
                    "number is available yet."
                )
                print(
                    "This is expected for newly processed orders."
                )
                print(
                    "No Tesco/Mirakl changes have been made."
                )
                return None

            tracking_number = validate_tracking_number(
                tracking_number
            )

            print()
            print("=" * 86)
            print("STREAM TRACKING FOUND")
            print("=" * 86)
            print(f"Search status:   {status_label}")
            print(f"RPii order:      {stream_order_number}")
            print(f"Postcode:        {postcode}")
            print(f"Tracking number: {tracking_number}")
            print("=" * 86)

            return tracking_number

        finally:
            browser.close()


# ============================================================
# MIRAKL WRITE - TRACKING
# ============================================================

def update_tesco_tracking(
    order_id,
    tracking_number,
    tracking_url,
):
    """
    Write Go2Stream tracking details to Tesco Mirakl.

    ONE PUT ATTEMPT ONLY.
    No automatic retry.
    """

    url = (
        f"{TESCO_MIRAKL_URL}"
        f"/api/orders/{order_id}/tracking"
    )

    payload = {
        "carrier_code": GO2STREAM_CARRIER_CODE,
        "carrier_name": GO2STREAM_CARRIER_NAME,
        "carrier_url": tracking_url,
        "tracking_number": tracking_number,
    }

    response = requests.put(
        url,
        headers={
            "Authorization": TESCO_MIRAKL_API_KEY,
            "Accept": "application/json",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=30,
    )

    if not 200 <= response.status_code < 300:
        raise RuntimeError(
            "Tesco tracking update failed.\n"
            "DO NOT automatically retry.\n"
            f"HTTP status: {response.status_code}\n"
            f"Response: {response.text}"
        )

    return response.status_code


# ============================================================
# MIRAKL WRITE - SHIP
# ============================================================

def ship_tesco_order(order_id):
    """
    Validate shipment in Tesco Mirakl.

    ONE PUT ATTEMPT ONLY.
    No automatic retry.
    """

    url = (
        f"{TESCO_MIRAKL_URL}"
        f"/api/orders/{order_id}/ship"
    )

    response = requests.put(
        url,
        headers={
            "Authorization": TESCO_MIRAKL_API_KEY,
            "Accept": "application/json",
        },
        timeout=30,
    )

    if not 200 <= response.status_code < 300:
        raise RuntimeError(
            "Tesco ship action failed.\n"
            "DO NOT automatically retry.\n"
            f"HTTP status: {response.status_code}\n"
            f"Response: {response.text}"
        )

    return response.status_code


# ============================================================
# HELPERS
# ============================================================

def get_first_order_line(order):
    lines = order.get("order_lines") or []

    if not lines:
        return {}

    return lines[0]


def show_order_summary(order):
    customer = order.get("customer") or {}
    shipping_address = (
        customer.get("shipping_address")
        or {}
    )
    line = get_first_order_line(order)

    customer_name = " ".join(
        part
        for part in (
            shipping_address.get("firstname"),
            shipping_address.get("lastname"),
        )
        if part
    ).strip()

    sku = (
        line.get("offer_sku")
        or line.get("product_shop_sku")
        or ""
    )

    postcode = shipping_address.get("zip_code") or ""

    price = order.get("total_price")

    print(f"Order ID:       {order.get('order_id')}")
    print(f"Order state:    {order.get('order_state')}")
    print(f"Customer:       {customer_name}")
    print(f"SKU:            {sku}")
    print(f"Postcode:       {postcode}")

    if price is None:
        print("Total price:    ")
    else:
        print(f"Total price:    £{float(price):.2f}")

    print(
        "Current carrier: "
        f"{order.get('shipping_company') or ''}"
    )
    print(
        "Current tracking: "
        f"{order.get('shipping_tracking') or ''}"
    )
    print(
        "Current URL:      "
        f"{order.get('shipping_tracking_url') or ''}"
    )


def validate_tracking_number(tracking_number):
    tracking_number = tracking_number.strip()

    if not tracking_number:
        raise ValueError(
            "Tracking number cannot be blank."
        )

    if any(char.isspace() for char in tracking_number):
        raise ValueError(
            "Tracking number must not contain spaces."
        )

    if "/" in tracking_number or ":" in tracking_number:
        raise ValueError(
            "Enter the Go2Stream tracking number only, "
            "not a tracking URL."
        )

    return tracking_number


def verify_tracking_readback(
    order,
    tracking_number,
    tracking_url,
):
    errors = []

    if order.get("shipping_company") != GO2STREAM_CARRIER_NAME:
        errors.append(
            "Carrier name did not read back as Go2Stream."
        )

    if order.get("shipping_carrier_code") != GO2STREAM_CARRIER_CODE:
        errors.append(
            "Carrier code did not read back as 0000."
        )

    if order.get("shipping_tracking") != tracking_number:
        errors.append(
            "Tracking number did not read back correctly."
        )

    if order.get("shipping_tracking_url") != tracking_url:
        errors.append(
            "Tracking URL did not read back correctly."
        )

    if errors:
        raise RuntimeError(
            "Tracking write could not be fully verified:\n"
            + "\n".join(
                f" - {error}"
                for error in errors
            )
        )


def verify_shipped_readback(
    order,
    tracking_number,
    tracking_url,
):
    errors = []

    if order.get("order_state") != "SHIPPED":
        errors.append(
            f"Order state is {order.get('order_state')!r}, "
            "not 'SHIPPED'."
        )

    lines = order.get("order_lines") or []

    if not lines:
        errors.append(
            "No order lines were returned."
        )
    else:
        bad_states = [
            line.get("order_line_state")
            for line in lines
            if line.get("order_line_state") != "SHIPPED"
        ]

        if bad_states:
            errors.append(
                "One or more order lines are not SHIPPED: "
                f"{bad_states}"
            )

    if order.get("shipping_company") != GO2STREAM_CARRIER_NAME:
        errors.append(
            "Carrier name is not Go2Stream."
        )

    if order.get("shipping_carrier_code") != GO2STREAM_CARRIER_CODE:
        errors.append(
            "Carrier code is not 0000."
        )

    if order.get("shipping_tracking") != tracking_number:
        errors.append(
            "Tracking number does not match."
        )

    if order.get("shipping_tracking_url") != tracking_url:
        errors.append(
            "Tracking URL does not match."
        )

    if errors:
        raise RuntimeError(
            "Final Tesco shipping verification failed:\n"
            + "\n".join(
                f" - {error}"
                for error in errors
            )
        )


# ============================================================
# MAIN
# ============================================================

def main():
    print()
    print("=" * 86)
    print("TESCO STREAM TRACKING / SHIPPING COMPLETION")
    print("=" * 86)
    print()
    print(
        "This tool can WRITE tracking details and mark "
        "a Tesco Mirakl order SHIPPED."
    )
    print(
        "No write occurs until the exact SHIP confirmation "
        "is entered."
    )
    print()

    print(
        f"Reading Tesco order queue from "
        f"{TESCO_ORDER_FILE}..."
    )

    queue = load_tracking_queue()

    print_tracking_queue(queue)

    selected = select_tracking_queue_order(
        queue
    )

    if selected is None:
        print()
        print("No order selected.")
        return

    spreadsheet_order = selected[
        "spreadsheet_order"
    ]
    order = selected["mirakl_order"]
    order_id = spreadsheet_order.order_id

    print()
    print(f"Selected Tesco order: {order_id}")
    print()

    if order.get("order_id") != order_id:
        raise RuntimeError(
            "ORDER ID SAFETY FAILURE\n"
            f"Requested: {order_id}\n"
            f"Returned:  {order.get('order_id')}"
        )

    print("=" * 86)
    print("TESCO ORDER")
    print("=" * 86)
    show_order_summary(order)
    print()

    current_state = order.get("order_state")

    if current_state == "SHIPPED":
        print("✓ This Tesco order is already SHIPPED.")
        print("No write action will be performed.")
        return

    if current_state != "SHIPPING":
        raise RuntimeError(
            "SAFETY STOP\n"
            "Expected Tesco order state SHIPPING before "
            f"completion, but received {current_state!r}."
        )

    existing_tracking = (
        order.get("shipping_tracking")
        or ""
    ).strip()

    existing_company = (
        order.get("shipping_company")
        or ""
    ).strip()

    if existing_tracking or existing_company:
        raise RuntimeError(
            "SAFETY STOP\n"
            "This SHIPPING order already contains carrier "
            "or tracking information.\n"
            f"Carrier:  {existing_company!r}\n"
            f"Tracking: {existing_tracking!r}\n"
            "Review it manually before making any change."
        )

    postcode = get_mirakl_postcode(order)

    if not postcode:
        raise RuntimeError(
            "SAFETY STOP\n"
            "Tesco Mirakl did not return a customer postcode. "
            "Stream lookup cannot continue safely."
        )

    print()
    print(
        "Looking for cached RPii order details..."
    )

    rp2_stream_details = (
        get_cached_rp2_stream_details(
            order_id
        )
    )

    if rp2_stream_details is None:
        print()
        print(
            "Cache miss - reading RPii order number "
            "and delivery route..."
        )

        rp2_stream_details = get_rp2_stream_details(
            order_id
        )

        details_source = "RPii fallback"
    else:
        details_source = "JSON cache"

    stream_order_number = (
        rp2_stream_details["rp2_order_number"]
    )

    courier = (
        rp2_stream_details["courier"]
    )

    print()
    print("=" * 86)
    print("RPii / STREAM DETAILS")
    print("=" * 86)
    print(f"Source:              {details_source}")
    print(f"RPii / Stream order: {stream_order_number}")
    print(f"Stream courier:      {courier}")
    print(f"Postcode:            {postcode}")
    print("=" * 86)

    tracking_number = fetch_stream_tracking(
        stream_order_number=stream_order_number,
        postcode=postcode,
        courier=courier,
    )

    if not tracking_number:
        print()
        print(
            "Tracking is not ready for this order. "
            "Run tesco_tracking again later."
        )
        print(
            "NO TESCO/MIRAKL WRITE ACTION HAS BEEN PERFORMED."
        )
        return

    tracking_url = GO2STREAM_TRACKING_URL.format(
        tracking_number=tracking_number
    )

    print()
    print("=" * 86)
    print("DRY RUN - NO WRITE YET")
    print("=" * 86)
    print(f"Order ID:         {order_id}")
    print(f"Carrier code:     {GO2STREAM_CARRIER_CODE}")
    print(f"Carrier name:     {GO2STREAM_CARRIER_NAME}")
    print(f"Tracking number:  {tracking_number}")
    print(f"Tracking URL:     {tracking_url}")
    print()
    print("Planned Mirakl actions:")
    print("  1. PUT tracking details")
    print("  2. GET order and verify tracking")
    print("  3. PUT ship")
    print("  4. GET order and verify SHIPPED")
    print()
    print(
        "No automatic write retry will be attempted "
        "if a request is ambiguous or fails."
    )
    print()

    expected_confirmation = f"SHIP {order_id}"

    confirmation = input(
        f"Type exactly '{expected_confirmation}' to continue: "
    ).strip()

    if confirmation != expected_confirmation:
        print()
        print("Confirmation did not match.")
        print("NO WRITE ACTION HAS BEEN PERFORMED.")
        return

    print()
    print("✓ LIVE SHIPPING AUTHORISATION ACCEPTED")
    print()

    # --------------------------------------------------------
    # WRITE 1 - TRACKING
    # --------------------------------------------------------

    print("=" * 86)
    print("STEP 1 - WRITE GO2STREAM TRACKING")
    print("=" * 86)
    print("One PUT attempt only.")
    print()

    tracking_status = update_tesco_tracking(
        order_id=order_id,
        tracking_number=tracking_number,
        tracking_url=tracking_url,
    )

    print(
        "✓ Tesco Mirakl tracking PUT returned success "
        f"(HTTP {tracking_status})"
    )
    print()

    # --------------------------------------------------------
    # READBACK 1 - VERIFY TRACKING
    # --------------------------------------------------------

    print("Reading order back to verify tracking...")
    tracking_order = get_tesco_order(order_id)

    verify_tracking_readback(
        order=tracking_order,
        tracking_number=tracking_number,
        tracking_url=tracking_url,
    )

    print("✓ Go2Stream tracking readback verified")
    print(f"  Carrier:  {tracking_order.get('shipping_company')}")
    print(f"  Code:     {tracking_order.get('shipping_carrier_code')}")
    print(f"  Tracking: {tracking_order.get('shipping_tracking')}")
    print(f"  URL:      {tracking_order.get('shipping_tracking_url')}")
    print()

    # --------------------------------------------------------
    # WRITE 2 - SHIP
    # --------------------------------------------------------

    print("=" * 86)
    print("STEP 2 - MARK TESCO ORDER SHIPPED")
    print("=" * 86)
    print("One PUT attempt only.")
    print()

    ship_status = ship_tesco_order(order_id)

    print(
        "✓ Tesco Mirakl ship PUT returned success "
        f"(HTTP {ship_status})"
    )
    print()

    # --------------------------------------------------------
    # READBACK 2 - VERIFY FINAL STATE
    # --------------------------------------------------------

    print("Reading order back to verify final state...")
    final_order = get_tesco_order(order_id)

    verify_shipped_readback(
        order=final_order,
        tracking_number=tracking_number,
        tracking_url=tracking_url,
    )

    print()
    print("=" * 86)
    print("✓ TESCO SHIPPING COMPLETION VERIFIED")
    print("=" * 86)
    print(f"Order ID:         {order_id}")
    print(f"Order state:      {final_order.get('order_state')}")
    print(f"Carrier code:     {final_order.get('shipping_carrier_code')}")
    print(f"Carrier name:     {final_order.get('shipping_company')}")
    print(f"Tracking number:  {final_order.get('shipping_tracking')}")
    print(f"Tracking URL:     {final_order.get('shipping_tracking_url')}")
    print()
    print("Tesco order is now verified as SHIPPED.")

    mark_cached_tracking_completed(
        order_id=order_id,
        tracking_number=tracking_number,
    )


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print()
        print("Cancelled by user.")
        sys.exit(1)
    except Exception as exc:
        print()
        print("=" * 86)
        print("SAFETY STOP")
        print("=" * 86)
        print(str(exc))
        print()
        print(
            "If a write action may already have been attempted, "
            "DO NOT simply run the script again."
        )
        print(
            "Read the Tesco Mirakl order first and confirm its "
            "current state/tracking before retrying."
        )
        sys.exit(1)
