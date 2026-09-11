import os
import re
import sys

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

STREAM_ORDER_SEARCH_URL = (
    "https://www.go2stream.net/stream/live/cons/view/OrderView.php"
)

STATUS_SEQUENCE = [
    ("UNCONFIRMED", "Unconfirmed"),
    ("UNPLANNED", "Unplanned"),
    ("", "All Active"),
]


load_dotenv()


def get_stream_credentials(courier):
    courier = str(courier or "").strip().upper()

    if courier not in {"SGK", "ED"}:
        raise ValueError(
            "Courier must be SGK or ED."
        )

    if courier == "SGK":
        username = (
            os.getenv("STREAM_SGK_USERNAME", "").strip()
            or os.getenv("STREAM_USERNAME", "").strip()
        )
        password = (
            os.getenv("STREAM_SGK_PASSWORD", "")
            or os.getenv("STREAM_PASSWORD", "")
        )
    else:
        username = os.getenv(
            "STREAM_ED_USERNAME",
            "",
        ).strip()
        password = os.getenv(
            "STREAM_ED_PASSWORD",
            "",
        )

    if not username or not password:
        raise RuntimeError(
            f"Stream {courier} login credentials are missing "
            "from .env."
        )

    return username, password


def login_to_stream_if_required(
    page,
    courier="SGK",
):
    """
    Log into the requested Stream account if the login page is shown.

    READ ONLY.

    courier:
        SGK -> STREAM_SGK_USERNAME / STREAM_SGK_PASSWORD
        ED  -> STREAM_ED_USERNAME  / STREAM_ED_PASSWORD

    For backward compatibility, SGK can also use the older
    STREAM_USERNAME / STREAM_PASSWORD names.
    """

    user_id = page.locator("#userId:visible")

    if user_id.count() == 0:
        return

    username, password = get_stream_credentials(
        courier
    )

    print(
        f"Logging into Stream ({courier})..."
    )

    user_id.first.fill(username)

    password_field = page.locator(
        "#password:visible"
    )

    if password_field.count() == 0:
        raise RuntimeError(
            "Stream login page is visible, but the password "
            "field could not be found."
        )

    password_field.first.fill(password)

    login_button = page.locator("#login:visible")

    if login_button.count() == 0:
        raise RuntimeError(
            "Stream login page is visible, but the Sign In "
            "button could not be found."
        )

    login_button.first.click()

    try:
        page.locator(
            "#userId:visible"
        ).wait_for(
            state="detached",
            timeout=15000,
        )
    except PlaywrightTimeoutError:
        error_text = ""

        for selector in (
            "#userIdErrors:visible",
            "#passwordErrors:visible",
        ):
            locator = page.locator(selector)

            if locator.count() > 0:
                candidate = (
                    locator.first.inner_text()
                    or ""
                ).strip()

                if candidate:
                    error_text = candidate
                    break

        if error_text:
            raise RuntimeError(
                f"Stream {courier} login failed: {error_text}"
            )

        raise RuntimeError(
            f"Stream {courier} login did not complete "
            "within 15 seconds."
        )

    print(
        f"✓ Stream {courier} login successful."
    )


def normalise_postcode(value):
    if not value:
        return ""
    return re.sub(r"\s+", "", str(value).upper().strip())


def extract_label_value(row, label_text):
    rows = row.locator(".row")

    for i in range(rows.count()):
        candidate = rows.nth(i)
        label = candidate.locator("label")

        if label.count() == 0:
            continue

        current_label = " ".join(
            label.first.inner_text().split()
        )

        if current_label.rstrip(":").lower() != label_text.rstrip(":").lower():
            continue

        columns = candidate.locator(
            ".col-xs-7, .col-sm-7, .col-md-7, .col-lg-7"
        )

        if columns.count() == 0:
            continue

        value = " ".join(
            columns.first.inner_text().split()
        )

        return value.strip()

    return ""


def find_matching_result(page, order_number, postcode):
    """
    Match against the visible Stream result row.

    The desktop and responsive Stream result layouts are different,
    so do not depend on the mobile <label> structure here. Instead,
    require BOTH the exact Stream/RPii order number and postcode to
    be present in the same visible row.
    """
    result_rows = page.locator("div.selectGridRow:visible")

    matches = []
    expected_order = str(order_number).strip()
    expected_postcode = normalise_postcode(postcode)

    for i in range(result_rows.count()):
        row = result_rows.nth(i)

        row_text = " ".join(
            row.inner_text().split()
        )

        row_text_upper = row_text.upper()
        row_text_no_spaces = re.sub(
            r"\s+",
            "",
            row_text_upper,
        )

        order_matches = (
            expected_order.upper()
            in row_text_upper
        )

        postcode_matches = (
            expected_postcode
            in row_text_no_spaces
        )

        if order_matches and postcode_matches:
            matches.append(row)

    if len(matches) == 0:
        return None

    if len(matches) > 1:
        raise RuntimeError(
            "SAFETY STOP\n"
            f"Found {len(matches)} visible Stream rows matching "
            f"order {expected_order} and postcode {postcode}.\n"
            "Expected exactly one result."
        )

    return matches[0]


def ensure_order_search_page(page):
    """
    Make sure Stream Order Search is loaded without unnecessarily
    navigating when Stream has just redirected there after login.
    """
    if "OrderView.php" in page.url:
        page.wait_for_load_state("domcontentloaded")
        return

    page.goto(
        STREAM_ORDER_SEARCH_URL,
        wait_until="domcontentloaded",
    )


def open_stream_search_form(page):
    """
    Open Stream's visible search modal and return it.

    If Stream is on the "No Results Found" screen, first use its
    Search Again button to clear the date-filter problem, then reopen
    the normal Search dialog.
    """
    search_again = page.locator("#searchAgain:visible")

    if search_again.count() > 0:
        print(
            "Stream returned 'No Results Found' - "
            "using Search Again."
        )
        search_again.first.click()
        page.wait_for_timeout(1000)

    open_button = page.locator(
        "#openSearchModalButton:visible"
    )

    open_button.wait_for(
        state="visible",
        timeout=10000,
    )
    open_button.click()

    search_modal = page.locator(
        "#searchFilterModal:visible"
    )

    search_modal.wait_for(
        state="visible",
        timeout=10000,
    )

    search_modal.locator(
        "#orderNumber"
    ).wait_for(
        state="visible",
        timeout=10000,
    )

    return search_modal


def run_stream_search(page, order_number, postcode):
    for status_value, status_label in STATUS_SEQUENCE:
        print()
        print(f"Searching Stream status: {status_label}")

        ensure_order_search_page(page)
        search_modal = open_stream_search_form(page)

        search_modal.locator(
            "#orderNumber"
        ).fill(
            str(order_number).strip()
        )

        search_modal.locator(
            "#orderStatus"
        ).select_option(
            status_value
        )

        activate_search = search_modal.get_by_role(
            "button",
            name="Activate Search",
        )

        activate_search.wait_for(
            state="visible",
            timeout=10000,
        )

        activate_search.click()

        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(1000)

        search_again = page.locator("#searchAgain:visible")
        if search_again.count() > 0:
            print(
                f"No results returned under {status_label}."
            )
            continue

        match = find_matching_result(
            page,
            order_number,
            postcode,
        )

        if match is not None:
            print(
                f"✓ Matching Stream order found under {status_label}"
            )
            return match, status_label

        print(
            f"No exact match found under {status_label}"
        )

    return None, None


def show_result_summary(row):
    customer = extract_label_value(row, "Customer")
    order_no = extract_label_value(row, "Order No.")
    postcode = extract_label_value(row, "Postal Code")
    status = extract_label_value(row, "Status")

    stream_key = ""
    key_locator = row.locator(".gridKey")

    if key_locator.count() > 0:
        stream_key = " ".join(
            key_locator.first.inner_text().split()
        )

    row_text = " ".join(
        row.inner_text().split()
    )

    print()
    print("=" * 80)
    print("MATCHED STREAM RESULT")
    print("=" * 80)

    if any([customer, order_no, postcode, status]):
        print(f"Customer:     {customer}")
        print(f"Order No:     {order_no}")
        print(f"Postal Code:  {postcode}")
        print(f"Status:       {status}")
    else:
        print(f"Matched row:  {row_text}")

    print(f"Stream key:   {stream_key}")
    print("=" * 80)


def read_tracking_from_order_page(page):
    tracking = page.locator("#deliverTo-tracking-id")

    tracking.wait_for(
        state="visible",
        timeout=15000,
    )

    tracking_number = " ".join(
        tracking.inner_text().split()
    ).strip()

    if not tracking_number:
        raise RuntimeError(
            "Stream order page opened, but the tracking "
            "number was blank."
        )

    return tracking_number


def main():
    print()
    print("=" * 80)
    print("STREAM TRACKING LOOKUP - READ ONLY")
    print("=" * 80)
    print()
    print(
        "This script searches Stream and reads an existing tracking number."
    )
    print(
        "It does not change any Stream or Tesco data."
    )
    print()

    order_number = input(
        "Enter RPii / Stream order number: "
    ).strip()

    if not order_number:
        print("No order number entered.")
        sys.exit(1)

    postcode = input(
        "Enter customer postcode: "
    ).strip()

    if not postcode:
        print("No postcode entered.")
        sys.exit(1)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        print()
        print("Opening Stream...")
        print()

        page.goto(
            STREAM_ORDER_SEARCH_URL,
            wait_until="domcontentloaded",
        )

        login_to_stream_if_required(page)

        # After login Stream may redirect elsewhere before settling.
        page.wait_for_timeout(1000)

        ensure_order_search_page(page)

        match, status_label = run_stream_search(
            page,
            order_number,
            postcode,
        )

        if match is None:
            print()
            print("=" * 80)
            print("NO EXACT STREAM MATCH FOUND")
            print("=" * 80)
            print(f"Order number: {order_number}")
            print(f"Postcode:     {postcode}")
            print()
            print(
                "Searched: Unconfirmed, Unplanned, and All Active."
            )
            print("No order was opened.")
            input(
                "\nPress Enter to close the browser..."
            )
            return

        show_result_summary(match)

        print()
        print("Opening matched Stream order...")

        edit_button = match.locator(
            ".editOrderIcon:visible, .editOrderIconMobile:visible"
        )

        if edit_button.count() == 0:
            raise RuntimeError(
                "Matched Stream order was found, but no visible "
                "Edit this order button was available."
            )

        edit_button.first.click()
        page.wait_for_timeout(1000)

        tracking_number = read_tracking_from_order_page(page)

        print()
        print("=" * 80)
        print("STREAM TRACKING FOUND")
        print("=" * 80)
        print(f"Search status:   {status_label}")
        print(f"Order number:    {order_number}")
        print(f"Postcode:        {postcode}")
        print(f"Tracking number: {tracking_number}")
        print("=" * 80)
        print()
        print(
            "READ ONLY - NO STREAM OR TESCO WRITE ACTION HAS BEEN PERFORMED."
        )

        input(
            "\nPress Enter to close the browser..."
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
        print("=" * 80)
        print("SAFETY STOP")
        print("=" * 80)
        print(str(exc))
        print()
        print("No automatic retry was attempted.")
        sys.exit(1)
