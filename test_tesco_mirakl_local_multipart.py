import json
import os
import re
from pathlib import Path
import requests

TESCO_FIXED_SUBJECT = "Electrical Discount UK - Delivery Information & Copy Invoice"
DEFAULT_ORDER_ID = "4431-2696-2301-A"

def section(title):
    print("\n" + "=" * 72)
    print(title)
    print("=" * 72)

def build_local_thread_object(message_body):
    return {
        "topic": {"type": "FREE_TEXT", "value": TESCO_FIXED_SUBJECT},
        "body": message_body,
    }

def build_prepared_request(base_url, order_id, message_body, invoice_path):
    invoice_path = Path(invoice_path)
    if not invoice_path.is_file():
        raise RuntimeError(f"Invoice file does not exist: {invoice_path}")

    url = f"{base_url.rstrip('/')}/api/orders/{order_id}/threads"
    thread_object = build_local_thread_object(message_body)

    # IMPORTANT: 'thread' and 'attachment' are hypotheses only.
    with invoice_path.open("rb") as invoice_file:
        files = {
            "thread": (None, json.dumps(thread_object), "application/json"),
            "attachment": (
                invoice_path.name,
                invoice_file,
                "application/pdf",
            ),
        }
        request = requests.Request(method="POST", url=url, files=files)
        prepared = request.prepare()

    return prepared, thread_object

def extract_boundary(content_type):
    match = re.search(r"boundary=([^;]+)", content_type or "", re.I)
    return match.group(1).strip().strip('"') if match else None

def sanitise_body_preview(body, invoice_name):
    raw = body.encode("utf-8") if isinstance(body, str) else (body or b"")
    text = raw.decode("latin-1", errors="replace")
    marker_pos = text.find(f'filename="{invoice_name}"')
    if marker_pos == -1:
        return text[:8000]
    header_end = text.find("\r\n\r\n", marker_pos)
    if header_end == -1:
        return text[:8000]
    content_start = header_end + 4
    next_boundary = text.find("\r\n--", content_start)
    if next_boundary == -1:
        next_boundary = len(text)
    return (
        text[:content_start]
        + "<PDF BINARY CONTENT REMOVED FROM PREVIEW>"
        + text[next_boundary:]
    )[:12000]

def main():
    section("TESCO LOCAL MULTIPART POST CONSTRUCTOR")
    print("SAFE LOCAL MODE")
    print("This constructs a POST request in memory but NEVER transmits it.")
    print("There is no requests.post(), requests.request(), Session.send(),")
    print("or prepared-request send call anywhere in this script.")

    base_url = input(
        "\nTesco Mirakl base URL [https://tescouk-prod.mirakl.net]: "
    ).strip() or "https://tescouk-prod.mirakl.net"

    order_id = input(
        f"Tesco order ID [{DEFAULT_ORDER_ID}]: "
    ).strip() or DEFAULT_ORDER_ID

    invoice_path = input(
        "Path to an existing test invoice PDF: "
    ).strip().strip('"')

    if not invoice_path:
        raise RuntimeError("An existing invoice PDF path is required.")

    message_body = "LOCAL TEST ONLY - proposed Tesco customer delivery message."

    prepared, thread_object = build_prepared_request(
        base_url, order_id, message_body, invoice_path
    )

    section("LOCAL THREAD OBJECT")
    print(json.dumps(thread_object, indent=2))

    section("PREPARED REQUEST - LOCAL ONLY")
    print(f"Method: {prepared.method}")
    print(f"URL:    {prepared.url}")
    print("\nHeaders:")
    for name, value in prepared.headers.items():
        print(f"  {name}: {value}")

    content_type = prepared.headers.get("Content-Type", "")
    print(f"\nMultipart boundary: {extract_boundary(content_type) or '(not found)'}")

    section("MULTIPART BODY PREVIEW")
    invoice_name = os.path.basename(invoice_path)
    print(sanitise_body_preview(prepared.body, invoice_name))

    section("LOCAL SAFETY CHECKS")
    failures = []
    if prepared.method == "POST":
        print("✓ Prepared method is POST")
    else:
        failures.append("Prepared method is not POST")

    expected_path = f"/api/orders/{order_id}/threads"
    if expected_path in prepared.url:
        print("✓ Prepared URL targets expected order thread endpoint")
    else:
        failures.append("Wrong endpoint")

    if "multipart/form-data" in content_type:
        print("✓ Request is multipart/form-data")
    else:
        failures.append("Not multipart/form-data")

    body = prepared.body.encode() if isinstance(prepared.body, str) else prepared.body
    checks = [
        (b"FREE_TEXT", "FREE_TEXT topic"),
        (TESCO_FIXED_SUBJECT.encode(), "fixed Tesco subject"),
        (invoice_name.encode(), "invoice filename"),
        (b"application/pdf", "PDF content type"),
    ]
    for marker, label in checks:
        if marker in body:
            print(f"✓ {label} is present")
        else:
            failures.append(f"Missing {label}")

    section("RESULT")
    if failures:
        print("✗ LOCAL REQUEST CONSTRUCTION FAILED")
        for failure in failures:
            print(f"- {failure}")
    else:
        print("✓ LOCAL MULTIPART REQUEST CONSTRUCTED SUCCESSFULLY")

    print("\nIMPORTANT:")
    print("The multipart field names 'thread' and 'attachment' are hypotheses.")
    print("This does NOT prove Tesco Mirakl accepts this payload.")
    print("\nNO HTTP REQUEST WAS SENT.")
    print("NO CUSTOMER MESSAGE WAS CREATED.")
    print("NO MIRAKL DATA WAS CHANGED.")

if __name__ == "__main__":
    main()
