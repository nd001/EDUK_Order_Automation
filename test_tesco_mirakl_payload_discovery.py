"""
Tesco Mirakl POST payload discovery - SAFE / READ ONLY.

This script is deliberately incapable of sending POST, PUT, PATCH or DELETE
requests. It uses GET, HEAD and OPTIONS only.

Purpose:
- Inspect the raw structure of existing Tesco Mirakl threads/messages.
- Inspect GET /api/orders/{order_id}/messages.
- Re-check messaging reason codes.
- Inspect OPTIONS metadata for known write endpoints.
- Probe common read-only OpenAPI/Swagger locations.
- If an API schema is found, extract any schema fragments for the known
  messaging endpoints.
- Build LOCAL candidate payload shapes for discussion only.

NO CUSTOMER MESSAGE CAN BE CREATED BY THIS SCRIPT.
"""

import json
from pprint import pprint
import requests

from config import (
    TESCO_MIRAKL_URL,
    TESCO_MIRAKL_API_KEY,
)


DEFAULT_ORDER_ID = "4431-2696-2301-A"

SAFE_METHODS = {"GET", "HEAD", "OPTIONS"}

KNOWN_MESSAGE_PATHS = [
    "/api/orders/{order_id}/threads",
    "/api/orders/{order_id}/messages",
    "/api/inbox/threads/{thread_id}/message",
]


def safe_request(method, url, **kwargs):
    method = method.upper()

    if method not in SAFE_METHODS:
        raise RuntimeError(
            f"BLOCKED UNSAFE HTTP METHOD: {method}. "
            "This discovery script permits GET, HEAD and OPTIONS only."
        )

    return requests.request(
        method,
        url,
        timeout=30,
        **kwargs,
    )


def auth_headers():
    if not TESCO_MIRAKL_API_KEY:
        raise RuntimeError(
            "TESCO_MIRAKL_API_KEY is missing."
        )

    return {
        "Authorization": TESCO_MIRAKL_API_KEY,
        "Accept": "application/json",
    }


def section(title):
    print()
    print("=" * 70)
    print(title)
    print("=" * 70)


def print_response_summary(response, body_limit=1500):
    print(f"HTTP:    {response.status_code}")
    print(f"Allow:   {response.headers.get('Allow', '(not supplied)')}")
    print(
        "Type:    "
        f"{response.headers.get('Content-Type', '(not supplied)')}"
    )

    text = response.text.strip()

    if not text:
        print("Body:    (empty)")
        return

    if len(text) > body_limit:
        text = text[:body_limit] + "\n... [truncated]"

    print("Body:")
    print(text)


def read_threads(order_id):
    url = (
        f"{TESCO_MIRAKL_URL.rstrip('/')}"
        "/api/inbox/threads"
    )

    params = {
        "entity_type": "MMP_ORDER",
        "entity_id": order_id,
        "with_messages": "true",
    }

    response = safe_request(
        "GET",
        url,
        headers=auth_headers(),
        params=params,
    )

    response.raise_for_status()
    return response.json()


def find_first_thread_id(data):
    threads = data.get("data", [])

    if not threads:
        return None

    return threads[0].get("id")


def print_thread_structure(data):
    threads = data.get("data", [])

    print(f"Threads found: {len(threads)}")

    if not threads:
        return

    for index, thread in enumerate(threads, start=1):
        print()
        print(f"THREAD {index}")
        print("-" * 70)

        print("Top-level thread keys:")
        print(sorted(thread.keys()))

        topic = thread.get("topic")
        print()
        print("Topic object:")
        pprint(topic)

        messages = thread.get("messages", [])
        print()
        print(f"Messages: {len(messages)}")

        for message_index, message in enumerate(
            messages,
            start=1,
        ):
            print()
            print(f"Message {message_index} keys:")
            print(sorted(message.keys()))

            metadata = {
                key: value
                for key, value in message.items()
                if key != "body"
            }

            pprint(metadata)

            body = message.get("body")
            if body:
                print("Body preview:")
                print(str(body)[:800])

        print()
        print("Raw thread JSON:")
        print(
            json.dumps(
                thread,
                indent=2,
                default=str,
            )[:10000]
        )


def inspect_order_messages(order_id):
    section(
        "GET /api/orders/{order_id}/messages - READ ONLY"
    )

    url = (
        f"{TESCO_MIRAKL_URL.rstrip('/')}"
        f"/api/orders/{order_id}/messages"
    )

    response = safe_request(
        "GET",
        url,
        headers=auth_headers(),
    )

    print_response_summary(
        response,
        body_limit=5000,
    )

    if response.status_code == 200:
        try:
            data = response.json()

            print()
            print("Parsed JSON structure:")
            if isinstance(data, dict):
                print("Top-level keys:", sorted(data.keys()))
            elif isinstance(data, list):
                print(f"Top-level list length: {len(data)}")

            pprint(data)
        except ValueError:
            pass


def inspect_reasons():
    section("GET /api/reasons - MESSAGING REASONS")

    url = (
        f"{TESCO_MIRAKL_URL.rstrip('/')}"
        "/api/reasons"
    )

    response = safe_request(
        "GET",
        url,
        headers=auth_headers(),
    )

    response.raise_for_status()
    data = response.json()

    reasons = data.get("reasons", [])

    messaging = [
        reason
        for reason in reasons
        if "MESSAGING" in str(
            reason.get("type", "")
        )
    ]

    for reason in messaging:
        print(
            json.dumps(
                reason,
                indent=2,
            )
        )

    return messaging


def inspect_options(order_id, thread_id):
    paths = [
        (
            "/api/orders/{order_id}/threads",
            f"/api/orders/{order_id}/threads",
        ),
        (
            "/api/orders/{order_id}/messages",
            f"/api/orders/{order_id}/messages",
        ),
    ]

    if thread_id:
        paths.append(
            (
                "/api/inbox/threads/{thread_id}/message",
                f"/api/inbox/threads/{thread_id}/message",
            )
        )

    for label, concrete_path in paths:
        section(f"OPTIONS {label}")

        url = (
            f"{TESCO_MIRAKL_URL.rstrip('/')}"
            f"{concrete_path}"
        )

        response = safe_request(
            "OPTIONS",
            url,
            headers=auth_headers(),
        )

        print_response_summary(
            response,
            body_limit=3000,
        )


def recursively_find_path_fragments(obj, wanted_paths):
    found = {}

    def walk(value):
        if isinstance(value, dict):
            for key, child in value.items():
                if key in wanted_paths:
                    found[key] = child
                walk(child)
        elif isinstance(value, list):
            for child in value:
                walk(child)

    walk(obj)
    return found


def probe_api_schemas():
    section("READ-ONLY OPENAPI / SWAGGER SCHEMA PROBE")

    candidates = [
        "/openapi.json",
        "/api/openapi.json",
        "/swagger.json",
        "/api/swagger.json",
        "/v3/api-docs",
        "/api-docs",
        "/swagger/v1/swagger.json",
    ]

    schema_hits = {}

    for path in candidates:
        url = (
            f"{TESCO_MIRAKL_URL.rstrip('/')}"
            f"{path}"
        )

        print()
        print(f"GET {path}")

        try:
            response = safe_request(
                "GET",
                url,
                headers=auth_headers(),
            )
        except requests.RequestException as exc:
            print(f"Request failed: {exc}")
            continue

        print(
            f"HTTP {response.status_code} | "
            f"{response.headers.get('Content-Type', '(no type)')}"
        )

        if response.status_code != 200:
            continue

        try:
            data = response.json()
        except ValueError:
            print("200 response was not JSON.")
            continue

        print("✓ JSON schema/document discovered")

        fragments = recursively_find_path_fragments(
            data,
            set(KNOWN_MESSAGE_PATHS),
        )

        if fragments:
            for endpoint, fragment in fragments.items():
                schema_hits[endpoint] = fragment
                print()
                print(f"SCHEMA FRAGMENT FOR {endpoint}")
                print("-" * 70)
                print(
                    json.dumps(
                        fragment,
                        indent=2,
                        default=str,
                    )[:15000]
                )
        else:
            print(
                "No exact known messaging paths were found "
                "in this JSON document."
            )

    if not schema_hits:
        print()
        print(
            "No usable public API schema was discovered from "
            "the common read-only locations."
        )

    return schema_hits


def print_local_payload_hypotheses(
    order_id,
    thread_id,
    messaging_reasons,
    schema_hits,
):
    section("LOCAL PAYLOAD HYPOTHESES - NOTHING IS SENT")

    delivery_reason = None

    for reason in messaging_reasons:
        if (
            reason.get("code") == "20"
            and reason.get("type")
            in {"MESSAGING", "ORDER_MESSAGING"}
        ):
            delivery_reason = reason
            break

    print(
        "These are LOCAL representations only. "
        "They are NOT claimed to be correct POST payloads unless "
        "the schema probe above proves them."
    )

    print()
    print("Observed facts from this Tesco instance:")
    print("- Existing manual thread topic type is FREE_TEXT.")
    print(
        "- OPTIONS exposes POST on "
        "/api/orders/{order_id}/threads."
    )
    print(
        "- OPTIONS exposes POST on "
        "/api/orders/{order_id}/messages."
    )

    if thread_id:
        print(
            "- OPTIONS exposes POST on "
            "/api/inbox/threads/{thread_id}/message."
        )

    if delivery_reason:
        print(
            "- Order messaging reason code 20 exists: "
            f"{delivery_reason.get('label')}"
        )

    print()
    print("Candidate A - FREE_TEXT new-thread concept")
    candidate_a = {
        "order_id": order_id,
        "topic": {
            "type": "FREE_TEXT",
            "value": (
                "Electrical Discount UK delivery information "
                "& copy invoice"
            ),
        },
        "body": "<customer message>",
        "attachment": "<invoice PDF>",
    }
    print(json.dumps(candidate_a, indent=2))

    print()
    print("Candidate B - reason-code new-thread concept")
    candidate_b = {
        "order_id": order_id,
        "topic": {
            "type": "REASON_CODE",
            "value": "20",
        },
        "body": "<customer message>",
        "attachment": "<invoice PDF>",
    }
    print(json.dumps(candidate_b, indent=2))

    if thread_id:
        print()
        print("Candidate C - existing-thread reply concept")
        candidate_c = {
            "thread_id": thread_id,
            "body": "<customer message>",
            "attachment": "<invoice PDF>",
        }
        print(json.dumps(candidate_c, indent=2))

    print()
    if schema_hits:
        print(
            "A schema fragment was discovered. Use the schema output "
            "above, not these candidates, to design the next test."
        )
    else:
        print(
            "No POST payload has been proven yet. The candidate objects "
            "above are deliberately only hypotheses for comparison."
        )


def main():
    section("TESCO MIRAKL POST PAYLOAD DISCOVERY")

    print("ABSOLUTELY NO POST REQUESTS")
    print("Permitted HTTP methods: GET, HEAD, OPTIONS only.")
    print(
        "The safe_request() function hard-blocks any other method."
    )

    order_id = input(
        f"\nTesco order ID [{DEFAULT_ORDER_ID}]: "
    ).strip()

    if not order_id:
        order_id = DEFAULT_ORDER_ID

    section("EXISTING THREAD RAW STRUCTURE - READ ONLY")

    thread_data = read_threads(order_id)
    print_thread_structure(thread_data)

    thread_id = find_first_thread_id(thread_data)

    inspect_order_messages(order_id)
    messaging_reasons = inspect_reasons()
    inspect_options(order_id, thread_id)
    schema_hits = probe_api_schemas()

    print_local_payload_hypotheses(
        order_id,
        thread_id,
        messaging_reasons,
        schema_hits,
    )

    section("DISCOVERY COMPLETE")
    print("NO POST REQUESTS WERE MADE.")
    print("NO CUSTOMER MESSAGE WAS CREATED.")
    print("NO MIRAKL DATA WAS CHANGED.")
    print()
    print(
        "Please send back the output, especially:\n"
        "1. Raw thread/message keys and attachment fields\n"
        "2. GET /api/orders/{order_id}/messages\n"
        "3. Any OPENAPI/SWAGGER schema fragment discovered"
    )


if __name__ == "__main__":
    main()
