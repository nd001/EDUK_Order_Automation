import json
import requests

from config import TESCO_MIRAKL_URL, TESCO_MIRAKL_API_KEY


TEST_ORDER = "4431-2696-2301-A"
TIMEOUT = 30


def auth_headers():
    if not TESCO_MIRAKL_API_KEY:
        raise RuntimeError("TESCO_MIRAKL_API_KEY is missing.")

    return {
        "Authorization": TESCO_MIRAKL_API_KEY,
        "Accept": "application/json",
    }


def print_response_summary(label, response):
    print()
    print("=" * 70)
    print(label)
    print("=" * 70)
    print(f"HTTP:    {response.status_code}")
    print(f"Allow:   {response.headers.get('Allow', '(not supplied)')}")
    print(f"Type:    {response.headers.get('Content-Type', '(not supplied)')}")

    text = (response.text or "").strip()
    if text:
        print("Body:")
        print(text[:2000])
    else:
        print("Body:    (empty)")


def get_existing_order_threads(order_id):
    url = f"{TESCO_MIRAKL_URL.rstrip('/')}/api/inbox/threads"
    params = {
        "entity_type": "MMP_ORDER",
        "entity_id": order_id,
        "with_messages": "true",
    }

    response = requests.get(
        url,
        headers=auth_headers(),
        params=params,
        timeout=TIMEOUT,
    )
    response.raise_for_status()

    data = response.json()
    threads = data.get("data", [])

    print()
    print("=" * 70)
    print("EXISTING TESCO MIRAKL THREADS - READ ONLY")
    print("=" * 70)
    print(f"Order:   {order_id}")
    print(f"Threads: {len(threads)}")

    for index, thread in enumerate(threads, start=1):
        topic = thread.get("topic") or {}
        print()
        print(f"Thread {index}")
        print(f"ID:          {thread.get('id')}")
        print(f"Topic type:  {topic.get('type')}")
        print(f"Topic value: {topic.get('value')}")
        print(f"Topic label: {topic.get('label') or topic.get('display_label')}")
        print(f"Messages:    {len(thread.get('messages') or [])}")

    return threads


def get_mirakl_reasons():
    url = f"{TESCO_MIRAKL_URL.rstrip('/')}/api/reasons"

    response = requests.get(
        url,
        headers=auth_headers(),
        timeout=TIMEOUT,
    )

    print_response_summary(
        "GET /api/reasons - READ ONLY",
        response,
    )

    if response.status_code != 200:
        return []

    data = response.json()
    reasons = data.get("reasons", [])

    print()
    print("MESSAGING-RELATED REASONS")
    print("-" * 70)

    messaging_reasons = []

    for reason in reasons:
        reason_type = str(reason.get("type", ""))
        if "MESSAG" in reason_type.upper():
            messaging_reasons.append(reason)
            print(json.dumps(reason, indent=2, ensure_ascii=False))

    if not messaging_reasons:
        print("No messaging-related reasons were identified in /api/reasons.")

    return messaging_reasons


def safe_options(label, path):
    url = f"{TESCO_MIRAKL_URL.rstrip('/')}{path}"

    response = requests.options(
        url,
        headers=auth_headers(),
        timeout=TIMEOUT,
    )

    print_response_summary(label, response)


def main():
    print("=" * 70)
    print("TESCO MIRAKL ENDPOINT DISCOVERY")
    print("=" * 70)
    print("SAFE DISCOVERY MODE")
    print("Only GET and OPTIONS requests are used.")
    print("NO POST, PUT, PATCH or DELETE requests are made.")
    print("No customer message can be created by this script.")

    order_id = input(
        f"\nTesco order ID [{TEST_ORDER}]: "
    ).strip() or TEST_ORDER

    threads = get_existing_order_threads(order_id)
    get_mirakl_reasons()

    thread_id = None
    if threads:
        thread_id = threads[0].get("id")

    safe_options(
        "OPTIONS /api/orders/{order_id}/threads",
        f"/api/orders/{order_id}/threads",
    )

    safe_options(
        "OPTIONS /api/orders/{order_id}/messages",
        f"/api/orders/{order_id}/messages",
    )

    if thread_id:
        safe_options(
            "OPTIONS /api/inbox/threads/{thread_id}/message",
            f"/api/inbox/threads/{thread_id}/message",
        )
    else:
        print()
        print("No existing thread found, so reply-endpoint OPTIONS probe was skipped.")

    print()
    print("=" * 70)
    print("DISCOVERY COMPLETE")
    print("=" * 70)
    print("NO WRITE REQUESTS WERE MADE.")
    print("Send me the output from this script before we implement any POST.")


if __name__ == "__main__":
    main()
