import os

import requests
from dotenv import load_dotenv


load_dotenv()

MIRAKL_API_KEY = os.getenv("MIRAKL_API_KEY")
MIRAKL_BASE_URL = os.getenv("MIRAKL_BASE_URL")

TEST_ORDER = "1068865081-A"


if not MIRAKL_API_KEY:
    raise RuntimeError(
        "MIRAKL_API_KEY is missing from .env"
    )

if not MIRAKL_BASE_URL:
    raise RuntimeError(
        "MIRAKL_BASE_URL is missing from .env"
    )


headers = {
    "Authorization": MIRAKL_API_KEY,
    "Accept": "application/json",
}


print("=" * 60)
print("MIRAKL READ-ONLY TEST")
print("=" * 60)
print(f"Order: {TEST_ORDER}")
print()


url = (
    f"{MIRAKL_BASE_URL.rstrip('/')}"
    f"/api/orders"
)

params = {
    "order_ids": TEST_ORDER
}


response = requests.get(
    url,
    headers=headers,
    params=params,
    timeout=30
)


print(
    f"HTTP Status: {response.status_code}"
)


if response.status_code != 200:

    print()
    print("Mirakl response:")
    print(response.text)

    raise SystemExit


data = response.json()

orders = data.get("orders", [])

if not orders:
    print()
    print("✗ Order was not found.")
    raise SystemExit

order = orders[0]

# ========================================================
# SAFETY CHECK - MIRAKL ORDER NUMBER
# ========================================================

mirakl_order_id = order.get("order_id")

print()
print("=" * 60)
print("MIRAKL SAFETY CHECK - ORDER NUMBER")
print("=" * 60)

print(f"Requested Order: {TEST_ORDER}")
print(f"Mirakl Order:    {mirakl_order_id}")

if mirakl_order_id != TEST_ORDER:

    print()
    print("✗ ORDER NUMBER MISMATCH")
    print("STOPPING - no further processing.")

    raise SystemExit

print()
print("✓ ORDER NUMBER MATCH")

customer = order.get("customer", {})
shipping = customer.get("shipping_address", {})

order_lines = order.get("order_lines", [])

if order_lines:
    line = order_lines[0]
else:
    line = {}

print()
print("=" * 60)
print("MIRAKL ORDER SUMMARY")
print("=" * 60)

print(
    f"Order ID:      {order.get('order_id')}"
)

print(
    f"Customer:      "
    f"{shipping.get('firstname', '')} "
    f"{shipping.get('lastname', '')}"
)

print(
    f"Postcode:      {shipping.get('zip_code')}"
)

print(
    f"Telephone:     {shipping.get('phone')}"
)

print(
    f"SKU:           {line.get('offer_sku')}"
)

print(
    f"Description:   {line.get('description')}"
)

print(
    f"Price:         £{order.get('total_price', 0):.2f}"
)

print(
    f"Order Status:  {order.get('order_state')}"
)

print(
    f"Has Invoice:   {order.get('has_invoice')}"
)

print(
    f"Customer Msg:  {order.get('has_customer_message')}"
)

print("=" * 60)

print()
print("=" * 60)
print("MIRAKL THREAD TEST")
print("=" * 60)

threads_url = (
    f"{MIRAKL_BASE_URL.rstrip('/')}"
    f"/api/inbox/threads"
)

thread_params = {
    "entity_type": "MMP_ORDER",
    "entity_id": TEST_ORDER,
    "with_messages": "true",
}

thread_response = requests.get(
    threads_url,
    headers=headers,
    params=thread_params,
    timeout=30,
)

print(
    f"HTTP Status: {thread_response.status_code}"
)

if thread_response.status_code != 200:
    print()
    print("Mirakl response:")
    print(thread_response.text)
    raise SystemExit

thread_data = thread_response.json()

print()
print("=" * 60)
print("MIRAKL EXISTING THREADS")
print("=" * 60)

threads = thread_data.get("data", [])

if not threads:

    print("No existing customer threads found.")

else:

    print(f"Threads found: {len(threads)}")

    for thread_number, thread in enumerate(
        threads,
        start=1
    ):

        print()
        print(f"THREAD {thread_number}")
        print("-" * 60)

        print(
            f"Thread ID:     {thread.get('id')}"
        )

        topic = thread.get("topic", {})

        print(
            f"Topic Type:    {topic.get('type')}"
        )

        print(
            f"Topic Code:    {topic.get('value')}"
        )

        messages = thread.get(
            "messages",
            []
        )

        print(
            f"Messages:      {len(messages)}"
        )

        for message_number, message in enumerate(
            messages,
            start=1
        ):

            sender = message.get(
                "from",
                {}
            )

            print()
            print(
                f"Message {message_number}:"
            )

            print(
                f"From:          "
                f"{sender.get('display_name')}"
            )

            print(
                f"Date:          "
                f"{message.get('date_created')}"
            )

            print(
                f"Body:          "
                f"{message.get('body')}"
            )

print("=" * 60)