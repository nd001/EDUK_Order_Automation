import json
import os
import sys
from typing import Any

import requests
from dotenv import load_dotenv


load_dotenv()

TESCO_MIRAKL_URL = os.getenv(
    "TESCO_MIRAKL_URL",
    "https://tescouk-prod.mirakl.net",
).rstrip("/")

TESCO_MIRAKL_API_KEY = os.getenv(
    "TESCO_MIRAKL_API_KEY",
    "",
).strip()

INTERESTING_WORDS = (
    "track",
    "carrier",
    "ship",
    "delivery",
    "logistic",
    "freight",
)


def is_interesting_key(key: str) -> bool:
    key_lower = str(key).lower()
    return any(word in key_lower for word in INTERESTING_WORDS)


def walk_for_interesting_fields(value: Any, path: str = "order"):
    if isinstance(value, dict):
        for key, child in value.items():
            child_path = f"{path}.{key}"

            if is_interesting_key(key):
                print(
                    f"{child_path}: "
                    f"{json.dumps(child, ensure_ascii=False, default=str)}"
                )

            walk_for_interesting_fields(child, child_path)

    elif isinstance(value, list):
        for index, child in enumerate(value):
            walk_for_interesting_fields(
                child,
                f"{path}[{index}]",
            )


def fetch_tesco_order(order_id: str) -> dict:
    if not TESCO_MIRAKL_API_KEY:
        raise RuntimeError(
            "TESCO_MIRAKL_API_KEY is missing from .env"
        )

    url = f"{TESCO_MIRAKL_URL}/api/orders"

    headers = {
        "Authorization": TESCO_MIRAKL_API_KEY,
        "Accept": "application/json",
    }

    response = requests.get(
        url,
        headers=headers,
        params={"order_ids": order_id},
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
            f"Expected exactly 1 order, found {len(orders)}"
        )

    return orders[0]


def main():
    print()
    print("=" * 80)
    print("TESCO MIRAKL TRACKING / CARRIER DISCOVERY")
    print("=" * 80)
    print()
    print("READ ONLY")
    print("This script performs GET requests only.")
    print("It cannot alter tracking or ship an order.")
    print()

    order_id = input(
        "Enter an ALREADY SHIPPED Tesco order ID: "
    ).strip()

    if not order_id:
        print("No order ID entered.")
        sys.exit(1)

    print()
    print(f"Reading Tesco Mirakl order: {order_id}")
    print()

    order = fetch_tesco_order(order_id)

    print("=" * 80)
    print("BASIC ORDER DETAILS")
    print("=" * 80)
    print(f"Order ID:    {order.get('order_id')}")
    print(f"Order state: {order.get('order_state')}")

    print()
    print("=" * 80)
    print("TRACKING / CARRIER / SHIPPING FIELDS")
    print("=" * 80)
    print()

    walk_for_interesting_fields(order)

    print()
    print("=" * 80)
    print("FULL ORDER JSON")
    print("=" * 80)
    print()
    print(
        json.dumps(
            order,
            indent=2,
            ensure_ascii=False,
            default=str,
        )
    )

    print()
    print("=" * 80)
    print("DISCOVERY COMPLETE - NO WRITE ACTION PERFORMED")
    print("=" * 80)


if __name__ == "__main__":
    main()
