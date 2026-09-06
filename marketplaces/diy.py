import html
import os
import re

import requests
from openpyxl import load_workbook

from config import (
    MIRAKL_BASE_URL,
    MIRAKL_API_KEY,
    DIY_DELIVERY_TOPIC_CODE,
)

from models import MarketplaceOrder


# ============================================================
# DIY / B&Q EXCEL IMPORT
# ============================================================

def read_first_bq_order(filename):
    """
    Read the first/top order from the B&Q unshipped-orders
    Excel export.
    """

    workbook = load_workbook(
        filename,
        data_only=True
    )

    sheet = workbook.active

    headers = {
        cell.value: cell.column
        for cell in sheet[1]
        if cell.value
    }

    required_columns = [
        "Order number",
        "Offer SKU",
        "Amount",
        "Billing address zip",
        "Shipping address phone",
        "Shipping address phone 2",
    ]

    for column in required_columns:

        if column not in headers:

            raise RuntimeError(
                f"Missing required B&Q Excel column: "
                f"{column}"
            )

    # First data row = oldest/top order
    row = 2

    order_id = sheet.cell(
        row=row,
        column=headers["Order number"]
    ).value

    sku = sheet.cell(
        row=row,
        column=headers["Offer SKU"]
    ).value

    price = sheet.cell(
        row=row,
        column=headers["Amount"]
    ).value

    postcode = sheet.cell(
        row=row,
        column=headers["Billing address zip"]
    ).value

    phone_1 = sheet.cell(
        row=row,
        column=headers["Shipping address phone"]
    ).value

    phone_2 = sheet.cell(
        row=row,
        column=headers["Shipping address phone 2"]
    ).value

    return MarketplaceOrder(
    marketplace="DIY",

    order_id=str(order_id).strip(),

    sku=str(sku).strip(),

    price=float(price),

    postcode=str(postcode).strip(),

    phone_1=(
        str(phone_1).strip()
        if phone_1
        else ""
    ),

    phone_2=(
        str(phone_2).strip()
        if phone_2
        else ""
    ),
)


# ============================================================
# MIRAKL ORDER LOOKUP
# ============================================================

def read_mirakl_order(order_id):
    """
    Read one B&Q/Kingfisher Mirakl order.

    READ ONLY.
    """

    url = (
        f"{MIRAKL_BASE_URL.rstrip('/')}"
        f"/api/orders"
    )

    headers = {
        "Authorization": MIRAKL_API_KEY,
        "Accept": "application/json",
    }

    params = {
        "order_ids": order_id
    }

    response = requests.get(
        url,
        headers=headers,
        params=params,
        timeout=30,
    )

    if response.status_code != 200:

        raise RuntimeError(
            f"Mirakl order lookup failed. "
            f"HTTP {response.status_code}: "
            f"{response.text}"
        )

    data = response.json()

    orders = data.get(
        "orders",
        []
    )

    if len(orders) != 1:

        raise RuntimeError(
            f"Expected exactly 1 Mirakl order, "
            f"found {len(orders)}."
        )

    order = orders[0]

    customer = order.get(
        "customer",
        {}
    )

    shipping = customer.get(
        "shipping_address",
        {}
    )

    order_lines = order.get(
        "order_lines",
        []
    )

    if not order_lines:

        raise RuntimeError(
            "Mirakl order contains no order lines."
        )

    # Current prototype assumes one active product line.
    line = order_lines[0]

    return {
        "order_id": order.get(
            "order_id"
        ),
        "sku": line.get(
            "offer_sku"
        ),
        "price": float(
            order.get(
                "total_price",
                0
            )
        ),
        "postcode": shipping.get(
            "zip_code"
        ),
        "phone_1": shipping.get(
            "phone"
        ),
        "phone_2": shipping.get(
            "phone_secondary"
        ),
        "customer_name": (
            f"{shipping.get('firstname', '')} "
            f"{shipping.get('lastname', '')}"
        ).strip(),
        "description": line.get(
            "description"
        ),
        "order_state": order.get(
            "order_state"
        ),
    }


# ============================================================
# MIRAKL THREADS
# ============================================================

def clean_mirakl_message_body(body):
    """
    Convert Mirakl HTML message content into readable
    terminal text.
    """

    if not body:
        return ""

    text = body

    text = re.sub(
        r"<br\s*/?>",
        "\n",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"<[^>]+>",
        "",
        text
    )

    text = html.unescape(
        text
    )

    return text.strip()


def read_mirakl_threads(order_id):
    """
    Read existing Mirakl threads/messages for a B&Q order.

    READ ONLY.
    """

    url = (
        f"{MIRAKL_BASE_URL.rstrip('/')}"
        f"/api/inbox/threads"
    )

    headers = {
        "Authorization": MIRAKL_API_KEY,
        "Accept": "application/json",
    }

    params = {
        "entity_type": "MMP_ORDER",
        "entity_id": order_id,
        "with_messages": "true",
    }

    response = requests.get(
        url,
        headers=headers,
        params=params,
        timeout=30,
    )

    if response.status_code != 200:

        raise RuntimeError(
            f"Mirakl thread lookup failed. "
            f"HTTP {response.status_code}: "
            f"{response.text}"
        )

    data = response.json()

    threads = data.get(
        "data",
        []
    )

    all_messages = []

    for thread in threads:

        topic = thread.get(
            "topic",
            {}
        )

        for message in thread.get(
            "messages",
            []
        ):

            sender = message.get(
                "from",
                {}
            )

            all_messages.append(
                {
                    "thread_id": thread.get(
                        "id"
                    ),
                    "topic_type": topic.get(
                        "type"
                    ),
                    "topic_code": topic.get(
                        "value"
                    ),
                    "message_id": message.get(
                        "id"
                    ),
                    "date": message.get(
                        "date_created"
                    ),
                    "sender_type": sender.get(
                        "type"
                    ),
                    "sender_name": sender.get(
                        "display_name"
                    ),
                    "body": clean_mirakl_message_body(
                        message.get(
                            "body"
                        )
                    ),
                }
            )

    all_messages.sort(
        key=lambda message: (
            message["date"] or ""
        )
    )

    return {
        "threads": threads,
        "messages": all_messages,
    }


# ============================================================
# MIRAKL DRY-RUN PACKAGE
# ============================================================

def build_mirakl_dry_run_payload(
    order_id,
    customer_name,
    message_body,
    invoice_path,
    existing_threads,
):
    """
    Build our LOCAL representation of the Mirakl delivery
    message.

    IMPORTANT:
    No network write occurs here.
    """

    thread_id = None

    if existing_threads:
        thread_id = existing_threads[0].get(
            "id"
        )

    return {
        "mode": "DRY_RUN",
        "order_id": order_id,
        "customer": customer_name,
        "topic": {
            "type": "REASON_CODE",
            "value": DIY_DELIVERY_TOPIC_CODE,
            "label": (
                "Information about delivery "
                "(incl. tracking)"
            ),
        },
        "existing_thread_id": thread_id,
        "message_body": message_body,
        "attachment": {
            "filename": os.path.basename(
                invoice_path
            ),
            "path": invoice_path,
            "exists": os.path.isfile(
                invoice_path
            ),
        },
    }