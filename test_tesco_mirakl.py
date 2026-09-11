import requests

from config import (
    TESCO_MIRAKL_URL,
    TESCO_MIRAKL_API_KEY,
)


TEST_ORDER_ID = "4431-2696-2301-A"


def read_tesco_mirakl_order(order_id):

    if not TESCO_MIRAKL_API_KEY:

        raise RuntimeError(
            "TESCO_MIRAKL_API_KEY is missing."
        )

    url = (
        f"{TESCO_MIRAKL_URL}"
        f"/api/orders"
    )

    params = {
        "order_ids": order_id,
    }

    headers = {
        "Authorization": TESCO_MIRAKL_API_KEY,
        "Accept": "application/json",
    }

    print()
    print("=" * 60)
    print("TESCO MIRAKL READ-ONLY TEST")
    print("=" * 60)

    print(
        f"Order: {order_id}"
    )

    print(
        f"Host:  {TESCO_MIRAKL_URL}"
    )

    print()
    print(
        "GET request only."
    )

    print(
        "NO MIRAKL DATA WILL BE CHANGED."
    )

    response = requests.get(
        url,
        headers=headers,
        params=params,
        timeout=30,
    )
    print()

    print(
        f"HTTP status: {response.status_code}"
    )

    if response.status_code != 200:

        print()
        print("Mirakl response:")
        print(response.text)

        return None

    data = response.json()

    print()
    print(
        "✓ Tesco Mirakl order retrieved successfully."
    )

    orders = data.get(
        "orders",
        []
    )

    if not orders:

        print()
        print(
            "✗ No Tesco Mirakl order was returned."
        )

        return None

    return orders[0]


if __name__ == "__main__":

    order = read_tesco_mirakl_order(
        TEST_ORDER_ID
    )

    if order is not None:

        print()
        print("=" * 60)
        print("TESCO MIRAKL ORDER SUMMARY")
        print("=" * 60)

        print(
            f"Order ID:          "
            f"{order.get('order_id', '')}"
        )

        print(
            f"Order state:       "
            f"{order.get('order_state', '')}"
        )

        print(
            f"Price:             "
            f"{order.get('price', '')}"
        )

        print(
            f"Total price:       "
            f"{order.get('total_price', '')}"
        )

        print(
            f"Shipping deadline: "
            f"{order.get('shipping_deadline', '')}"
        )

        print()

        print()
        print("=" * 60)
        print("TESCO MIRAKL CUSTOMER")
        print("=" * 60)

        customer = order.get(
            "customer",
            {}
        )

        shipping_address = customer.get(
            "shipping_address",
            {}
        )

        print(
            f"Name:      "
            f"{customer.get('firstname', '')} "
            f"{customer.get('lastname', '')}"
        )

        print(
            f"Postcode:  "
            f"{shipping_address.get('zip_code', '')}"
        )

        print(
            f"Phone:     "
            f"{shipping_address.get('phone', '')}"
        )

        print(
            f"Phone 2:   "
            f"{shipping_address.get('phone_secondary', '')}"
        )

        print()

        print("=" * 60)
        print("TESCO MIRAKL ORDER LINE")
        print("=" * 60)

        order_lines = order.get(
            "order_lines",
            []
        )

        if order_lines:

            first_line = order_lines[0]

            print(
                f"Offer SKU:      "
                f"{first_line.get('offer_sku', '')}"
            )

            print(
                f"Product SKU:    "
                f"{first_line.get('product_sku', '')}"
            )

            print(
                f"Shop SKU:       "
                f"{first_line.get('product_shop_sku', '')}"
            )

            print(
                f"Quantity:       "
                f"{first_line.get('quantity', '')}"
            )

            print(
                f"Line price:     "
                f"{first_line.get('price', '')}"
            )

            print(
                f"Line total:     "
                f"{first_line.get('total_price', '')}"
            )

            print(
                f"Line state:     "
                f"{first_line.get('order_line_state', '')}"
            )

        else:

            print(
                "No order lines returned."
            )

        print()
        print("=" * 60)
        print("READ-ONLY TEST COMPLETE")
        print("=" * 60)

        if order_lines:

            first_line = order_lines[0]

            for key in sorted(
                first_line.keys()
            ):

                print(key)

        else:

            print(
                "No order lines returned."
            )

        print()
        print("=" * 60)
        print("READ-ONLY TEST COMPLETE")
        print("=" * 60)