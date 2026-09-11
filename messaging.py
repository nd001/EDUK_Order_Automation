def build_customer_address(delivery_customer):
    """
    Format the RPii delivery address for the customer message.
    """

    name = delivery_customer.get("name") or ""
    address = delivery_customer.get("address") or ""
    postcode = delivery_customer.get("postcode") or ""
    telephone = delivery_customer.get("telephone") or ""

    lines = []

    if name:
        lines.append(name)

    address_parts = [
        part.strip()
        for part in address.split(",")
        if part.strip()
    ]

    address_parts = [
        part
        for part in address_parts
        if part.upper() != postcode.upper()
    ]

    if len(address_parts) >= 2:
        lines.append(
            f"{address_parts[0]}, {address_parts[1]}"
        )

        lines.extend(
            address_parts[2:]
        )

    elif address_parts:
        lines.extend(
            address_parts
        )

    if postcode:
        lines.append(
            postcode
        )

    if telephone:
        lines.append(
            f"T: {telephone}"
        )

    return "\n".join(lines)



def build_sgk_message(
    delivery_customer,
    product_sku,
    product_description,
    collection_date,
):
    address_block = build_customer_address(
        delivery_customer
    )

    product_description_text = str(
        product_description or ""
    ).strip()

    if product_description_text:
        item_text = (
            f"{product_sku} - "
            f"{product_description_text}"
        )
    else:
        item_text = str(product_sku).strip()

    return f"""Good Morning,

Thank you for placing your order with Electrical Discount UK.

***For any queries regarding delivery information, please call our office on 01282 443850 - Mon-Fri 9-5/Sat 9-3***

Please find attached a copy of your invoice.

Delivery Details:

Item: {item_text}

{address_block}

Please check that these details are correct and let us know if any changes are needed.

Your item(s) will be collected by SGK Distribution on {collection_date}. Once collected, the courier will contact you within 48 hours to arrange a convenient delivery date.

⚠️ Important: Missed or cancelled deliveries, once a date has been agreed, may incur additional charges. Please ensure someone is available to receive the delivery.

If you have any questions regarding your order or delivery, please contact us on 01282 443850 and we will be happy to help.

Kind regards,

Electrical Discount UK

Customer Service"""


def build_ed_message(
    delivery_customer,
    product_sku,
    product_description,
    delivery_date,
):
    address_block = build_customer_address(
        delivery_customer
    )

    product_description_text = str(
        product_description or ""
    ).strip()

    if product_description_text:
        item_text = (
            f"{product_sku} - "
            f"{product_description_text}"
        )
    else:
        item_text = str(product_sku).strip()

    return f"""Thank you for purchasing from Electrical Discount UK.

*** For ANY queries regarding delivery information, please call our office on 01282 443850 - Mon-Fri 9-5 / Sat 9-3 ***

Your item has been dispatched and is due for delivery on {delivery_date} between 7am and 7pm.

Item: {item_text}

The delivery address we have for you is:

{address_block}

We will email you the day before with an estimated delivery window.

Please note that circumstances such as traffic or weather may affect the time window given.

It is very important that you contact us if there are any steps or stairs leading up to your property. We may have to use a different courier altering the delivery day.

Please note this is a delivery only, no installation.

If you have paid for a scrap collection please make sure the item is disconnected and placed outside (unless other arrangements have been agreed) prior to the arrival of the driver.

If you require any further information please do not hesitate to contact our sales team on 01282 443850 (direct landline).

Kind Regards,

Electrical Discount UK"""

def build_sgk_sms():
    """
    Build the customer SMS for an SGK delivery.
    """

    return (
        "Electrical Discount UK: "
        "Thanks for your B&Q order. "
        "Your delivery will be made by SGK. "
        "SGK will contact you directly about delivery. "
        "We have emailed further information. "
        "If you have any issues, call 01282 443850."
    )


def build_ed_sms(
    sku,
    delivery_date,
    description=None,
):
    """
    Build the customer SMS for an EDUK delivery.

    Prefer the RPii product description when available.
    Fall back safely to the marketplace SKU if the
    description is missing.
    """

    product_text = str(
        description or ""
    ).strip()

    if not product_text:
        product_text = str(sku or "").strip()

    if not product_text:
        product_text = "item"

    return (
        "Electrical Discount UK: "
        "Thanks for your B&Q order. "
        f"Your {product_text} is due "
        f"{delivery_date}, 7am-7pm. "
        "Further delivery information has been emailed to you. "
        "We will email an estimated delivery window "
        "the day before. "
        "Queries: 01282 443850."
    )
