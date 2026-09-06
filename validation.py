# ============================================================
# VALIDATION HELPERS
# ============================================================


def normalise_phone(number):
    """
    Normalise UK telephone numbers for comparison.

    Examples:
        07761022352
        +44 7761 022352
        447761022352

    all become:
        07761022352
    """

    if not number:
        return ""

    digits = "".join(
        char
        for char in str(number)
        if char.isdigit()
    )

    if digits.startswith("44"):
        digits = "0" + digits[2:]

    return digits


def normalise_postcode(postcode):
    """
    Normalise a postcode for comparison.

    Example:
        "DN4 0SR" -> "DN40SR"
    """

    if not postcode:
        return ""

    return (
        str(postcode)
        .replace(" ", "")
        .strip()
        .upper()
    )


def normalise_sku(sku):
    """
    Normalise a SKU for comparison.
    """

    if sku is None:
        return ""

    return str(sku).strip().upper()


def normalise_price(value):
    """
    Normalise a monetary value to 2 decimal places.
    """

    return round(
        float(value),
        2
    )


def build_three_way_safety_result(
    excel_order,
    rp2_data,
    mirakl_order,
):
    """
    Compare Excel, RPii and Mirakl.

    Returns:
        {
            "failures": [...],
            "values": {...}
        }

    This function performs NO override and NO prompting.
    It only evaluates the data.
    """

    failures = []

    # --------------------------------------------------------
    # ORDER NUMBER
    # --------------------------------------------------------

    excel_order_id = str(
        excel_order.get("order_id") or ""
    ).strip()

    rp2_order_id = str(
        rp2_data.get("order_id") or ""
    ).strip()

    mirakl_order_id = str(
        mirakl_order.get("order_id") or ""
    ).strip()

    if not (
        excel_order_id
        == rp2_order_id
        == mirakl_order_id
    ):
        failures.append(
            "Order number"
        )

    # --------------------------------------------------------
    # SKU
    # --------------------------------------------------------

    excel_sku = normalise_sku(
        excel_order.get("sku")
    )

    rp2_sku = normalise_sku(
        rp2_data.get("sku")
    )

    mirakl_sku = normalise_sku(
        mirakl_order.get("sku")
    )

    if not (
        excel_sku
        == rp2_sku
        == mirakl_sku
    ):
        failures.append(
            "SKU"
        )

    # --------------------------------------------------------
    # PRICE
    # --------------------------------------------------------

    excel_price = normalise_price(
        excel_order.get("price", 0)
    )

    rp2_price = normalise_price(
        rp2_data.get("price", 0)
    )

    mirakl_price = normalise_price(
        mirakl_order.get("price", 0)
    )

    if not (
        excel_price
        == rp2_price
        == mirakl_price
    ):
        failures.append(
            "Price"
        )

    # --------------------------------------------------------
    # POSTCODE
    # --------------------------------------------------------

    excel_postcode = normalise_postcode(
        excel_order.get("postcode")
    )

    rp2_postcode = normalise_postcode(
        rp2_data.get("postcode")
    )

    mirakl_postcode = normalise_postcode(
        mirakl_order.get("postcode")
    )

    if not (
        excel_postcode
        == rp2_postcode
        == mirakl_postcode
    ):
        failures.append(
            "Postcode"
        )

    # --------------------------------------------------------
    # TELEPHONE
    # --------------------------------------------------------

    excel_phones = {
        normalise_phone(
            excel_order.get("phone_1")
        ),
        normalise_phone(
            excel_order.get("phone_2")
        ),
    }

    excel_phones.discard("")

    rp2_phone = normalise_phone(
        rp2_data.get("telephone")
    )

    mirakl_phones = {
        normalise_phone(
            mirakl_order.get("phone_1")
        ),
        normalise_phone(
            mirakl_order.get("phone_2")
        ),
    }

    mirakl_phones.discard("")

    phone_match = (
        bool(rp2_phone)
        and rp2_phone in excel_phones
        and rp2_phone in mirakl_phones
    )

    if not phone_match:
        failures.append(
            "Telephone"
        )

    return {
        "failures": failures,
        "values": {
            "excel_order_id": excel_order_id,
            "rp2_order_id": rp2_order_id,
            "mirakl_order_id": mirakl_order_id,

            "excel_sku": excel_sku,
            "rp2_sku": rp2_sku,
            "mirakl_sku": mirakl_sku,

            "excel_price": excel_price,
            "rp2_price": rp2_price,
            "mirakl_price": mirakl_price,

            "excel_postcode": excel_postcode,
            "rp2_postcode": rp2_postcode,
            "mirakl_postcode": mirakl_postcode,

            "excel_phones": excel_phones,
            "rp2_phone": rp2_phone,
            "mirakl_phones": mirakl_phones,
        },
    }


def can_override_sku_only(
    safety_result,
):
    """
    SKU override is allowed only when:

    - SKU is the ONLY failed safety check.
    - Excel SKU and RPii SKU agree.

    Mirakl may therefore be treated as the bad source
    only in this tightly controlled case.
    """

    failures = safety_result[
        "failures"
    ]

    values = safety_result[
        "values"
    ]

    return (
        failures == ["SKU"]
        and values["excel_sku"]
        == values["rp2_sku"]
    )