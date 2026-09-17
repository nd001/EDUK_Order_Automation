import re
from datetime import datetime, timedelta


# ============================================================
# RPii PAYMENT
# ============================================================

def read_payment_summary(sales_frame):
    """
    Read the RPii Payment Summary.

    A sale is considered payment-processed if Payment Summary
    contains at least one payment row beneath the heading.

    PREPAL may legitimately show 0.00, so the payment amount
    itself is not used to determine processing status.
    """

    payment_table = sales_frame.locator("#psum")

    if payment_table.count() == 0:
        return {
            "processed": False,
            "details": None,
        }

    rows = payment_table.locator("tr")

    payment_details = []

    for i in range(rows.count()):
        row = rows.nth(i)

        text = " ".join(
            row.inner_text().split()
        )

        if not text:
            continue

        if text.lower() == "payment summary":
            continue

        payment_details.append(text)

    if payment_details:
        return {
            "processed": True,
            "details": payment_details,
        }

    return {
        "processed": False,
        "details": None,
    }


# ============================================================
# RPii WEB REFERENCE
# ============================================================

def extract_web_ref(text):
    """
    Extract:

        Web Ref: 1068814250-A

    from the larger RPii sale-information panel.
    """

    match = re.search(
        r"Web Ref:\s*([A-Za-z0-9\-]+)",
        text
    )

    return match.group(1) if match else None

# ============================================================
# RPii     Extract the RPii internal order number from the sale header.
# ============================================================

def extract_rp2_order_number(text):
    """
    Extract the RPii internal order number from the sale header.

    Example:

        992819 • 11 • 10/ 9/26 11.00
        • TESCO Sale
        ...
        Web Ref: 8431-2631-101-A

    Returns:
        "992819"
    """

    cleaned = " ".join(
        str(text or "").split()
    )

    match = re.match(
        r"^(\d+)\s*•",
        cleaned
    )

    if not match:
        return None

    return match.group(1)


# ============================================================
# RPii CUSTOMER
# ============================================================

def extract_customer_name(text):
    """
    Extract the customer name from an RPii customer panel.
    """

    parts = [
        part.strip()
        for part in text.split("•")
    ]

    if len(parts) >= 2:
        return parts[1]

    return None


def extract_delivery_customer(text):
    """
    Example:

        JEEV9737 • Jackie Jeeves •
        36, Smith Square, Doncaster, DN4 0SR
        T: 07761022352 • e: ...
    """

    cleaned = " ".join(
        text.split()
    )

    parts = [
        part.strip()
        for part in cleaned.split("•")
    ]

    name = None
    address = None
    telephone = None

    if len(parts) >= 2:
        name = parts[1]

    if len(parts) >= 3:
        address = parts[2]

        address = re.split(
            r"\bT:\s*",
            address
        )[0].strip()

    tel_match = re.search(
        r"T:\s*(\+?[0-9 ]+)",
        cleaned
    )

    if tel_match:
        telephone = tel_match.group(1).strip()

    return {
        "name": name,
        "address": address,
        "telephone": telephone,
    }


def extract_uk_postcode(address):
    """
    Extract and normalise a UK postcode.
    """

    if not address:
        return None

    match = re.search(
        r"\b("
        r"[A-Z]{1,2}\d[A-Z\d]?"
        r"\s*"
        r"\d[A-Z]{2}"
        r")\b",
        address.upper()
    )

    if not match:
        return None

    postcode = match.group(1).replace(
        " ",
        ""
    )

    return (
        postcode[:-3]
        + " "
        + postcode[-3:]
    )


# ============================================================
# RPii DELIVERY
# ============================================================

def parse_delivery_block(text):
    """
    Example SGK:

        DELIVER SK
        BN:09
        070926:AM
        SK

    Example EDUK:

        DELIVER ED
        BN:09
        120926:PM
        ED

    Business rule:

        SGK:
            use the RPii date exactly.

        ED:
            customer delivery date is one day before
            the RPii date.
    """

    cleaned = " ".join(
        text.split()
    )

    courier = None

    # --------------------------------------------------------
    # Pre-manifest RPii format
    # --------------------------------------------------------

    if "DELIVER SK" in cleaned:
        courier = "SK"

    elif "DELIVER ED" in cleaned:
        courier = "ED"

    # --------------------------------------------------------
    # On-manifest RPii format
    #
    # Example:
    # ON MANFST: 67094 D SK
    # BN:09
    # 090926:AM • 67094:SGK
    # --------------------------------------------------------

    elif "ON MANFST:" in cleaned and (
        " D SK" in cleaned
        or ":SGK" in cleaned
    ):
        courier = "SK"

    elif "ON MANFST:" in cleaned and (
        " D ED" in cleaned
        or ":ED" in cleaned
    ):
        courier = "ED"

    # --------------------------------------------------------
    # Delivered RPii format
    # --------------------------------------------------------

    elif re.search(
        r"\bDELVRD\b.*\bSK\b",
        cleaned
    ):
        courier = "SK"

    elif re.search(
        r"\bDELVRD\b.*\bED\b",
        cleaned
    ):
        courier = "ED"
    date_match = re.search(
        r"(\d{6}):(AM|PM)",
        cleaned
    )

    rp2_date = None
    customer_date = None
    slot = None

    if date_match:
        date_text = date_match.group(1)
        slot = date_match.group(2)

        rp2_date = datetime.strptime(
            date_text,
            "%d%m%y"
        ).date()

        if courier == "ED":
            customer_date = (
                rp2_date
                - timedelta(days=1)
            )

        else:
            customer_date = rp2_date

    return {
        "courier": courier,
        "rp2_date": rp2_date,
        "customer_date": customer_date,
        "slot": slot,
        "raw": cleaned,
    }


# ============================================================
# RPii PRODUCT
# ============================================================

def read_active_product(sales_frame):
    """
    Identify the first RPii product line.

    RPii displays the product SKU before the first opening
    parenthesis. Anything inside the parentheses is RPii
    internal information and must not be interpreted as
    quantity.

    Quantity is deliberately not parsed here yet.
    """

    product_cells = sales_frame.locator(
        "td.SALES-lines-left-wrap"
    )

    product_cell_count = product_cells.count()

    sku = None
    description = None
    quantity = None

    for i in range(
        0,
        product_cell_count - 1,
        2
    ):

        sku_cell = product_cells.nth(i)
        desc_cell = product_cells.nth(i + 1)

        sku_text = " ".join(
            sku_cell.inner_text().split()
        )

        raw_description_text = (
            desc_cell.inner_text()
            or ""
        )

        description_text = (
            " ".join(
                raw_description_text.split()
            )
            or None
        )

        # RPii SKU is everything before the first
        # opening parenthesis.
        #
        # Examples:
        #
        # HUSHU269 (HUS-HU269)
        # -> HUSHU269
        #
        # EY 48SB8 (...)
        # -> EY48SB8
        #
        # Anything inside the parentheses is RPii
        # internal information and is NOT quantity.

        sku_part = sku_text.split(
            "(",
            1
        )[0].strip()

        sku = re.sub(
            r"\s+",
            "",
            sku_part,
        )

        # Ignore empty/non-product cells.
        if not sku:
            continue

        print()
        print("RPii SKU PARSE DIAGNOSTIC")
        print("-" * 60)

        print(
            f"Raw SKU cell:    {sku_text}"
        )

        print(
            f"Cleaned SKU:     {sku}"
        )

        print(
            f"Raw description: "
            f"{raw_description_text!r}"
        )

        print(
            f"Description:     "
            f"{description_text!r}"
        )

        print(
            "Quantity:        "
            "Not currently parsed"
        )

        print("-" * 60)

        description = description_text

        # Quantity parsing deliberately deferred.
        quantity = None

        break

    return {
        "sku": sku,
        "description": description,
        "quantity": quantity,
    }