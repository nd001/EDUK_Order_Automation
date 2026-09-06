from dataclasses import dataclass
from typing import Optional


@dataclass
class MarketplaceOrder:
    """
    Common order format used by the core EDUK order workflow.

    Every marketplace adapter should eventually convert its
    marketplace-specific order data into this structure.

    Dictionary-style access is temporarily supported so the
    existing workflow can be migrated gradually.
    """

    marketplace: str
    order_id: str
    sku: str
    price: float
    postcode: str
    phone_1: str = ""
    phone_2: str = ""
    customer_name: Optional[str] = None
    description: Optional[str] = None

    def __getitem__(self, key):
        """
        Temporary compatibility with existing code such as:

            order["order_id"]
        """

        return getattr(self, key)

    def get(self, key, default=None):
        """
        Temporary compatibility with existing dictionary .get()
        calls.
        """

        return getattr(
            self,
            key,
            default,
        )