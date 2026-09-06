from abc import ABC, abstractmethod

from models import MarketplaceOrder


class MarketplaceAdapter(ABC):
    """
    Base interface for marketplace order sources.

    Each marketplace adapter is responsible for reading its
    marketplace-specific data and converting it into the common
    MarketplaceOrder format.

    The core EDUK workflow should not need to know how the
    marketplace stores or supplies its order data.
    """

    @property
    @abstractmethod
    def name(self) -> str:
        """
        Human-readable marketplace name.
        """
        raise NotImplementedError

    @abstractmethod
    def read_first_order(self) -> MarketplaceOrder:
        """
        Read the next marketplace order to be processed.

        Returns:
            MarketplaceOrder
        """
        raise NotImplementedError