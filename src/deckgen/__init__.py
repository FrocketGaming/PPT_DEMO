"""deckgen: config-driven PowerPoint generation."""

from .builder import build
from .components import register

__all__ = ["build", "register"]
