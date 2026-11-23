"""pptgen - turn a topic + style description into a PowerPoint deck."""
import re
from typing import Callable, Optional, Tuple

from .builder import DeckBuilder
from .content import ContentError, generate_deck
from .themes import PRESETS, pick_theme

__all__ = ["make_presentation", "ContentError", "PRESETS"]


def make_presentation(topic: str, theme: str = "minimal", num_slides: int = 8,
                      audience: str = "general audience", api_key: Optional[str] = None,
                      log: Callable[[str], None] = print) -> Tuple[str, bytes]:
    """Used by both the CLI and the website. Returns (file name, .pptx bytes)."""
    log("Planning slides with Gemini...")
    deck = generate_deck(topic, theme, num_slides, audience, api_key, log)
    log(f"Got {len(deck.slides)} slides. Fetching images and building PowerPoint...")
    data = DeckBuilder(pick_theme(theme, deck.theme)).build(deck, log)
    slug = re.sub(r"[^a-z0-9]+", "-", deck.title.lower()).strip("-")[:60] or "deck"
    return f"{slug}.pptx", data
