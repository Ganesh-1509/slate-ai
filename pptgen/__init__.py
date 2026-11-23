"""pptgen - turn a topic + style description into a PowerPoint deck."""
import re
from typing import Callable, Optional, Tuple

from .builder import DeckBuilder
from .content import ContentError, generate_deck
from .template import TemplateDeckBuilder, TemplateError, describe_template, load_template
from .themes import PRESETS, pick_theme

__all__ = ["make_presentation", "describe_template", "ContentError", "TemplateError", "PRESETS"]


def make_presentation(topic: str, theme: str = "minimal", num_slides: int = 8,
                      audience: str = "general audience", api_key: Optional[str] = None,
                      log: Callable[[str], None] = print,
                      template: Optional[bytes] = None) -> Tuple[str, bytes]:
    """Used by both the CLI and the website. Returns (file name, .pptx bytes).

    template: optional .pptx/.potx file contents; its layouts and colours are used.
    """
    if template is not None:
        try:
            load_template(template)  # fail fast, before spending AI quota
        except TemplateError as e:
            raise ContentError(str(e)) from e
        theme = theme or "match the uploaded template"

    log("Planning slides with Gemini...")
    deck = generate_deck(topic, theme, num_slides, audience, api_key, log)
    log(f"Got {len(deck.slides)} slides. Fetching images and building PowerPoint...")
    if template is not None:
        builder = TemplateDeckBuilder(template, pick_theme(theme, deck.theme))
    else:
        builder = DeckBuilder(pick_theme(theme, deck.theme))
    data = builder.build(deck, log)
    slug = re.sub(r"[^a-z0-9]+", "-", deck.title.lower()).strip("-")[:60] or "deck"
    return f"{slug}.pptx", data
