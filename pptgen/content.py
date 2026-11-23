"""Ask Google Gemini (free tier) to plan the whole deck as a validated `Deck` object."""
import os
import time
from typing import Callable, List, Optional

from google import genai
from google.genai import errors, types
from pydantic import ValidationError

from .models import Deck

# Free-tier models, tried in order. If one is overloaded (503) or its free quota
# is used up (429), we move on to the next one - each model has its own quota.
# Set GEMINI_MODEL in .env / Secrets to put a different model first.
FALLBACK_MODELS = [
    "gemini-3.8-flash",
    "gemini-3.6-flash",
    "gemini-3.5-flash-lite",
    "gemini-3.1-flash-lite",
]
RETRIES_PER_MODEL = 2      # quick retries on a busy model before moving on
RETRY_WAIT_SECONDS = 2

SYSTEM_PROMPT = """You are an expert presentation writer. You plan complete, informative
PowerPoint decks that someone could present directly with minimal editing.

CONTENT
- Be specific and informative for the requested audience: explain *why* and *how*, and
  include concrete examples, real applications, context, figures and consequences.
- Each point = a bold `heading` (2-6 words) + a `detail` that actually explains it.
  Never write bare keywords. How long `detail` is depends on the slide kind:
  * `bullets` slides are text-only slides, so they must be RICH: 5-6 points, and each
    `detail` is 2-3 full sentences (35-50 words) covering what it is, why it matters,
    and a concrete example, figure or real-world application. A bullets slide should
    fill 10-15 lines of text in total.
  * `cards` slides: 3-4 points, each `detail` 1-2 sentences (15-30 words).
  * `chart` / `stats` insight points: 2-3 points, each `detail` 15-25 words.
- Add a `subtitle` (one framing sentence) to most content slides.
- Add a `callout` (key insight, striking example or takeaway, max 30 words) to about
  half of the slides - only when it adds something the points don't already say.
- `speaker_notes`: 3-5 natural sentences that expand on the slide.

STRUCTURE (logical flow)
- Slide 1: an introduction - what the topic is, why it matters, what the deck covers.
- Middle slides: build understanding step by step (background -> how it works ->
  types/examples/applications -> data/impact -> challenges -> future/solutions),
  adapting this to the topic.
- Last slide: a strong `bullets` conclusion ("Key Takeaways" or similar) with 5 rich
  points that summarise the main lessons, plus a memorable `callout`.
- Only use `section` dividers for decks of 10+ slides (at most 2 of them).

LAYOUT VARIETY - pick the kind that fits the content and never use the same kind on
more than 2 slides in a row:
- `bullets`: explanations, causes, how something works.
- `cards`: 3-4 parallel items - types, examples, applications, steps, pillars.
- `two_column`: comparisons - pros/cons, before/after, myths/facts, A vs B. Fill `left`
  and `right` (4-5 items each, full sentences of 12-22 words). `points` may be empty.
- `stats`: 2-4 headline figures in `stats` plus 2-3 `points` interpreting them, and a
  `source` with year.
- `chart`: numeric data over categories/time, plus 2-3 `points` with key insights.
Use `stats` and `chart` ONLY for real, well-established figures you are confident about,
always with a source and year. Never invent numbers; if unsure, use another kind.
Pie charts have exactly one series. Each series has one value per category.

IMAGES
- Give `bullets` and `section` slides an `image_query`: a concrete, photographable 2-4
  word search term (e.g. "solar panel farm", not "renewable energy benefits").
  Use null for cards, two_column, stats and chart slides.

THEME
- Design `theme` to match the user's style description: strong contrast between
  background and text, an accent colour for highlights. Hex colours like "#1A2B3C" only.
"""


class ContentError(RuntimeError):
    """A problem we can show to the user as a friendly message."""


def _models_to_try() -> List[str]:
    preferred = os.getenv("GEMINI_MODEL", "").strip()
    models = ([preferred] if preferred else []) + FALLBACK_MODELS
    return list(dict.fromkeys(models))  # remove duplicates, keep order


def generate_deck(topic: str, theme: str, num_slides: int, audience: str,
                  api_key: Optional[str] = None,
                  log: Callable[[str], None] = print) -> Deck:
    key = api_key or os.getenv("GEMINI_API_KEY")
    if not key:
        raise ContentError("No Gemini API key set. Get a free one at aistudio.google.com/apikey")
    client = genai.Client(api_key=key)

    user_msg = (
        f"Topic: {topic}\n"
        f"Audience: {audience}\n"
        f"Visual style requested: {theme}\n"
        f"Number of content slides (excluding the cover): {num_slides}"
    )
    config = types.GenerateContentConfig(
        system_instruction=SYSTEM_PROMPT,
        response_mime_type="application/json",
        response_json_schema=Deck.model_json_schema(),
    )

    problems = set()  # what went wrong, used for the final error message
    for model in _models_to_try():
        for attempt in range(RETRIES_PER_MODEL):
            try:
                response = client.models.generate_content(
                    model=model, contents=user_msg, config=config)
            except errors.APIError as e:
                if e.code in (401, 403) or "API_KEY_INVALID" in str(e):
                    raise ContentError("Your Gemini API key was rejected. "
                                       "Check it at aistudio.google.com/apikey") from e
                if e.code in (500, 502, 503, 504):
                    problems.add("busy")
                    if attempt + 1 < RETRIES_PER_MODEL:
                        log(f"{model} is busy, retrying...")
                        time.sleep(RETRY_WAIT_SECONDS * (attempt + 1))
                        continue
                    log(f"{model} is busy, trying another model...")
                elif e.code == 429:
                    problems.add("quota")
                    log(f"{model} free quota used up, trying another model...")
                else:  # 400 / 404: this model can't handle the request - skip it
                    problems.add("other")
                    log(f"{model} not usable (error {e.code}), trying another model...")
                break  # next model

            # Got a response - check it's a valid deck
            try:
                if not response.text:
                    raise ValueError("empty response")
                deck = Deck.model_validate_json(response.text)
                log(f"Slides written by {model}.")
                return deck
            except (ValueError, ValidationError):
                problems.add("format")
                log(f"{model} returned an unusable answer, trying another model...")
                break  # next model

    if problems == {"quota"}:
        raise ContentError("Today's free AI quota is used up. Try again later, "
                           "or paste your own free Gemini key.")
    if "busy" in problems:
        raise ContentError("Google's free AI models are overloaded right now. "
                           "Please try again in a minute or two.")
    raise ContentError("Couldn't generate the slides. Try rewording your topic, "
                       "or try again in a minute.")
