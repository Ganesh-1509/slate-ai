"""Data shapes shared by every part of the app.

The AI is asked to return JSON that matches `Deck` exactly (structured output),
so the builder never has to guess what the AI meant.
"""
from typing import List, Literal, Optional

from pydantic import BaseModel, Field

FontName = Literal[
    "Calibri", "Arial", "Segoe UI", "Verdana", "Trebuchet MS",
    "Georgia", "Cambria", "Garamond", "Century Gothic",
]


class Series(BaseModel):
    name: str
    values: List[float]


class ChartSpec(BaseModel):
    chart_type: Literal["column", "bar", "line", "pie"]
    title: str
    categories: List[str] = Field(description="X-axis labels, e.g. years or countries")
    series: List[Series] = Field(description="Each series has one value per category")
    source: str = Field(description="Where the numbers come from, with year")


class Point(BaseModel):
    heading: str = Field(description="Bold lead-in of 2-6 words")
    detail: str = Field(
        description="Explanation with a concrete example, figure, application or "
                    "consequence. bullets slides: 2-3 sentences (35-50 words). "
                    "cards: 15-30 words. chart/stats insights: 15-25 words."
    )


class Column(BaseModel):
    heading: str = Field(description="Column title, e.g. 'Advantages' or 'Before'")
    points: List[str] = Field(description="4-5 items, each a full sentence of 12-22 words")


class Stat(BaseModel):
    value: str = Field(description="Short figure such as '68%', '$1.2 trillion', '3x', '2030'")
    label: str = Field(description="What the figure means, 5-12 words")


class SlideSpec(BaseModel):
    kind: Literal["bullets", "cards", "two_column", "stats", "chart", "section"] = Field(
        description="bullets: explained points. cards: 3-4 parallel ideas (types, examples, "
                    "applications, steps). two_column: comparison (pros/cons, before/after). "
                    "stats: 2-4 key real figures. chart: real numeric data. section: divider."
    )
    title: str
    subtitle: Optional[str] = Field(
        default=None, description="One sentence (max 20 words) framing the slide's message")
    points: List[Point] = Field(
        default_factory=list,
        description="bullets: 5-6 rich points. cards: 3-4. chart/stats: 2-3 key insights. "
                    "two_column/section: may be empty")
    left: Optional[Column] = Field(default=None, description="two_column only")
    right: Optional[Column] = Field(default=None, description="two_column only")
    stats: Optional[List[Stat]] = Field(default=None, description="stats only: 2-4 figures")
    callout: Optional[str] = Field(
        default=None,
        description="Optional key insight, real-world example or takeaway (max 30 words)")
    image_query: Optional[str] = Field(
        default=None, description="2-4 word stock-photo search term, or null for no image")
    chart: Optional[ChartSpec] = Field(default=None, description="chart only")
    source: Optional[str] = Field(default=None, description="stats only: source with year")
    speaker_notes: str = Field(default="", description="3-5 sentences the presenter can say")


class ThemeSpec(BaseModel):
    """Colours as hex strings like '#1A2B3C'."""
    background: str
    title_color: str
    text_color: str
    accent: str
    accent2: str
    title_font: FontName
    body_font: FontName


class Deck(BaseModel):
    title: str
    subtitle: str
    cover_image_query: Optional[str]
    theme: ThemeSpec
    slides: List[SlideSpec]
