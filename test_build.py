"""Offline test: builds a sample deck in every theme. No API keys needed.

    python test_build.py        -> output\test-<theme>.pptx
"""
from pathlib import Path

from pptgen.builder import DeckBuilder
from pptgen.models import ChartSpec, Column, Deck, Point, Series, SlideSpec, Stat
from pptgen.themes import PRESETS


def P(h, d):
    return Point(heading=h, detail=d)


SAMPLE = Deck(
    title="Solar Energy for Indian Households",
    subtitle="How rooftop solar cuts bills, supports the grid and what it takes to get started",
    cover_image_query="rooftop solar panels",
    theme=PRESETS["minimal"],
    slides=[
        SlideSpec(
            kind="bullets", title="Why Solar Matters Now",
            subtitle="Falling costs and rising demand make rooftop solar a practical choice.",
            image_query="solar panels house",
            points=[
                P("Rising electricity demand", "Household consumption keeps climbing as more homes add ACs, fridges and electric vehicles."),
                P("Panel prices have fallen", "Module costs dropped sharply over the last decade, shortening payback periods for families."),
                P("Abundant sunshine", "Most of India receives strong sunlight for 250-300 days a year, ideal for generation."),
                P("Government support", "Central and state subsidies reduce the upfront cost of residential rooftop systems."),
                P("Energy independence", "Generating power at home protects families from tariff hikes and outages."),
            ],
            speaker_notes="Start by framing why this is relevant today.",
        ),
        SlideSpec(
            kind="cards", title="Types of Rooftop Solar Systems",
            subtitle="Choose a system based on grid reliability and budget.",
            points=[
                P("On-grid", "Connected to the utility grid; excess power is exported through net metering. Lowest cost, but shuts off during outages."),
                P("Off-grid", "Runs independently with batteries. Suits remote homes with unreliable supply, but batteries raise the cost."),
                P("Hybrid", "Combines grid connection with battery backup, giving savings plus power during outages at a moderate premium."),
            ],
            callout="Most city households choose on-grid systems because net metering offers the fastest payback.",
            speaker_notes="Walk through each type.",
        ),
        SlideSpec(
            kind="two_column", title="Benefits vs. Challenges",
            left=Column(heading="Benefits", points=[
                "Electricity bills can fall substantially once the system is running",
                "Very low maintenance: occasional panel cleaning and inspection",
                "Panels typically carry 25-year performance warranties",
                "Reduces household carbon emissions every year",
            ]),
            right=Column(heading="Challenges", points=[
                "Upfront cost is still significant without subsidies",
                "Requires shade-free roof space facing the sun",
                "Net-metering approval can take weeks in some states",
                "Batteries for backup add cost and need replacement",
            ]),
            speaker_notes="Balance the picture.",
        ),
        SlideSpec(
            kind="stats", title="Solar in Numbers",
            subtitle="Illustrative figures for testing the layout.",
            stats=[Stat(value="25 yrs", label="Typical panel performance warranty"),
                   Stat(value="4-6 yrs", label="Common payback period with subsidy"),
                   Stat(value="300", label="Sunny days per year in many regions")],
            points=[P("Long life", "Systems keep producing for decades, long after paying for themselves."),
                    P("Quick payback", "Subsidies and net metering shorten the time to recover the investment.")],
            source="Sample data for layout testing",
            speaker_notes="Highlight the key numbers.",
        ),
        SlideSpec(
            kind="chart", title="Example: Monthly Generation",
            points=[P("Summer peak", "Generation is highest in April and May when sunlight is strongest."),
                    P("Monsoon dip", "Cloud cover in July and August reduces output noticeably."),
                    P("Plan around it", "Net metering credits from summer offset the weaker monsoon months.")],
            chart=ChartSpec(chart_type="column", title="kWh",
                            categories=["Jan", "Mar", "May", "Jul", "Sep", "Nov"],
                            series=[Series(name="kWh", values=[380, 450, 520, 310, 360, 390])],
                            source="Sample data for layout testing"),
            speaker_notes="Explain seasonality.",
        ),
        SlideSpec(
            kind="bullets", title="Key Takeaways",
            points=[
                P("Solar pays for itself", "Most households recover their investment within a few years and save for decades."),
                P("Pick the right system", "On-grid suits reliable cities; hybrid or off-grid suits outage-prone areas."),
                P("Use subsidies", "Government schemes cut upfront costs, so check eligibility before buying."),
                P("Plan the roof", "A shade-free, south-facing roof gives the best output."),
            ],
            callout="Rooftop solar turns every sunny day into savings - the best time to start is now.",
            speaker_notes="Close with a summary.",
        ),
    ],
)

if __name__ == "__main__":
    Path("output").mkdir(exist_ok=True)
    for name, theme in PRESETS.items():
        out = Path("output") / f"test-{name}.pptx"
        out.write_bytes(DeckBuilder(theme).build(SAMPLE.model_copy(deep=True)))
        print("wrote", out)
