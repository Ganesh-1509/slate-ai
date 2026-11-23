"""Command-line version (for testing on your own computer).

Example:
    python cli.py "Electric vehicles in India" --theme dark --slides 8
    python cli.py "History of the Internet" --theme "retro, warm browns, serif fonts"
"""
import argparse
import sys
from pathlib import Path

from dotenv import load_dotenv

from pptgen import ContentError, make_presentation


def main():
    load_dotenv()  # loads GEMINI_API_KEY / PEXELS_API_KEY from .env
    ap = argparse.ArgumentParser(description="Generate a PowerPoint from a topic.")
    ap.add_argument("topic", help="What the presentation is about (be detailed!)")
    ap.add_argument("--theme", default="minimal",
                    help="minimal | dark | lively | academic, or describe your own style")
    ap.add_argument("--slides", type=int, default=8, help="Number of content slides")
    ap.add_argument("--audience", default="general audience")
    ap.add_argument("--out", default="output", help="Output folder")
    args = ap.parse_args()

    try:
        name, data = make_presentation(args.topic, args.theme, args.slides, args.audience)
    except ContentError as e:
        sys.exit(f"Error: {e}")

    out = Path(args.out) / name
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(data)
    print(f"\nDone! Saved to: {out.resolve()}")


if __name__ == "__main__":
    main()
