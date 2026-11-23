"""Download free stock photos from Pexels (https://www.pexels.com/api/)."""
import hashlib
import os
from pathlib import Path
from typing import Optional, Tuple

import requests

CACHE_DIR = Path(".image_cache")

Photo = Tuple[Path, str]  # (local file, credit text)


def fetch_image(query: Optional[str]) -> Optional[Photo]:
    """Return (local path, 'Photo: Name on Pexels') for `query`, or None."""
    api_key = os.getenv("PEXELS_API_KEY")
    if not query or not api_key:
        return None

    CACHE_DIR.mkdir(exist_ok=True)
    stem = hashlib.md5(query.lower().encode()).hexdigest()
    target, credit_file = CACHE_DIR / f"{stem}.jpg", CACHE_DIR / f"{stem}.txt"
    if target.exists() and credit_file.exists():
        return target, credit_file.read_text(encoding="utf-8")

    try:
        r = requests.get(
            "https://api.pexels.com/v1/search",
            headers={"Authorization": api_key},
            params={"query": query, "per_page": 1, "orientation": "landscape"},
            timeout=15,
        )
        r.raise_for_status()
        photos = r.json().get("photos", [])
        if not photos:
            return None
        photo = photos[0]
        img = requests.get(photo["src"]["large"], timeout=30)
        img.raise_for_status()
        credit = f"Photo: {photo.get('photographer', 'Unknown')} on Pexels ({photo.get('url', '')})"
        target.write_bytes(img.content)
        credit_file.write_text(credit, encoding="utf-8")
        return target, credit
    except (requests.RequestException, KeyError, ValueError) as e:
        print(f"  ! image search failed for '{query}': {e}")
        return None
