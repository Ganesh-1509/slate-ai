"""Built-in themes. If the user's theme text names one of these, we use it;
otherwise we use the theme Claude designed from the user's description."""
from .models import ThemeSpec

PRESETS = {
    "minimal": ThemeSpec(
        background="#FFFFFF", title_color="#111111", text_color="#333333",
        accent="#2563EB", accent2="#94A3B8",
        title_font="Segoe UI", body_font="Segoe UI",
    ),
    "dark": ThemeSpec(
        background="#0F172A", title_color="#F8FAFC", text_color="#CBD5E1",
        accent="#38BDF8", accent2="#A78BFA",
        title_font="Segoe UI", body_font="Segoe UI",
    ),
    "lively": ThemeSpec(
        background="#FFF7ED", title_color="#9A3412", text_color="#431407",
        accent="#F97316", accent2="#14B8A6",
        title_font="Trebuchet MS", body_font="Verdana",
    ),
    "academic": ThemeSpec(
        background="#FBFAF7", title_color="#1E3A5F", text_color="#2D2D2D",
        accent="#8B1E3F", accent2="#1E3A5F",
        title_font="Georgia", body_font="Cambria",
    ),
}


def pick_theme(user_theme: str, ai_theme: ThemeSpec) -> ThemeSpec:
    """Exact preset name wins; anything else ('dark with neon green') uses the AI theme."""
    key = user_theme.strip().lower()
    return PRESETS.get(key, ai_theme)
