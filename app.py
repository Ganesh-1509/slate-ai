"""Slate - the website. Run locally with:  streamlit run app.py"""
import os

import streamlit as st
from dotenv import load_dotenv

from pptgen import PRESETS, ContentError, make_presentation

APP_NAME = "Slate"
MAX_SLIDES = 15
CUSTOM = "Custom"

load_dotenv()  # local: read keys from .env
# Streamlit Cloud: keys live in the app's "Secrets" settings
for name in ("GEMINI_API_KEY", "PEXELS_API_KEY", "GEMINI_MODEL"):
    try:
        if name in st.secrets and not os.getenv(name):
            os.environ[name] = str(st.secrets[name])
    except Exception:  # no secrets file locally - that's fine
        pass

st.set_page_config(page_title=APP_NAME, page_icon=":material/slideshow:", layout="centered")

# ---------------------------------------------------------------- header
st.title(APP_NAME)
st.caption("Describe a topic. Get a finished, presentation-ready deck with charts, "
           "photos and speaker notes.")
st.space("small")

# ---------------------------------------------------------------- inputs
with st.container(border=True):
    topic = st.text_area(
        "What's your presentation about?",
        placeholder="e.g. Growth of electric vehicles in India 2015-2025: sales, "
                    "government policy, charging infrastructure and challenges",
        height=130,
        key="topic",
    )
    st.caption("The more detail you give, the better the slides.")

    preset = st.segmented_control(
        "Style",
        [p.capitalize() for p in PRESETS] + [CUSTOM],
        default="Minimal",
        required=True,
        key="style",
    )
    custom_theme = ""
    if preset == CUSTOM:
        custom_theme = st.text_input(
            "Describe your style",
            placeholder="e.g. dark navy with gold accents, elegant serif fonts",
            key="custom_style",
        )

    left, right = st.columns(2, vertical_alignment="bottom")
    with left:
        audience = st.text_input("Audience", value="General audience", key="audience")
    with right:
        num_slides = st.slider("Slides", 3, MAX_SLIDES, 8, key="slides")

    with st.expander("Use your own Gemini key", icon=":material/key:"):
        st.caption("Only needed if the shared free quota runs out. Get one free at "
                   "[aistudio.google.com/apikey](https://aistudio.google.com/apikey). "
                   "It's used for this request only and never stored.")
        user_key = st.text_input("Gemini API key", type="password",
                                 label_visibility="collapsed", placeholder="Paste your key",
                                 key="user_key")

    generate = st.button("Generate deck", type="primary", icon=":material/auto_awesome:",
                         width="stretch", key="generate")

# ---------------------------------------------------------------- generation
if generate:
    theme = custom_theme.strip() if preset == CUSTOM else (preset or "minimal").lower()
    if len(topic.strip()) < 5:
        st.error("Describe your topic first.", icon=":material/edit_note:")
    elif not theme:
        st.error("Describe your custom style, or pick one of the presets.",
                 icon=":material/palette:")
    else:
        with st.status("Creating your deck...", expanded=True) as status:
            try:
                file_name, data = make_presentation(
                    topic.strip(), theme, num_slides, audience.strip() or "general audience",
                    api_key=user_key.strip() or None, log=st.caption,
                )
                st.session_state["result"] = (file_name, data)
                status.update(label="Your deck is ready", state="complete", expanded=False)
            except ContentError as e:
                status.update(label="Couldn't create the deck", state="error")
                st.error(str(e), icon=":material/error:")
            except Exception as e:
                status.update(label="Couldn't create the deck", state="error")
                st.error("Unexpected error while building the presentation. Please try again.",
                         icon=":material/error:")
                with st.expander("Error details"):
                    st.exception(e)

# ---------------------------------------------------------------- result
if "result" in st.session_state:
    file_name, data = st.session_state["result"]
    with st.container(border=True):
        st.subheader("Ready to download", icon=":material/check_circle:")
        st.caption(f"{file_name}  ·  {len(data) / 1024:.0f} KB  ·  opens in PowerPoint, "
                   "Google Slides and Keynote")
        st.download_button(
            "Download .pptx", data=data, file_name=file_name, type="primary",
            icon=":material/download:", width="stretch",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    st.caption("Double-check statistics against their listed sources before presenting.")

# ---------------------------------------------------------------- footer
st.space("large")
st.caption("Text by Google Gemini  ·  Photos by [Pexels](https://www.pexels.com) "
           "(credits in each slide's speaker notes)", text_alignment="center")
