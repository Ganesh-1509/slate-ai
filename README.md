# Slate - AI presentation maker (100% free)

Type a topic and a style, get a downloadable `.pptx` with themed slides,
editable charts and stock photos.

- Text: **Google Gemini** free tier
- Photos: **Pexels** free API
- Website: **Streamlit**, hosted free on **Streamlit Community Cloud**

## 1. Get the free keys
- Gemini: https://aistudio.google.com/apikey (sign in with Google, click "Create API key")
- Pexels: https://www.pexels.com/api/ (sign up, request a key; optional, no photos without it)

## 2. Run on your computer (VS Code terminal)
```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env     # then paste your keys into .env
python test_build.py       # offline test, makes output\test-*.pptx
streamlit run app.py       # opens the website at http://localhost:8501
```
Command-line version: `python cli.py "Your topic" --theme dark --slides 8`

## 3. Put it on GitHub
```powershell
git init
git add .
git commit -m "AI PowerPoint maker"
```
Create an empty repo on github.com, then run the two `git remote add` / `git push` lines it shows you.
`.env` is in `.gitignore`, so your keys are NOT uploaded.

## 4. Deploy free on Streamlit Community Cloud
1. Go to https://share.streamlit.io and sign in with GitHub.
2. Click **Create app** and choose your repo, branch `main`, main file `app.py`.
3. Open **Advanced settings -> Secrets** and paste:
   ```
   GEMINI_API_KEY = "your-gemini-key"
   PEXELS_API_KEY = "your-pexels-key"
   ```
4. Click **Deploy**. You get a public link like `https://your-app.streamlit.app`.

## Use your own template
Pick **Your template** under Style and upload a `.pptx` or `.potx` file (max 25 MB).
Its layouts, backgrounds, logos, fonts and colours are used for every slide; any
slides already in the file are replaced. Command line: `python cli.py "Topic" --template my.pptx`

## Good to know
- Free Gemini quota is shared by everyone using your site. When it runs out, users see
  a message and can paste their own free key (it is not stored).
- On the free tier, Google may use the prompts to improve its products. Don't enter private info.
- Free Streamlit apps go to sleep when unused; the first visitor wakes them up (~30 s).
- If Google renames the model, set `GEMINI_MODEL` in `.env` / Secrets to a current free model
  from https://aistudio.google.com/docs/pricing.
- Statistics come from the AI's knowledge, not a live web search: check them against the listed source.
