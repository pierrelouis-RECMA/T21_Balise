# NBB Report Generator

Upload an Excel file → get a filled PowerPoint presentation.

## Files

```
app.py                          Flask web app + API
fill_template.py                Fills slides 1–6 with {{placeholders}}
generate_pptx_v2.py             Generates agency card slides 7+
T21_HK_Agencies_Glass_v13.pptx Template (slides 1–6, balises injectées)
requirements.txt
Dockerfile
```

## Deploy on Render.com

### Option A — Docker (recommended)

1. Push this folder to a GitHub repo (private is fine)
2. Render Dashboard → **New** → **Web Service**
3. Connect your GitHub repo
4. Settings:
   - **Environment**: Docker
   - **Dockerfile path**: `Dockerfile`  (auto-detected)
   - **Instance type**: Free tier works; use Starter for faster boot
5. Click **Deploy**

Render auto-detects the `$PORT` env variable — no manual config needed.

### Option B — Native Python (no Docker)

1. Push to GitHub
2. Render → New → Web Service
3. Settings:
   - **Environment**: Python 3
   - **Build command**: `pip install -r requirements.txt`
   - **Start command**: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`

## Local test

```bash
pip install -r requirements.txt
python app.py
# → open http://localhost:5000
```

## Excel format required

| Column              | Type   | Required | Notes                                |
|---------------------|--------|----------|--------------------------------------|
| Agency              | string | ✅        | MINDSHARE, HAVAS MEDIA, etc.         |
| NewBiz              | string | ✅        | WIN / DEPARTURE / RETENTION          |
| Advertiser          | string | ✅        | Client / brand name                  |
| Integrated Spends   | float  | ✅        | $m — positive WIN, negative DEPART.  |
| Date of announcement| date   | optional | Shown on agency card                 |
| Incumbent           | string | optional | Previous agency                      |

## Adding a new market

Edit `AGENCY_GROUP` in both `fill_template.py` and `generate_pptx_v2.py`
to add the local agency names and their group (Publicis Media, Omnicom Media,
Dentsu, Havas Media Network, WPP Media, Independant).

## Slide structure

| Slide   | Content                                  | How filled              |
|---------|------------------------------------------|-------------------------|
| 1       | Cover / Sommaire                         | Static (template)       |
| 2       | Key Findings — top 4 + Key Takeaways     | fill_template.py        |
| 3       | TOP moves — wins / departures / retentions| fill_template.py       |
| 4       | Agency overview (bar chart + table ×14)  | fill_template.py        |
| 5       | Group overview (dynamic ranking)         | fill_template.py        |
| 6       | Retentions ranking                       | fill_template.py        |
| 7+      | Agency detail cards (4 per slide)        | generate_pptx_v2.py     |
