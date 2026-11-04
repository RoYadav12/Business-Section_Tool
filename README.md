# Fill Deck Generator

Automatically fills a PowerPoint template using data from a RefractOne Excel workbook.

## What it does

Given:
- a RefractOne-style Excel workbook (`.xlsm`/`.xlsx`) with `Overview`, `IT Vendor`,
  and `Business Themes` sheets (optionally `Tech Themes`, `Key Triggers & Recent
  Developments`, `Sustainability & ESG Themes` and `Challenges & Growth
  Opportunities`), and
- a matching PPTX template with `xxxx` placeholders,

it produces a filled `.pptx` deck, auto-duplicating slides (business themes,
tech themes, sustainability themes, challenges, vendor cards, key developments)
when a company has more items than the template's base layout has room for.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints, upload your Excel workbook and PPTX
template, enter a company name, and click **Generate Presentation**.

## Run from the command line (no UI)

```bash
python3 fill_deck.py --excel COMPANY.xlsm --pptx Sample.pptx \
    --company "Company Name" --output Filled_Deck.pptx
```

## Deploy on Streamlit Community Cloud

1. Push this folder to a GitHub repo.
2. Go to [share.streamlit.io](https://share.streamlit.io), click **New app**,
   and point it at this repo with `app.py` as the main file.
3. Streamlit Cloud installs `requirements.txt` automatically — no extra config needed.

## Files

| File | Purpose |
|---|---|
| `app.py` | Streamlit UI: upload files, call `fill_deck.generate_ppt`, offer the result for download |
| `fill_deck.py` | All the parsing/slide-filling logic, plus a CLI (`python3 fill_deck.py ...`) |
| `requirements.txt` | Python dependencies |
| `Sample.pptx` / `New_Sample_Business.pptx` | Example templates (older / current, with Challenges and Sustainability slides) |
| `RefractOne_Business_Demo_Final.xlsm` / `RefractOne_Business_Final_Template_test.xlsm` | Example workbooks (older / current) |

## Notes

- Wherever the Excel data doesn't contain a value the template asks for, the
  script writes `N/A` or a flagged placeholder (`xxxx`) instead of guessing —
  search the output deck for `xxxx` to find anything that still needs a human touch.
- The first six slides (cover to IT Vendor Summary) are filled by position. Every
  slide after that is found by its **title** (`Core Business Themes`, `Technology
  Themes`, `Sustainability Themes`, `Challenges & Growth Prospects`, `Key
  Developments`, `IT Vendors – Software / ICT Services`), so those slides can be
  added, removed or reordered in the template without code changes.
- Sheet formats: both the older layout (Business Themes with one row per supporting
  instance, Tech Themes as Headline / Commentary / Supporting Instances / Impact)
  and the current layout (one theme per row: Theme / Summary Intro / Key Points /
  Strategic Impact) are read. `Key Triggers` dates can be real dates or
  month text such as `October 2026`.
- If the workbook has no `Challenges & Growth Opportunities` or `Sustainability & ESG
  Themes` sheet, those slides are dropped from the output rather than left as
  `xxxx` slides.
- If the template contains two copies of the same slide family (e.g. two
  `Technology Themes` slides left over from editing), the **last** copy is used and
  the others are removed (see `KEEP_WHICH_DUPLICATE` in `fill_deck.py`).
- Challenges & Growth Prospects produces one slide per theme (Key Challenges on the
  left, Opportunities on the right).
- Long bullets shrink their font (down to 7.5pt) so text stays inside its panel.
- A structurally different PPTX template would need the shape-index mappings in
  `fill_deck.py` updated.
