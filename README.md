# Bank XYZ — Branch Customer Experience Dashboard

Survey analytics for an Indonesian bank's branch customer-experience study
(anonymised as "Bank XYZ"). 1,730 respondents × 632 variables, collected across
branches, provinces and five service touchpoints.

The deliverable is a Streamlit dashboard for CX / branch stakeholders, driven by a
metadata layer that an analyst can re-shape in Excel without touching code.

```
data/Deka_project_dataset_BankXYZ.csv   raw survey export (semicolon-delimited, 2-row header)
pipeline.ipynb                          raw data -> metadata (the whole prep pipeline)
metadata/metadata.csv                   semantic layer: all 632 variables tagged
metadata/metadata_dashboard.{csv,xlsx}  what the app reads; the .xlsx is the editable control surface
dashboard.py                            the Streamlit app
smoke_test.py                           headless render check for every page
```

## Run it

```bash
# 1. regenerate the metadata (only needed if the data or the tagging rules change)
uv run --with pandas,openpyxl,nbconvert,nbclient,ipykernel \
    jupyter nbconvert --to notebook --execute --inplace pipeline.ipynb

# 2. serve the dashboard
uv run --with pandas,plotly,streamlit,openpyxl streamlit run dashboard.py

# 3. verify every page still renders (no browser needed)
uv run --with pandas,plotly,streamlit,openpyxl python smoke_test.py
```

The notebook is the only thing that writes `metadata/`. Set `DEKA_METADATA_OUT` to
write elsewhere when experimenting.

## Reading the numbers

Three properties of this dataset shape every chart in the app. They are worth knowing
before interpreting anything.

**1. Every attribute was asked twice, on two different scales.** Each touchpoint has an
*importance* battery (`1 SANGAT TIDAK PENTING … 6 SANGAT PENTING`) and a *satisfaction*
battery (`1 SANGAT TIDAK PUAS … 6 SANGAT PUAS`), with near-identical question text. The
metadata keeps them apart as `role=Importance` and `role=Atribut`, and links each pair
with a `pair_key`. **Never merge the two** — averaging them produces a number that
answers no question. `pipeline.ipynb` §2.3 shows the two scales side by side.

**2. Scores sit at the ceiling.** All 169 satisfaction attributes fall between 5.59 and
5.90 out of 6; the four sections are separated by 0.08 points, the eight touchpoints by
0.16. Absolute scores therefore carry almost no information. What does vary:

| Measure | Spread across the 169 satisfaction attributes |
|---|---|
| Mean score | 5.59 – 5.90 (σ 0.06) |
| Top-2-box % (answered 5 or 6) | 92.8% – 100% (σ 1.3) — **cannot discriminate** |
| Top-box % (answered 6) | 69.8% – 90.1% (σ 4.9) — 4.5× the relative spread of T2B |
| Importance − satisfaction gap | −0.18 to +0.13 — the actual priority signal |
| Emotion items | XYZ 1.58 vs competitor 2.18 on negatives — the widest gap in the study |

So the dashboard leads with **gap** and **Importance-Performance Analysis**, reports
**top-box** instead of the useless top-2-box, and zooms every axis to the real data
range. Charts that compare magnitude use **dot plots, not bars** — a bar whose baseline
is not zero overstates the differences.

**3. Base sizes differ by a factor of 25.** ATM attributes rest on ~1,600 respondents;
Customer Advisor and Sarana Elektronik on ~70, because of survey routing (n runs 68 to
1,730, and 35 of the 169 attributes sit under 100). Every section and sub-category
average is therefore **weighted by n**, the real base is shown on the KPI cards, and any
view containing an attribute with n < 100 raises a warning. Competitor KPIs (`E1B`,
`F1B`, `G1C`) come from a 546-respondent sub-sample — not the same base as the XYZ
figures, which is disclosed wherever the two appear together.

Missing cells (~42%) are structural — routing, the competitor subset, and `99`/`999`
"not applicable" codes. Nothing is ever imputed; metrics are computed on answered rows
only.

## Re-shaping the dashboard without code

Open `metadata/metadata_dashboard.xlsx` (sheet **Metadata**; the **Petunjuk** sheet has
the instructions in Indonesian) and edit:

- `include` — `1` shows the variable, `0` hides it
- `subgroup` — the sub-category a chart groups by
- `label` — the text on chart axes

Save, then re-run the app. The app prefers the `.xlsx` over the `.csv`. Do not rename
`variable`, swap `role=Atribut` with `role=Importance`, or edit `pair_key` — the app
filters on those exact values and pages will silently render empty.

See `CHANGES.md` for what changed in the current revision and why.
