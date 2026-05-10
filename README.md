# Cover + Density Data Cleaner

A Streamlit app for cleaning vegetation cover and density data from CSV or Excel files.

## Features

### Cover Data mode

Use this mode for point-intercept cover files with columns such as:

- `YEAR` or `DATE`
- `BLOCK`
- `PLOT`
- optional `TRANSECT`
- `BASAL`
- optional `1ST FOLIAR`, `2ND FOLIAR`, `3RD FOLIAR`

The app follows the same logic as the cover program:

1. Clean `BLOCK` and `PLOT`.
2. Choose exactly one hit per row using this priority:
   - `1ST FOLIAR`
   - `2ND FOLIAR`
   - `3RD FOLIAR`
   - `BASAL`
   - `NO_HIT` or `0` for completely empty rows, depending on the option selected in the app
3. Count hits by `BLOCK + PLOT + HIT`.
4. Force every block to include `BN`, `BS`, `UN`, and `US` rows.
5. Export only the clean sheets:
   - `Cover Proportion`
   - `Cover Percent`

The output layout matches the cleaned example style:

`YEAR, BLOCK, PLOT, FIRE, RODENTS, species/category columns..., TOTAL_HITS, ROW_SUM`

### Density Data mode

Use this mode for density sheets with columns such as:

- `DATE` or `YEAR`
- `BLOCK`
- `PLOT`
- `SPECIES` or `PLANT SPECIES`
- count columns `1` through `12`
- optional `TRANSECT`

The density mode now follows the original density calculator structure:

1. Sum count columns `1` through `12` for each species row.
2. Add species totals by `BLOCK + PLOT + SPECIES`.
3. Output rows in this order: all `BN` blocks, all `BS` blocks, all `UN` blocks, all `US` blocks.
4. Export one sheet named `Species Totals Wide`.

Blank cells become `0`. Asterisks are ignored completely. For example, `2*` becomes `2`, not `10`, and `*` by itself becomes `0`. There is no `*5` multiplier.

The output layout matches the provided density example:

`DATE, BLOCK, PLOT, FIRE, RODENTS, species columns...`

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy

Upload `app.py`, `requirements.txt`, and `README.md` to a GitHub repository, then deploy the repo on Streamlit Community Cloud.
