# AEM Translation URL Builder

Streamlit app for Thermo Fisher AEM translation URL preparation.

## Deploy on Streamlit Community Cloud
1. Create a GitHub repository.
2. Upload `app.py` and `requirements.txt` (README is optional).
3. In Streamlit Community Cloud, create an app from the repo and set the main file to `app.py`.

## Behavior
- Reads **only the first worksheet** of the uploaded Excel workbook.
- Finds the `URL in AEM` header automatically.
- Detects language columns and rows marked `x`/`X`.
- Generates Original, Local Author, Regional Author, Regional Sites/Rollout, and Global Author URLs.
- Adds language flags and wraps URL text in the output workbook.
- Includes a `Locale Map` sheet for reference.

## Locale roots
- Europe: de-DE, es-ES, fr-FR
- Japan: ja-JP
- IPAC: ko-KR, zh-TW
- Greater China: zh-CN
- Latin America: pt-BR and es-LATAM (Argentina, Chile, Mexico)
