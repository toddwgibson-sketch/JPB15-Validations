# LV Portal Validation Suite

Unified Streamlit toolkit for processing LV Portal and Slack validation reports.

## Tools Included

| Tool | Purpose | Key Features |
|------|---------|--------------|
| **Compute Link Validator** | GPU/Compute to T0 validation | Ghost host detection, Compute Optics, Summary dashboard |
| **QFAB / T1→T0 Validator** | Newer LV Portal T1-T0 exports | Per-channel Optics, strong cutsheet fallback |
| **Slack Report Highlighter** | Classic Slack validation reports | Full LLDP + recurring issue detection |

## Running Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploying to Streamlit Community Cloud

1. Push this folder to a **public** GitHub repository
2. Go to [https://share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub account
4. Select the repo and set **Main file path** to `app.py`
5. Deploy

## Adding Your Logo

1. Place your logo as `assets/logo.png` (recommended size: ~200×80 px or square)
2. The app will automatically display it in the header of every page

## Notes

- All processing happens in memory / temporary files
- Nothing is stored on the server after you download the output
- Multiple cutsheets are supported in all tools (highly recommended)