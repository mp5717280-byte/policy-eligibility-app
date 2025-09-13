
# LIC Plan Eligibility — Step 1 & 2 (Mini App)

A shareable Streamlit app that replicates your in-chat logic:

- **Step 1:** From *LIC 10 Year Plan* → filter plans active on the purchase date (optionally only GREEN highlighted).
- **Step 2:** From those, filter by age eligibility using *Excel Sheet 2* (parses `days/months/years`).

## How to Run Locally
1. Install Python 3.10+
2. Create a virtual env (optional) and install requirements:
   ```bash
   pip install -r requirements.txt
   ```
3. Launch the app:
   ```bash
   streamlit run app.py
   ```
4. In your browser:
   - Upload **LIC 10 Year Plan** workbook
   - Upload **Excel Sheet 2**
   - Enter **Name, DOB, Purchase Date**
   - Click **Run Eligibility**
   - Download Step-1 and Step-2 CSVs

## Deploy & Share
### Option A — Streamlit Community Cloud (free)
1. Push these files to a public GitHub repo.
2. Go to https://streamlit.io/cloud → **New app** → select your repo → set **Main file path** to `app.py`.
3. Deploy. Share the URL with clients.

### Option B — Hugging Face Spaces (Gradio/Streamlit)
1. Create a new **Space** → choose **Streamlit**.
2. Upload these files (`app.py`, `requirements.txt`).
3. Deploy and share the Space URL.

### Option C — Internal Server
Run `streamlit run app.py` on a VM and expose the port via nginx.

## Notes
- For GREEN detection we check Excel cell fill color; if your file uses different shades or merged structures, uncheck **Use GREEN filter** in the sidebar to rely only on date ranges.
- Age parsing supports formats like `90 days`, `3 months`, `18 years`. Everything converts to **years** for comparison.
- If your Sheet 2 tabs vary (e.g., `Policy 1`, `Policy 2`, …), the app aggregates them automatically.
