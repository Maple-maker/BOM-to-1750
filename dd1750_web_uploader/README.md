# DD1750 Web Uploader (Streamlit)

Upload a BOM (PDF or Excel) + a flat DD1750 template PDF (`blank_flat.pdf`) and generate:
- a **multi-page DD1750** PDF (18 items/page)
- an **AUDIT CSV** (flags suspicious quantities, typically OCR noise)
- (Optional) merge many DD1750 PDFs into one file

## Local Run (macOS)
```bash
python3 -m pip install -r requirements.txt
brew install tesseract   # needed for scanned PDFs
streamlit run app.py
```

## Deploy (recommended): Docker
```bash
docker build -t dd1750-uploader .
docker run -p 8501:8501 dd1750-uploader
```
Open http://localhost:8501
