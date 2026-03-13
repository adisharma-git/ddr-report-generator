# 🏗️ DDR Report Generator — AI-Powered Detailed Diagnostic Reports

An AI system that reads site inspection documents and thermal imaging reports, then generates a professional **Detailed Diagnostic Report (DDR)** with images, thermal analysis, severity assessment, and actionable recommendations.

> **Built for:** AI Generalist Applied AI Builder Assignment  
> **AI Model:** Google Gemini 2.0 Flash (Free Tier, 1M+ token context window)  
> **Stack:** Python, Flask, PyMuPDF, python-docx, ReportLab

---

## 🎯 What It Does

1. **Reads** two PDF documents — an Inspection Report (site observations + photos) and a Thermal Imaging Report (temperature readings + thermal images)
2. **Extracts** text content and all embedded images from both PDFs using PyMuPDF
3. **Analyzes** combined data using Google Gemini AI to identify issues, root causes, severity, and recommendations
4. **Generates** a structured DDR report in three formats:
   - **DOCX** (Word document with images, tables, thermal data)
   - **PDF** (Professional formatted report)
   - **JSON** (Structured data for programmatic use)
5. **Displays** an interactive web preview with tabbed sections, image galleries, and downloadable files

---

## 📋 DDR Output Structure

The generated report contains all required sections:

| Section | Description |
|---------|-------------|
| **Property Issue Summary** | Comprehensive overview of all identified issues |
| **Area-wise Observations** | Negative side (damage) + Positive side (source) per area, with photo/thermal references |
| **Probable Root Causes** | Detailed root cause analysis with affected areas |
| **Severity Assessment** | Critical/High/Moderate/Low rating with reasoning |
| **Recommended Actions** | Prioritized repair actions with urgency timelines |
| **Thermal Analysis** | Temperature readings, differentials, moisture interpretation |
| **Summary Table** | Impacted area ↔ Exposed area correlation table |
| **Additional Notes** | Extra observations and warnings |
| **Missing/Unclear Info** | Explicitly flags "Not Available" data |
| **Appendix** | Extracted inspection photographs + thermal images |

---

## 🛠️ Tech Stack

| Component | Technology | Why |
|-----------|-----------|-----|
| **AI Model** | Google Gemini 2.0 Flash | Free tier, 1M+ token window, excellent accuracy, multimodal |
| **Backend** | Flask (Python) | Lightweight, production-ready, simple deployment |
| **PDF Processing** | PyMuPDF (fitz) | Fast text + image extraction from PDFs |
| **DOCX Generation** | python-docx | Professional Word documents with tables, images, styling |
| **PDF Generation** | ReportLab | Programmatic PDF creation with custom layouts |
| **Image Processing** | Pillow | Image validation and format handling |
| **Frontend** | Vanilla HTML/CSS/JS | Zero dependencies, fast load, responsive |

---

## 🚀 Setup & Run

### Prerequisites
- Python 3.10+
- Google Gemini API key (free at [aistudio.google.com](https://aistudio.google.com))

### Installation

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/ddr-report-generator.git
cd ddr-report-generator

# Install dependencies
pip install -r requirements.txt

# Run the application
python app.py
```

The app will start at **http://localhost:5000**

### Usage

1. Open `http://localhost:5000` in your browser
2. Enter your Google Gemini API key (free tier works)
3. Upload the Inspection Report PDF
4. Upload the Thermal Images PDF
5. Click "Generate DDR Report"
6. Review the interactive report preview
7. Download DOCX, PDF, or JSON outputs

---

## 🏗️ Architecture

```
User uploads 2 PDFs
        │
        ▼
┌─────────────────────┐
│   PDF Extraction     │  PyMuPDF extracts text + images
│   (PyMuPDF/fitz)     │  from both documents
└────────┬────────────┘
         │
         ▼
┌─────────────────────┐
│   AI Analysis        │  Gemini processes combined text,
│   (Google Gemini)    │  generates structured DDR JSON
└────────┬────────────┘
         │
         ▼
┌─────────────────────┐
│   Report Generation  │  DOCX (python-docx) + PDF (ReportLab)
│   + Image Mapping    │  with extracted images placed in sections
└────────┬────────────┘
         │
         ▼
┌─────────────────────┐
│   Web Preview        │  Interactive tabbed UI with
│   + Downloads        │  galleries, tables, downloads
└─────────────────────┘
```

---

## 🧠 How It Handles Edge Cases

| Scenario | Handling |
|----------|----------|
| **Missing data** | AI explicitly writes "Not Available" |
| **Conflicting info** | AI mentions the conflict in the report |
| **No thermal data for an area** | Reports "Thermal data not available for this area" |
| **Small/irrelevant images** | Filtered out (< 50px threshold) |
| **Large PDFs** | PyMuPDF handles efficiently; Gemini's 1M token window accommodates large documents |
| **API failures** | Clear error messages with retry guidance |
| **Duplicate observations** | AI deduplicates during analysis phase |

---

## 📁 Project Structure

```
ddr-report-generator/
├── app.py                  # Main Flask application
├── templates/
│   └── index.html          # Frontend UI
├── requirements.txt        # Python dependencies
├── README.md               # This file
├── sample_inputs/          # Sample input documents (for testing)
│   ├── Inspection_Report.pdf
│   └── Thermal_Images.pdf
└── sample_outputs/         # Example generated reports
    ├── DDR_Report.docx
    ├── DDR_Report.pdf
    └── DDR_Data.json
```

---

## ⚠️ Limitations

1. **Image-to-Area Mapping**: Currently maps images sequentially by page order. A more sophisticated approach would use OCR + spatial analysis to precisely match each photo to its labeled area in the source PDF.
2. **Thermal Image Correlation**: Maps thermal images to inspection areas based on text analysis. Physical overlap detection (using GPS/coordinates) would improve accuracy.
3. **Free Tier Rate Limits**: Google Gemini free tier has rate limits (15 RPM). For production use, a paid API key is recommended.
4. **PDF Complexity**: Highly complex PDFs with unusual layouts may have imperfect text extraction.
5. **Single Session**: Each generation creates a new session. No persistent storage across browser refreshes.

---

## 🔮 How I Would Improve It

1. **Vision API Integration**: Send actual images to Gemini's vision model for direct visual analysis (crack severity, moisture pattern recognition)
2. **Multi-model Pipeline**: Use a smaller model for extraction + a larger model for reasoning/recommendations
3. **Template System**: Allow custom DDR templates per client/company with configurable sections
4. **Batch Processing**: Support processing multiple properties in one go
5. **Database Backend**: Store reports in PostgreSQL for historical comparison and trend analysis
6. **OCR Enhancement**: Add Tesseract OCR for scanned/handwritten inspection notes
7. **Automated Image Labeling**: Use object detection to auto-tag images (crack, dampness, mold, etc.)
8. **Export to Google Docs**: Direct integration with Google Workspace APIs
9. **Comparison Mode**: Compare current DDR with previous reports to track deterioration/improvement
10. **WhatsApp/Email Delivery**: Auto-send reports to clients

---

## 📝 Design Decisions

- **Gemini over GPT**: Chosen for the free tier with 1M+ token context window — can process entire inspection reports without chunking
- **Flask over Streamlit**: More control over UI/UX, easier to deploy as a standalone web app, better for the demo link requirement
- **PyMuPDF over pdfplumber**: Faster image extraction and better handling of complex PDF layouts
- **Separate DOCX + PDF**: DOCX for client editing, PDF for final delivery — matching real-world workflow
- **JSON intermediate**: Structured JSON output enables downstream integrations and programmatic access

---

## 📄 License

MIT License — Free to use, modify, and distribute.
