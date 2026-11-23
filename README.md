# PDF to Word Converter: CortexLabs by Zaheer
https://pdf2doc-tool.vercel.app/

A single-file Flask web utility (`app.py`) that converts PDF documents into editable Word (`.docx`) files with drag-and-drop upload, light/dark theme switching, and page/word count extraction feedback.

## Features

- **Drag & Drop + File Picker:** Upload `.pdf` files directly in the browser with instant file size feedback.
- **Editable `.docx` Generation:** Extracts text page-by-page into clean Word documents using `pypdf` and `python-docx`.
- **Conversion Readout:** Displays extracted page count and word count alongside status states.
- **Light & Dark Themes:** Switch themes on the fly with smooth transitions.

## Tech Stack

- **Backend:** Python, Flask, `pypdf`, `python-docx`
- **Frontend:** Embedded HTML5, CSS3, Vanilla JavaScript in `app.py`

## Local Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/ZaheerSaqib/pdf-to-word-converter.git
   cd pdf-to-word-converter
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the application:
   ```bash
   python app.py
   ```

4. Open `http://127.0.0.1:5000` in your browser.

## Author

**CortexLabs by Zaheer** — [GitHub @ZaheerSaqib](https://github.com/ZaheerSaqib)
