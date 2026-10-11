# 📄 Offline PDF Utility

> **A modern, privacy-focused, offline PDF and document toolkit for Kali Linux.**

**Offline PDF Utility** is a desktop application designed to bring essential PDF and document operations into one simple, modern interface. It works completely offline, allowing users to process documents locally without uploading files to websites or cloud services.

Created by **AakashSingh988**.

---

## ✨ Features

### 🔄 Document Conversion

* 🖼️ JPG / PNG → PDF
* 📄 PDF → Images
* 📝 Word → PDF
* 📊 Excel → PDF
* 📽️ PowerPoint → PDF
* 📄 PDF → Word
* 📊 PDF → Excel
* 📽️ PDF → PowerPoint
* 📝 Text → PDF

### 🛠️ PDF Tools

* 🔗 Merge multiple PDFs
* ✂️ Split PDF
* 🔄 Rotate PDF pages
* 🗑️ Delete PDF pages
* 🗜️ Compress PDF
* 🖼️ Extract images from PDF

---

## 🔐 Privacy First

Offline PDF Utility is designed with privacy in mind.

* ✅ Works completely offline
* ✅ No cloud upload
* ✅ No external server required
* ✅ Documents remain on your computer
* ✅ No account required
* ✅ No internet required for normal operation

Your files are processed locally instead of being sent to an online PDF service.

---

## 🖥️ Platform

Currently designed for:

**Kali Linux / Linux**

Windows and Android support may be considered for future versions.

---

## 🎨 Interface

The application includes a modern dashboard-style interface with:

* Dark theme
* Organized tool sections
* Conversion and PDF tool categories
* Clear tool descriptions
* Native Linux file selection
* Multi-file selection
* Simple status information
* Creator branding

---

## 🧰 Technologies

The project is built primarily with **Python** and uses several open-source libraries and utilities:

* Python
* Tkinter
* GTK
* Pillow
* PyMuPDF
* pypdf
* python-docx
* openpyxl
* python-pptx
* ReportLab
* LibreOffice
* Ghostscript

---

## 📁 Project Structure

```text
OfflinePDFUtility/
├── core/
│   ├── file_utils.py
│   ├── office_utils.py
│   ├── pdf_utils.py
│   └── image_utils.py
│
├── converters/
│   ├── image_to_pdf.py
│   ├── pdf_to_image.py
│   ├── office_to_pdf.py
│   ├── pdf_to_word.py
│   ├── pdf_to_ppt.py
│   └── pdf_to_excel.py
│
├── gui/
│   ├── file_dialog.py
│   ├── dialogs.py
│   └── main_window.py
│
├── pdf_tools/
│   ├── delete_pages.py
│   ├── split.py
│   ├── reorder_pages.py
│   ├── text_to_pdf.py
│   ├── merge.py
│   ├── compress.py
│   ├── extract_images.py
│   └── rotate.py
│
├── output/
├── main.py
├── requirements.txt
└── README.md
```

---

## 🚀 Installation

### 1. Clone the repository

```bash
git clone https://github.com/AakashSingh988/OfflinePDFUtility.git
cd OfflinePDFUtility
```

### 2. Install system dependencies

```bash
sudo apt update
sudo apt install -y python3-tk python3-gi gir1.2-gtk-3.0 libreoffice ghostscript
```

### 3. Install Python dependencies

```bash
python3 -m pip install --break-system-packages \
pillow \
pypdf \
pymupdf \
python-docx \
openpyxl \
reportlab \
python-pptx
```

---

## ▶️ Run

From the project directory:

```bash
python3 main.py
```

The application will open the Offline PDF Utility dashboard.

---

## 🧪 Check for Python Syntax Errors

Before running the application, you can check the main file with:

```bash
python3 -m py_compile main.py
```

If there is no output, the file passed the Python syntax check.

---

## 📌 Current Status

**Version:** `1.0`

**Status:** Active Development

The project is continuously being improved with new PDF tools, interface improvements, better error handling, and additional document-processing capabilities.

---

## 🗺️ Future Plans

Possible future improvements include:

* Drag-and-drop file support
* PDF preview
* Progress indicators
* Custom output directory
* Improved compression controls
* Better PDF page reordering
* Keyboard shortcuts
* Theme customization
* More document formats
* Windows support
* Android version

---

## 🤝 Contributing

Contributions, suggestions, bug reports, and improvements are welcome.

If you find a problem, please open an **Issue** and describe:

1. What happened
2. What you expected
3. Your operating system
4. The tool you were using
5. Any error message shown

Pull requests are also welcome.

---



## 👨‍💻 Creator

**Aakash998**

Built with **Python**, open-source technologies, and a focus on offline document processing.

---

## ⭐ Support the Project

If you find **Offline PDF Utility** useful, consider giving the repository a ⭐ on GitHub.

**Offline. Private. Simple. Powerful.**
