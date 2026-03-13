# Metadata Extractor

A **GUI-based Metadata Extraction Tool** built using **Python and CustomTkinter** that allows users to extract, view, copy, and modify metadata from multiple file formats.

The application provides a **drag-and-drop interface** for quick file analysis and supports common file types such as **images, PDFs, and Microsoft Office documents**. This tool can be useful for **digital forensics, cybersecurity research, document inspection, and metadata analysis**.

---

## Features

### Drag and Drop Interface

* Simple drag-and-drop file upload
* Automatically detects the file path
* Validates the file before processing

### Metadata Extraction

The tool can extract metadata from several file formats:

| File Type               | Metadata Extracted                                  |
| ----------------------- | --------------------------------------------------- |
| **Images (JPG / JPEG)** | EXIF data, camera information, GPS coordinates      |
| **PDF**                 | Author, creator, title, and other document metadata |
| **DOCX**                | Core document properties                            |
| **XLSX / XLS**          | Workbook metadata                                   |
| **PPTX / PPT**          | Presentation metadata                               |

---

## Image Metadata

For image files, the tool extracts EXIF information such as:

* Image width and height
* Camera make and model
* Date and time
* GPS coordinates
* GPS timestamp

GPS data is converted into **readable decimal coordinates**.

---

## PDF Metadata Editing

The application allows users to:

* View existing PDF metadata
* Edit metadata fields
* Save a new updated PDF file

---

## Copy Metadata

A **Copy button** is included to quickly copy the extracted metadata to the clipboard.

---

## Supported File Formats

* JPG / JPEG
* PDF
* DOCX
* XLSX / XLS
* PPTX / PPT

---

## Technologies Used

* Python
* CustomTkinter
* TkinterDnD2
* PyPDF2
* pikepdf
* python-docx
* openpyxl
* python-pptx
* piexif
* python-magic

---

## Installation

### Clone the Repository

```bash
git clone https://github.com/yourusername/metadata-extractor.git
cd metadata-extractor
```

### Install Dependencies

```bash
pip install -r requirements.txt
```

### Run the Application

```bash
python main.py
```

---

## Use Cases

* Digital Forensics
* Cybersecurity Analysis
* Metadata Investigation
* Document Inspection
* Privacy Auditing

---

## Project Structure

```
metadata-extractor/
│
├── main.py
├── requirements.txt
└── README.md
```

---

## Author

Rahul lakra
Cybersecurity Student | Python Developer
