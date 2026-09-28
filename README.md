# Trackon-API: Courier Automation & Weight Extraction Engine

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?style=flat&logo=python)](https://python.org)
[![Selenium](https://img.shields.io/badge/Selenium-Automation-43B02A?style=flat&logo=selenium)](https://www.selenium.dev/)
[![Docker](https://img.shields.io/badge/Docker-Ready-2496ED?style=flat&logo=docker)](https://www.docker.com/)

> **An automated scraping, tracking, and reconciliation system for the Trackon Courier enterprise portal with CAPTCHA handling, Excel synchronization, and headless Docker execution.**

---

## 📌 Problem Statement
E-commerce businesses and logistics managers frequently suffer from shipping weight discrepancies charged by courier aggregators. Manually querying hundreds of Airway Bill (AWB) numbers on the Trackon portal to extract billed vs. actual weights is tedious, error-prone, and rate-limited by CAPTCHAs.

`Trackon-API` automates this entire lifecycle:
1. Ingests bulk booking shipments from Excel (`bookings.xlsx`).
2. Programmatically handles Trackon portal authentication and CAPTCHA solving.
3. Scrapes actual shipment weights, delivery milestones, and POD (Proof of Delivery) records.
4. Reconciles and writes updated verification metrics back into Excel.

---

## 🚀 Key Features

* **Headless Browser Automation**: Built using Selenium WebDriver with custom anti-detection parameters and resilient retry policies.
* **CAPTCHA Solver Module (`captcha_solver.py`)**: Automatic image pre-processing, OCR extraction, and manual fallback handling for Trackon's security challenges.
* **Bi-Directional Excel Synchronization**:
  - Automatically reads input consignments from `bookings.xlsx`.
  - Enriches sheets with live status, actual weight, volumetric weight, and billing charges.
* **Docker Containerized**: Fully containerized with a self-contained headless Chromium and ChromeDriver environment for cloud worker deployments.
* **REST API Server (`main.py`)**: Exposes API endpoints to trigger automated bulk extraction runs or query individual AWB tracking numbers on demand.

---

## 📁 Repository Structure
```text
Trackon-API/
├── Dockerfile                  # Container definition with Python & headless Chrome
├── requirements.txt            # Python dependencies (Selenium, pandas, openpyxl, etc.)
├── main.py                     # API server interface and job runner
├── trackon_automation.py       # Core Selenium scraping and session orchestrator
├── captcha_solver.py           # CAPTCHA image processing and OCR solver
├── bookings.xlsx               # Sample workbook for batch consignment input/output
├── .env.template               # Template for portal credentials and scraper configs
└── .gitignore
```

---

## 🛠️ Tech Stack
- **Python 3.10+**
- **Selenium WebDriver** (Browser automation & DOM traversal)
- **Pandas & OpenPyXL** (High-throughput Excel manipulation)
- **OpenCV / Pillow & Tesseract** (CAPTCHA image cleaning and OCR)
- **Docker** (Production deployment environment)

---

## 🚀 Quick Start

### 1. Local Setup
```bash
git clone https://github.com/01mrabhi/Trackon-API.git
cd Trackon-API
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Configure Credentials
Copy `.env.template` to `.env`:
```bash
cp .env.template .env
```
Provide your Trackon corporate portal credentials:
```env
TRACKON_USERNAME=your_username
TRACKON_PASSWORD=your_password
HEADLESS=true
```

### 3. Run Automation
To execute the automated Excel batch extraction:
```bash
python trackon_automation.py
```
Or to run the API service:
```bash
python main.py
```

### 4. Run via Docker
```bash
docker build -t trackon-api .
docker run -d --env-file .env -v $(pwd)/data:/app/data trackon-api
```

---

## 📄 License
Maintained and authored by [01mrabhi](https://github.com/01mrabhi).
