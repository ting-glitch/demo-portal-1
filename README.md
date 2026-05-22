# Enterprise Logistics & RMA Management System
### *Customized Automation Workflow Demo*

[![GitHub Pages](https://img.shields.io/badge/Live-Demo-blue?style=for-the-badge&logo=github)](https://ting-glitch.github.io/demo-portal-1/)
[![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.100+-05998b?style=for-the-badge&logo=fastapi&logoColor=white)](https://fastapi.tiangolo.com/)

A high-performance, enterprise-grade management dashboard designed to automate complex logistics workflows and **Return Merchandise Authorization (RMA)** processes. This platform bridges the gap between unstructured email communications and structured data management by integrating **FastAPI**, **Microsoft Outlook COM Interfaces**, and **OpenPyXL**.

> **Live Portfolio Demo:** [Experience the Interactive Interface](https://ting-glitch.github.io/demo-portal-1/) (Simulated environment optimized for static hosting).

---

## 🌟 Solution Highlights

### 1. Intelligent Outlook Integration
* Leverages a robust COM API layer (`pywin32`) to interface directly with local Microsoft Outlook installations.
* Features an asynchronous scanning engine that parses incoming communications to extract critical metadata, including batch identifiers, quantities, and tracking references through advanced regex patterns.

### 2. Multi-Stage Process Automation
* **Automated Data Ingestion**: Systematically categorizes email attachments into localized directories structured by Calendar Weeks (CW) and custom batch logic.
* **Logistics Synchronization**: Orchestrates the transition of records from initial receipt to shipping confirmation, ensuring real-time status accuracy.
* **Verification & Closure**: Streamlines the extraction of authorization codes from vendor receipts to provide a complete audit trail for each transaction.

### 3. Professional Glassmorphic UI/UX
* **Precision Engineering**: A fully responsive interface optimized for enterprise viewports, from 4K desktop monitors to mobile devices.
* **Dynamic Interaction**: Supports double-click inline editing for rapid data overrides and features real-time status visualizations via a progressive stepper component.
* **Demonstration Integrity**: Built-in "Simulation Mode" allows for a comprehensive showcase of automated features without requiring live mail server dependencies.

---

## 🏗️ System Architecture

```text
.
├── backend/                # High-concurrency FastAPI Service
│   ├── app.py              # Core Business Logic & API Endpoints
│   ├── .env.example        # Configuration Template
│   └── requirements.txt    # System Dependencies
├── frontend/               # UI Layer
│   └── templates/          # Responsive Glassmorphic Templates
├── docs/                   # GitHub Pages Deployment (Static Demo)
└── run.bat                 # Enterprise Startup Script
```

---

## 🛠️ Technical Implementation

*   **Backend Architecture**: FastAPI, Pydantic, Uvicorn, Win32COM (MAPI Integration).
*   **Data Reporting**: OpenPyXL-driven Relational Excel Generation (Styled for professional delivery).
*   **Frontend Technologies**: Semantic HTML5, Vanilla CSS3 (Custom Variables/Flexbox), ES6+ JavaScript.

---

## 🚀 Deployment Guide

### 📋 Prerequisites
* **Windows Environment** (Required for native Outlook COM integration).
* **Python 3.8+** Runtime.

### ⚡ Local Setup
1. **Initialize the Environment**:
   ```bash
   git clone https://github.com/ting-glitch/demo-portal-1.git
   cd demo-portal-1
   ```
2. **Launch the Platform**:
   ```bash
   run.bat
   ```
3. **Access the Dashboard**: Navigate to **`http://127.0.0.1:8000`** in your preferred browser.

---

## 🛡️ Data Privacy & Compliance
*   **Anonymized Datasets**: All demonstration data, including domains (`@client.example`) and tracking identifiers, are strictly fictional and de-identified to protect proprietary information.
*   **Environment Security**: Pre-configured `.gitignore` protocols ensure that local environment variables and internal logs remain excluded from version control.

---
*Developed by [ting-glitch](https://github.com/ting-glitch)*
