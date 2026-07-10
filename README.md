# 📊 AI Project Health Reporting Agent

## Overview

The AI Project Health Reporting Agent is a Python-based application that analyzes project data from an Excel workbook and automatically generates a project health report.

The application extracts important project metrics, calculates an overall health score using a custom scoring engine, generates an executive summary using Groq Llama 3.3, and creates a PowerPoint report.

---

## Features

- Read project data from an Excel workbook
- Extract important project metrics
- Calculate project health score
- Classify project status using RAG (Green / Amber / Red)
- Generate an AI-based executive summary using Groq
- Automatically create a PowerPoint report

---

## Project Workflow

```
Excel Workbook
      │
      ▼
Excel Parser
      │
      ▼
Metrics Extraction
      │
      ▼
Health Scoring Engine
      │
      ▼
RAG Status
      │
      ▼
Groq LLM
      │
      ▼
Executive Summary
      │
      ▼
PowerPoint Report
```

---

## Project Structure

```
app/
│
├── models/
│
├── parser/
│
├── presentation/
│
├── prompts/
│
├── services/
│
├── utils/
│
└── main.py

data/

outputs/
```

---

## Health Scoring

The overall project score is calculated using four metrics.

| Metric | Weight |
|---------|--------|
| Schedule Score | 35% |
| Completion Score | 25% |
| Milestone Score | 20% |
| Blocker Score | 20% |

Based on the final score:

| Score | Status |
|---------|--------|
| ≥ 75 | 🟢 Green |
| 50 – 74 | 🟡 Amber |
| < 50 | 🔴 Red |

---

## Technologies Used

- Python
- Pandas
- OpenPyXL
- python-pptx
- Groq API
- Llama 3.3 70B
- python-dotenv

---

## Installation

Clone the repository

```bash
git clone <repository-url>
```

Create a virtual environment

```bash
python -m venv .venv
```

Activate the environment

### Windows

```bash
.venv\Scripts\activate
```

Install dependencies

```bash
pip install -r requirements.txt
```

Create a `.env` file

```env
GROQ_API_KEY=your_api_key
```

---

## Run the Project

```bash
python -m app.main
```

---

## Output

The application generates:

- Project Health Score
- RAG Status
- Executive Summary
- PowerPoint Report

The PowerPoint report is saved in:

```
outputs/
Monthly_Project_Report.pptx
```

---

## Future Improvements

- Improve the health scoring logic
- Support multiple Excel files
- Add more visualizations to the PowerPoint report
- Build a simple web interface

---

## Author

**Naman Upreti**

Computer Science & Engineering (Data Science)