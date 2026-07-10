# 📊 AI Project Health Reporting Agent

## Overview

The AI Project Health Reporting Agent is a Python-based application that analyzes project planning data from an Excel workbook, calculates project health using a custom scoring engine, generates an AI-powered executive summary using Groq Llama 3.3, and automatically creates a PowerPoint report.

The goal of the project is to automate project health assessment and generate an executive-ready report from project data.

---

## Features

- Read project data from an Excel workbook
- Extract project metrics
- Calculate project health score
- Classify project status using RAG (Green / Amber / Red)
- Generate an AI-powered executive summary using Groq
- Automatically generate a PowerPoint report

---

## System Architecture

```text
              Excel Workbook (.xlsx)
                      │
                      ▼
                Excel Parser
                      │
                      ▼
           Metrics Extraction Engine
                      │
                      ▼
         Project Health Scoring Engine
                      │
                      ▼
             Health Assessment (RAG)
                      │
                      ▼
          Groq LLM (Llama 3.3 70B)
                      │
                      ▼
        Executive Summary Generator
                      │
                      ▼
        PowerPoint Report Generator
                      │
                      ▼
      Executive Project Health Report
```

---

## Project Structure

```text
AI-Project-Health-Reporting-Agent/

│
├── app/
│   ├── models/
│   ├── parser/
│   ├── presentation/
│   ├── prompts/
│   ├── services/
│   ├── utils/
│   └── main.py
│
├── data/
│
├── outputs/
│
├── requirements.txt
├── README.md
├── LICENSE
└── .gitignore
```

---

## Health Scoring

The overall project health score is calculated using four weighted metrics.

| Metric | Weight |
|---------|--------|
| Schedule Score | 35% |
| Completion Score | 25% |
| Milestone Score | 20% |
| Blocker Score | 20% |

Overall Score:

```
Overall Score =
0.35 × Schedule Score +
0.25 × Completion Score +
0.20 × Milestone Score +
0.20 × Blocker Score
```

---

## RAG Classification

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
- Llama 3.3 70B Versatile
- python-dotenv

---

## Installation

Clone the repository

```bash
git clone <repository-url>
```

Navigate to the project directory

```bash
cd AI-Project-Health-Reporting-Agent
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

- Overall Project Health Score
- RAG Status
- KPI Scores
- AI-generated Executive Summary
- Executive PowerPoint Report

Generated PowerPoint:

```text
outputs/
└── Monthly_Project_Report.pptx
```

---

## Future Improvements

- Improve the health scoring model
- Support multiple project reports
- Add more PowerPoint visualizations
- Build a simple web interface
- Export reports as PDF

---

## Author

**Naman Upreti**

Computer Science & Engineering (Data Science)

Interests:
- Artificial Intelligence
- Backend Development
- Generative AI
- Python

---

## License

This project is released under the MIT License.
