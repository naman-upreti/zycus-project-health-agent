# 📊 AI Project Health Reporting Agent

> An AI-powered Project Health Reporting system that automatically analyzes Microsoft Project Excel exports, evaluates project health using a custom scoring engine, generates executive insights with an LLM, and produces a professional PowerPoint report.

---

## 🚀 Overview

Project managers often spend hours manually reviewing project plans, calculating KPIs, preparing executive summaries, and creating presentation decks.

The **AI Project Health Reporting Agent** automates this entire workflow.

Given an exported Microsoft Project Excel file, the system:

- Parses project data
- Extracts project KPIs
- Calculates project health using a custom scoring engine
- Classifies project status using a RAG (Red-Amber-Green) model
- Uses an LLM (Groq + Llama 3.3) to generate an executive summary
- Automatically creates a professional PowerPoint report for stakeholders

---

# ✨ Features

✅ Automated Excel Parsing

✅ Intelligent Project Metrics Extraction

✅ Custom Project Health Scoring Engine

✅ RAG (Red / Amber / Green) Classification

✅ AI-generated Executive Summary

✅ Professional PowerPoint Report Generation

✅ Modular Clean Architecture

---

# 🏗️ System Architecture

```text
                 Microsoft Project
                        │
                        ▼
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
        Executive Project Health Report (.pptx)
```

---

# 📂 Project Structure

```text
Project-Health-Reporting-Agent/

│
├── app/
│   ├── agents/
│   ├── models/
│   │     ├── health_result.py
│   │     └── project_metrics.py
│   │
│   ├── parser/
│   │     └── excel_parser.py
│   │
│   ├── presentation/
│   │     └── ppt_generator.py
│   │
│   ├── prompts/
│   │
│   ├── services/
│   │     ├── health_scoring_engine.py
│   │     ├── metrics_extractor.py
│   │     └── reasoning_service.py
│   │
│   ├── utils/
│   │
│   └── main.py
│
├── data/
│     └── S2P Project.xlsx
│
├── outputs/
│     └── Monthly_Project_Report.pptx
│
├── requirements.txt
├── README.md
└── .env
```

---

# ⚙️ Workflow

### Step 1

Read Microsoft Project Excel Export

↓

### Step 2

Extract

- Total Tasks
- Completed Tasks
- Delayed Tasks
- Critical Tasks
- Milestones
- Average Completion

↓

### Step 3

Calculate

- Completion Score
- Schedule Score
- Milestone Score
- Blocker Score

↓

### Step 4

Generate

Overall Project Health Score

↓

### Step 5

Determine

- 🟢 Green
- 🟡 Amber
- 🔴 Red

↓

### Step 6

Generate Executive Summary using Groq Llama 3.3

↓

### Step 7

Generate Executive PowerPoint Report

---

# 📈 Health Scoring Logic

The project health score is calculated using weighted KPIs:

| Metric | Weight |
|---------|--------|
| Schedule Score | 35% |
| Completion Score | 25% |
| Milestone Score | 20% |
| Blocker Score | 20% |

Final Health Score:

```
Overall Health =
0.35 × Schedule
+ 0.25 × Completion
+ 0.20 × Milestones
+ 0.20 × Blockers
```

---

# 🟢 RAG Classification

| Score | Status |
|---------|--------|
| ≥ 75 | 🟢 Green |
| 50–74 | 🟡 Amber |
| < 50 | 🔴 Red |

---

# 🤖 AI Integration

The project uses:

- **Groq API**
- **Llama 3.3 70B Versatile**

The LLM transforms numerical project metrics into a concise executive summary covering:

- Current Project Status
- Key Risks
- Business Impact
- Recommendations

---

# 📑 PowerPoint Report

The system automatically generates a multi-slide executive report including:

- Executive Cover
- KPI Dashboard
- Risk Assessment
- Recommendations
- AI Executive Summary
- Project Snapshot & Conclusion

---

# 🛠️ Tech Stack

### Programming

- Python 3

### Data Processing

- Pandas
- OpenPyXL

### AI

- Groq API
- Llama 3.3 70B

### Presentation

- python-pptx

### Environment

- python-dotenv

---

# 🚀 Installation

Clone the repository

```bash
git clone <repository-url>
```

Create a virtual environment

```bash
python -m venv .venv
```

Activate it

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
GROQ_API_KEY=your_api_key_here
```

---

# ▶️ Run

```bash
python -m app.main
```

---

# 📁 Output

The generated report will be saved in:

```text
outputs/
    Monthly_Project_Report.pptx
```

---

# 🔮 Future Improvements

- Interactive Dashboard
- Multi-project Portfolio Analysis
- Predictive Project Health Forecasting
- Risk Trend Analysis
- Email Report Automation
- PDF Report Generation
- Web Dashboard using FastAPI

---

# 👨‍💻 Author

**Naman**

Computer Science & Engineering (Data Science)

AI • Backend Engineering • Generative AI • Python

---

# 📜 License

This project was developed for educational and portfolio purposes.