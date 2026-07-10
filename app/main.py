from pathlib import Path

from app.parser.excel_parser import ExcelParser
from app.services.metrics_extractor import MetricsExtractor
from app.services.health_scoring_engine import HealthScoringEngine
from app.services.reasoning_service import ReasoningService

from app.presentation.monthly_ppt_generator import MonthlyPPTGenerator

def main():

    print("=" * 60)
    print("Project Health Reporting Agent")
    print("=" * 60)

    BASE_DIR = Path(__file__).resolve().parent.parent

    # Read Excel
    reports = []

    data_folder = BASE_DIR / "data"

    for file_path in data_folder.glob("*.xlsx"):

        print("\n" + "=" * 60)
        print(f"Processing: {file_path.name}")
        print("=" * 60)

        # Read Excel
        parser = ExcelParser(file_path)
        sheet_name = parser.get_sheet_names()[0]
        df = parser.read_sheet(sheet_name)

        # Extract Metrics
        extractor = MetricsExtractor(
            dataframe=df,
            file_path=file_path
        )

        metrics = extractor.extract()

        # Calculate Health
        engine = HealthScoringEngine(metrics)
        health_result = engine.calculate()

        # Generate Executive Summary
        reasoner = ReasoningService()
        summary = reasoner.generate(health_result)

        reports.append({
            "metrics": metrics,
            "health": health_result,
            "summary": summary,
        })
    # Console Report
    print("\n")
    print("=" * 60)
    print("Processed Projects")
    print("=" * 60)

    print("\n")
    print("=" * 70)
    print("PROJECT PORTFOLIO SUMMARY")
    print("=" * 70)

    for report in reports:

        metrics = report["metrics"]
        health = report["health"]

        print(f"\nProject Name     : {metrics.project_name}")
        print(f"Project Manager  : {metrics.project_manager}")
        print(f"Overall Score    : {health.overall_score}")
        print(f"RAG Status       : {health.rag_status}")

    ppt = MonthlyPPTGenerator(reports)

    output_path = BASE_DIR / "outputs" / "Monthly_Project_Report.pptx"

    ppt.generate(output_path)

    print("\nPowerPoint generated successfully!")
    print(f"Saved to: {output_path}")

if __name__ == "__main__":
    main()