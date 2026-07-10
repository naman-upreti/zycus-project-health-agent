from pathlib import Path

from app.parser.excel_parser import ExcelParser
from app.services.metrics_extractor import MetricsExtractor
from app.services.health_scoring_engine import HealthScoringEngine
from app.services.reasoning_service import ReasoningService
from app.presentation.ppt_generator import PPTGenerator

def main():

    print("=" * 60)
    print("Project Health Reporting Agent")
    print("=" * 60)

    BASE_DIR = Path(__file__).resolve().parent.parent

    file_path = BASE_DIR / "data" / "S2P Project.xlsx"
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

    # Generate PowerPoint
    ppt = PPTGenerator(
        metrics=metrics,
        health=health_result,
        summary=summary
    )

    ppt.generate("outputs/Monthly_Project_Report.pptx")

    print("\nPowerPoint generated successfully!")
    print("Saved to: outputs/Monthly_Project_Report.pptx")

    # Console Report
    print("\n" + "=" * 60)
    print("PROJECT HEALTH REPORT")
    print("=" * 60)

    print(f"\nProject Name      : {metrics.project_name}")
    print(f"Project Manager   : {metrics.project_manager}")

    print(f"\nOverall Score     : {health_result.overall_score}/100")
    print(f"RAG Status        : {health_result.rag_status}")
    print(f"Confidence Score  : {health_result.confidence_score}%")

    print("\n" + "-" * 60)
    print("DECISION TRACE")
    print("-" * 60)

    print(f"Completion Score  : {health_result.completion_score}")
    print(f"Schedule Score    : {health_result.schedule_score}")
    print(f"Milestone Score   : {health_result.milestone_score}")
    print(f"Blocker Score     : {health_result.blocker_score}")

    print("\n" + "-" * 60)
    print("KEY FINDINGS")
    print("-" * 60)

    for reason in health_result.reasons:
        print(f"• {reason}")

    print("\n" + "-" * 60)
    print("RECOMMENDATIONS")
    print("-" * 60)

    for recommendation in health_result.recommendations:
        print(f"• {recommendation}")

    print("\n" + "-" * 60)
    print("EXECUTIVE SUMMARY")
    print("-" * 60)

    print(summary)


if __name__ == "__main__":
    main()