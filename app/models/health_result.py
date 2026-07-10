from dataclasses import dataclass


@dataclass
class HealthResult:
    """
    Represents the final health assessment of a project.
    """

    overall_score: float

    rag_status: str

    schedule_score: float

    milestone_score: float

    blocker_score: float

    completion_score: float

    confidence_score: float

    reasons: list[str]

    recommendations: list[str]