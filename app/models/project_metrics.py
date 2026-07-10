from dataclasses import dataclass


@dataclass
class ProjectMetrics:

    project_name: str

    project_manager: str

    total_tasks: int

    completed_tasks: int

    in_progress_tasks: int

    not_started_tasks: int

    on_hold_tasks: int

    critical_tasks: int

    delayed_tasks: int

    average_completion: float

    schedule_health: str

    total_milestones: int

    completed_milestones: int

    high_priority_tasks: int