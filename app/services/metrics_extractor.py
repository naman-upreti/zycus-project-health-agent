from pathlib import Path

import pandas as pd

from app.models.project_metrics import ProjectMetrics

class MetricsExtractor:
    """
    Extracts useful project metrics from the project plan.
    """

    def __init__(self, dataframe: pd.DataFrame, file_path: str):

        self.df = dataframe.copy()
        self.file_name = Path(file_path).stem

    # -------------------------
    # Project Information
    # -------------------------

    def _get_project_name(self):

        if "Project Name" in self.df.columns:

            values = self.df["Project Name"].dropna()

            if not values.empty:
                return values.iloc[0]

        # Fallback
        return self.file_name

    def _get_project_manager(self):

        if "Project Manager" in self.df.columns:

            values = self.df["Project Manager"].dropna()

            if not values.empty:
                return values.iloc[0]

        return "Unknown"

    # -------------------------
    # Task Counts
    # -------------------------

    def _total_tasks(self):

        return len(self.df)

    def _completed_tasks(self):

        return (self.df["Status"] == "Completed").sum()

    def _in_progress_tasks(self):

        return (self.df["Status"] == "In Progress").sum()

    def _not_started_tasks(self):

        return (self.df["Status"] == "Not Started").sum()

    def _on_hold_tasks(self):

        return (self.df["Status"] == "On Hold").sum()

    # -------------------------
    # Critical Tasks
    # -------------------------

    def _critical_tasks(self):

        if "Critical ?" not in self.df.columns:
            return 0

        return self.df["Critical ?"].fillna(0).astype(int).sum()

    # -------------------------
    # Delayed Tasks
    # -------------------------

    def _delayed_tasks(self):

        if "Variance" not in self.df.columns:
            return 0

        variance = pd.to_numeric(
            self.df["Variance"],
            errors="coerce"
        )

        return (variance > 0).sum()

    # -------------------------
    # Completion
    # -------------------------

    def _average_completion(self):

        completion = pd.to_numeric(
            self.df["% Complete"],
            errors="coerce"
        )

        return round(completion.mean(), 2)

    # -------------------------
    # Schedule Health
    # -------------------------

    def _schedule_health(self):

        if "Schedule Health" not in self.df.columns:
            return "Unknown"

        return (
            self.df["Schedule Health"]
            .dropna()
            .mode()
            .iloc[0]
        )

    # -------------------------
    # Milestones
    # -------------------------

    def _total_milestones(self):

        if "Phase/Milestone" not in self.df.columns:
            return 0

        return self.df["Phase/Milestone"].notna().sum()

    def _completed_milestones(self):

        if "Phase/Milestone" not in self.df.columns:
            return 0

        milestone_df = self.df[
            self.df["Phase/Milestone"].notna()
        ]

        return (
            milestone_df["Status"] == "Completed"
        ).sum()

    # -------------------------
    # High Priority
    # -------------------------

    def _high_priority_tasks(self):

        if "Priority" not in self.df.columns:
            return 0

        return (
            self.df["Priority"]
            .fillna("")
            .str.lower()
            .eq("high")
            .sum()
        )

    # -------------------------
    # Final Output
    # -------------------------

    def extract(self):

        return ProjectMetrics(

            project_name=self._get_project_name(),

            project_manager=self._get_project_manager(),

            total_tasks=self._total_tasks(),

            completed_tasks=self._completed_tasks(),

            in_progress_tasks=self._in_progress_tasks(),

            not_started_tasks=self._not_started_tasks(),

            on_hold_tasks=self._on_hold_tasks(),

            critical_tasks=self._critical_tasks(),

            delayed_tasks=self._delayed_tasks(),

            average_completion=self._average_completion(),

            schedule_health=self._schedule_health(),

            total_milestones=self._total_milestones(),

            completed_milestones=self._completed_milestones(),

            high_priority_tasks=self._high_priority_tasks()
        )