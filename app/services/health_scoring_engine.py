from app.models.project_metrics import ProjectMetrics
from app.models.health_result import HealthResult

class HealthScoringEngine:
    """
    Calculates the overall project health (RAG) based on extracted metrics
    """
    def __init__(self, metrics: ProjectMetrics):
        self.metrics = metrics 
        
    def _calculate_completion_score(self):
        completion = self.metrics.average_completion

        if completion <= 1:
            completion *= 100

        return round(completion, 2)
    
    def _calculate_milestone_score(self):

        if self.metrics.total_milestones == 0:
            return 100

        score = (
            self.metrics.completed_milestones
            / self.metrics.total_milestones
        ) * 100

        return round(score, 2)
    
    
    def _calculate_blocker_score(self):

        penalty = (
            self.metrics.critical_tasks * 3
            + self.metrics.on_hold_tasks * 5
        )

        score = max(0, 100 - penalty)

        return score
    
    
    def _calculate_schedule_score(self):

        if self.metrics.total_tasks == 0:
            return 100

        delayed_ratio = (
            self.metrics.delayed_tasks
            / self.metrics.total_tasks
        )

        score = 100 - (delayed_ratio * 100)

        return round(max(score, 0), 2)
        
    def _calculate_overall_score(self) -> float:

        schedule = self._calculate_schedule_score()
        completion = self._calculate_completion_score()
        milestone = self._calculate_milestone_score()
        blocker = self._calculate_blocker_score()

        overall = (
            schedule * 0.35 +
            completion * 0.25 +
            milestone * 0.20 +
            blocker * 0.20
        )

        return round(overall, 2)
    
    
        
    def calculate(self) -> HealthResult:

        schedule_score = self._calculate_schedule_score()
        completion_score = self._calculate_completion_score()
        milestone_score = self._calculate_milestone_score()
        blocker_score = self._calculate_blocker_score()

        overall_score = self._calculate_overall_score() 
        
        if overall_score >= 75:
            rag_status = "Green"    
        elif overall_score >= 50:  
            rag_status = "Amber"
        else:
            rag_status = "Red"        
            
            
        reasons = []
        
        
        if completion_score < 80:
            reasons.append(
                f"Project completion is only {completion_score:.1f}%."
            )

        if self.metrics.critical_tasks > 0:
            reasons.append(
                f"{self.metrics.critical_tasks} critical tasks require attention."
            )

        if self.metrics.on_hold_tasks > 0:
            reasons.append(
                f"{self.metrics.on_hold_tasks} tasks are currently on hold."
            )

        if milestone_score < 75:
            reasons.append(
                "Milestone completion is behind schedule."
            )

        if schedule_score < 80:
            reasons.append(
                "Schedule slippage has been detected."
            )

        if not reasons:
            reasons.append(
                "Project is progressing as planned."
            )

        recommendations = []

        if rag_status == "Green":

            recommendations.append(
                "Continue monitoring critical tasks to avoid schedule slippage."
            )

            recommendations.append(
                "Maintain the current execution plan and track milestone progress."
            )

        elif rag_status == "Amber":

            recommendations.append(
                "Review delayed tasks and resolve blockers as soon as possible."
            )

            recommendations.append(
                "Increase milestone review frequency."
            )

        else:

            recommendations.append(
                "Immediate management attention is required."
            )

            recommendations.append(
                "Re-plan the schedule and resolve critical blockers."
            )

        confidence_score = 100.0

        if self.metrics.total_tasks == 0:
            confidence_score -= 20

        if self.metrics.total_milestones == 0:
            confidence_score -= 10
        
        
        overall_score = float(overall_score)
        schedule_score = float(schedule_score)
        milestone_score = float(milestone_score)
        blocker_score = float(blocker_score)
        completion_score = float(completion_score)
        return HealthResult(

            overall_score=overall_score,

            rag_status=rag_status,

            schedule_score=schedule_score,

            milestone_score=milestone_score,

            blocker_score=blocker_score,

            completion_score=completion_score,

            confidence_score=confidence_score,

            reasons=reasons,

            recommendations=recommendations
        )