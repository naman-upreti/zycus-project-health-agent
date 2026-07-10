import os

from dotenv import load_dotenv
from groq import Groq

from app.models.health_result import HealthResult

load_dotenv()


class ReasoningService:
    """
    Generates an executive summary using the Groq API.
    """

    def __init__(self):

        api_key = os.getenv("GROQ_API_KEY")

        if not api_key:
            raise ValueError(
                "GROQ_API_KEY not found in .env file"
            )

        self.client = Groq(api_key=api_key)

    def generate(self, result: HealthResult) -> str:

        prompt = f"""
You are a Senior PMO Executive at Zycus.

Generate a concise executive project health summary.

Project Health Details

Overall Score: {result.overall_score}
RAG Status: {result.rag_status}

Schedule Score: {result.schedule_score}
Completion Score: {result.completion_score}
Milestone Score: {result.milestone_score}
Blocker Score: {result.blocker_score}

Reasons:
{chr(10).join("- " + reason for reason in result.reasons)}

Recommendations:
{chr(10).join("- " + rec for rec in result.recommendations)}

Write the response in this format:

Executive Summary

Current Project Status

Key Risks

Business Impact

Recommendations

Keep it professional.
Maximum 250 words.
"""

        response = self.client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            temperature=0.3,
            max_tokens=500,
            messages=[
                {
                    "role": "system",
                    "content": "You are a senior Project Management Office executive."
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ]
        )

        return response.choices[0].message.content