from __future__ import annotations

import base64
import re
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pandas as pd

from advanced_engine import tailoring_eligibility


SUITABLE = "Suitable for this Job"
NOT_SUITABLE = "Not Suitable for this Job"


def add_suitability(results: pd.DataFrame, job_description: str) -> pd.DataFrame:
    """Return a copy with the same suitability decision used by candidate cards."""
    enriched = results.copy()
    enriched["Suitability"] = [
        SUITABLE if tailoring_eligibility(row, job_description)[0] else NOT_SUITABLE
        for _, row in enriched.iterrows()
    ]
    return enriched


def filter_recruiter_results(
    results: pd.DataFrame,
    suitability: str = "All candidates",
    experience_rule: str = "Any experience",
    experience_years: float = 0,
    minimum_match: float = 0,
    minimum_skill_coverage: float = 0,
) -> pd.DataFrame:
    """Apply recruiter-selected filters without modifying ranks or scores."""
    filtered = results.copy()
    if suitability != "All candidates":
        filtered = filtered[filtered["Suitability"] == suitability]

    experience = pd.to_numeric(filtered["Experience"], errors="coerce")
    if experience_rule == "At least":
        filtered = filtered[experience >= experience_years]
    elif experience_rule == "More than":
        filtered = filtered[experience > experience_years]
    elif experience_rule == "At most":
        filtered = filtered[experience <= experience_years]

    match_score = pd.to_numeric(filtered["Match %"], errors="coerce").fillna(0)
    skill_score = pd.to_numeric(filtered["Skill coverage %"], errors="coerce").fillna(0)
    filtered = filtered[(match_score >= minimum_match) & (skill_score >= minimum_skill_coverage)]
    return filtered.reset_index(drop=True)


def recruiter_display_table(results: pd.DataFrame) -> pd.DataFrame:
    """Select and label columns shown and downloaded from recruiter screening."""
    columns = [
        "Rank", "Candidate", "Suitability", "Match %", "Semantic %",
        "Skill coverage %", "Experience",
    ]
    return results[columns].rename(columns={
        "Match %": "Overall Match %",
        "Semantic %": "Job Description Match %",
    })


def original_resume_payloads(
    results: pd.DataFrame,
    original_files: dict,
) -> list[dict]:
    """Return safely decoded originals for only the currently filtered candidates."""
    payloads = []
    for candidate in results.get("Candidate", []):
        original = original_files.get(candidate)
        if not original or not original.get("content_b64"):
            continue
        try:
            content = base64.b64decode(original["content_b64"], validate=True)
        except (ValueError, TypeError):
            continue
        payloads.append({
            "candidate": str(candidate),
            "name": Path(str(original.get("name") or f"{candidate}_resume")).name,
            "mime": str(original.get("mime") or "application/octet-stream"),
            "content": content,
        })
    return payloads


def build_filtered_resume_zip(display_table: pd.DataFrame, payloads: list[dict]) -> bytes:
    """Package every filtered original résumé together with the visible table."""
    output = BytesIO()
    used_names: set[str] = set()
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr("filtered_candidate_table.csv", display_table.to_csv(index=False).encode("utf-8-sig"))
        for payload in payloads:
            source_name = re.sub(r"[^A-Za-z0-9._ -]+", "_", payload["name"]).strip(" ._") or "resume"
            archive_name = f"resumes/{source_name}"
            stem, suffix = Path(source_name).stem, Path(source_name).suffix
            duplicate_number = 2
            while archive_name.casefold() in used_names:
                archive_name = f"resumes/{stem}_{duplicate_number}{suffix}"
                duplicate_number += 1
            used_names.add(archive_name.casefold())
            archive.writestr(archive_name, payload["content"])
    return output.getvalue()
