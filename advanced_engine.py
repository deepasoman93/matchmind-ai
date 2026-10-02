from __future__ import annotations

import re
from typing import Callable

import numpy as np
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


MODEL_NAME = "sentence-transformers/all-mpnet-base-v2"
SKILL_ALIASES = {
    "power bi": ["power bi", "powerbi"], "business intelligence": ["business intelligence", "bi reporting"],
    "python": ["python"], "sql": ["sql", "t-sql", "mysql", "postgresql"], "pandas": ["pandas"],
    "numpy": ["numpy"], "excel": ["excel", "microsoft excel"], "tableau": ["tableau"],
    "data validation": ["data validation", "data quality", "quality checks"],
    "etl": ["etl", "data pipeline", "data pipelines"], "machine learning": ["machine learning", "ml"],
    "deep learning": ["deep learning"], "streamlit": ["streamlit"], "aws": ["aws", "amazon web services"],
    "azure": ["azure"], "gcp": ["gcp", "google cloud"], "spark": ["spark", "pyspark"],
    "selenium": ["selenium"], "cypress": ["cypress"], "javascript": ["javascript", "js"],
    "project management": ["project management", "project manager"],
    "stakeholder management": ["stakeholder management", "stakeholder communication"],
    "powerpoint": ["powerpoint", "power point", "ppt", "slide deck", "slide templates"],
    "spss": ["spss"],
    "market research": ["market research", "survey data", "quantitative research"],
    "reporting": ["reporting", "business reports", "analytical reports", "deliverables"],
    "quality control": ["quality control", "quality checks", "qc", "validation", "reconciliation"],
    "data analysis": ["data analysis", "data analyst", "analytics"],
    "communication": ["communication", "communicate", "collaborate", "stakeholder"],
    "clickup": ["clickup"], "sharepoint": ["sharepoint"], "onedrive": ["onedrive"],
    "vba": ["vba"], "power automate": ["power automate"], "power apps": ["power apps"],
}
EDUCATION_LEVELS = {
    "phd": 4, "doctorate": 4, "master": 3, "mba": 3, "m.tech": 3, "mtech": 3,
    "bachelor": 2, "b.tech": 2, "btech": 2, "degree": 2, "diploma": 1,
}


def clean_text(text: str) -> str:
    text = str(text or "").lower().replace("\n", " ").replace("–", "-").replace("—", "-")
    text = re.sub(r"[^a-z0-9+#./ -]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def chunk_text(text: str, size: int = 220, overlap: int = 40) -> list[str]:
    words = clean_text(text).split()
    if not words:
        return []
    step = max(1, size - overlap)
    return [" ".join(words[i:i + size]) for i in range(0, len(words), step)]


def extract_skills(text: str) -> set[str]:
    normalized = clean_text(text)
    return {skill for skill, aliases in SKILL_ALIASES.items() if any(alias in normalized for alias in aliases)}


def extract_required_skills(text: str) -> tuple[set[str], set[str]]:
    """Separate core requirements from explicitly preferred/advantage skills."""
    parts = re.split(
        r"preferred\s*/?\s*(?:added advantage)?\s*skills|preferred\s+(?:or\s+)?added\s+advantage|added\s+advantage",
        text,
        maxsplit=1,
        flags=re.IGNORECASE,
    )
    required = extract_skills(parts[0])
    preferred = extract_skills(parts[1]) if len(parts) > 1 else set()
    return required, preferred


def extract_required_years(text: str) -> int | None:
    minimum, _ = extract_required_experience_range(text)
    return minimum


def extract_required_experience_range(text: str) -> tuple[int | None, int | None]:
    normalized = clean_text(text)
    ranged = re.findall(
        r"(\d{1,2})\s*(?:-|–|to)\s*(\d{1,2})\s*(?:years?|yrs?)",
        normalized,
    )
    valid_ranges = [
        (int(low), int(high))
        for low, high in ranged
        if 0 <= int(low) <= int(high) <= 50
    ]
    if valid_ranges:
        return min(valid_ranges, key=lambda item: item[0])

    values = re.findall(r"(\d{1,2})\+?\s*(?:years?|yrs?)", clean_text(text))
    valid = [int(value) for value in values if int(value) <= 50]
    if not valid:
        return None, None
    minimum = min(valid)
    # A single value such as "1 year experience" states the minimum required
    # experience. Only an explicit range such as "1-2 years" supplies a maximum.
    return minimum, None


def extract_candidate_years(text: str) -> int | None:
    values = re.findall(r"(\d{1,2})\+?\s*(?:years?|yrs?)", clean_text(text))
    valid = [int(value) for value in values if int(value) <= 50]
    return max(valid) if valid else None


def education_level(text: str) -> tuple[int, str]:
    normalized = clean_text(text)
    found = [(level, term) for term, level in EDUCATION_LEVELS.items() if term in normalized]
    return max(found, default=(0, "not detected"))


def semantic_coverage(jd_embeddings: np.ndarray, resume_embeddings: np.ndarray) -> float:
    similarities = cosine_similarity(jd_embeddings, resume_embeddings)
    jd_coverage = similarities.max(axis=1).mean()
    resume_relevance = similarities.max(axis=0).mean()
    return float(np.clip(0.75 * jd_coverage + 0.25 * resume_relevance, 0, 1))


def tfidf_scores(job_description: str, resume_texts: list[str]) -> np.ndarray:
    """Score every resume against the JD independently of the candidate batch."""
    job_text = clean_text(job_description)
    scores = []
    for resume_text in resume_texts:
        documents = [job_text, clean_text(resume_text)]
        matrix = TfidfVectorizer(
            stop_words="english",
            ngram_range=(1, 2),
            sublinear_tf=True,
        ).fit_transform(documents)
        scores.append(float(cosine_similarity(matrix[0:1], matrix[1:2])[0, 0]))
    return np.asarray(scores, dtype=float)


def rank_resumes_advanced(
    job_description: str,
    resumes: list[tuple[str, str]],
    encode: Callable[[list[str]], np.ndarray],
) -> pd.DataFrame:
    if not clean_text(job_description):
        raise ValueError("Job description is empty.")
    resumes = [(name, text) for name, text in resumes if clean_text(text)]
    if not resumes:
        raise ValueError("No readable résumé text was found.")

    jd_chunks = chunk_text(job_description)
    all_chunks = jd_chunks[:]
    boundaries = []
    for _, text in resumes:
        chunks = chunk_text(text)
        start = len(all_chunks)
        all_chunks.extend(chunks)
        boundaries.append((start, len(all_chunks)))
    embeddings = np.asarray(encode(all_chunks))
    jd_embeddings = embeddings[:len(jd_chunks)]
    exact_scores = tfidf_scores(job_description, [text for _, text in resumes])
    jd_skills, preferred_skills = extract_required_skills(job_description)
    required_years = extract_required_years(job_description)
    required_education, required_education_name = education_level(job_description)

    rows = []
    for index, (candidate, resume_text) in enumerate(resumes):
        start, end = boundaries[index]
        semantic = semantic_coverage(jd_embeddings, embeddings[start:end])
        resume_skills = extract_skills(resume_text)
        matched = sorted(jd_skills & resume_skills)
        missing = sorted(jd_skills - resume_skills)
        preferred_matched = sorted(preferred_skills & resume_skills)
        skill_score = len(matched) / len(jd_skills) if jd_skills else float(exact_scores[index])

        candidate_years = extract_candidate_years(resume_text)
        experience_score = 1.0 if required_years is None else (
            min((candidate_years or 0) / required_years, 1.0) if required_years else 1.0
        )
        candidate_education, candidate_education_name = education_level(resume_text)
        education_score = 1.0 if required_education == 0 else min(candidate_education / required_education, 1.0)

        score = 100 * (
            0.55 * semantic + 0.20 * skill_score + 0.10 * float(exact_scores[index])
            + 0.10 * experience_score + 0.05 * education_score
        )
        strengths = []
        if matched:
            strengths.append("Matched required skills: " + ", ".join(matched))
        if preferred_matched:
            strengths.append("Matched preferred skills: " + ", ".join(preferred_matched))
        if required_years is not None and candidate_years is not None and candidate_years >= required_years:
            strengths.append(f"Experience evidence meets the {required_years}-year requirement")
        if required_education and candidate_education >= required_education:
            strengths.append(f"Education evidence ({candidate_education_name}) meets the stated requirement")
        gaps = []
        if missing:
            gaps.append("Missing or unclear required skills: " + ", ".join(missing))
        if required_years is not None and (candidate_years is None or candidate_years < required_years):
            gaps.append(f"Required {required_years} years; detected {candidate_years if candidate_years is not None else 'no clear duration'}")
        if required_education and candidate_education < required_education:
            gaps.append(f"Required {required_education_name}; detected {candidate_education_name}")

        rows.append({
            "Candidate": candidate, "Match %": round(float(np.clip(score, 0, 100)), 1),
            "Semantic %": round(semantic * 100, 1), "Skill coverage %": round(skill_score * 100, 1),
            "Experience": candidate_years if candidate_years is not None else "Not clearly detected",
            "Strengths": ". ".join(strengths) or "Relevant meaning detected, but no structured requirement was confirmed.",
            "Gaps": ". ".join(gaps) or "No major structured gap detected; verify manually.",
        })

    result = pd.DataFrame(rows).sort_values(["Match %", "Candidate"], ascending=[False, True]).reset_index(drop=True)
    result.insert(0, "Rank", range(1, len(result) + 1))
    return result


def tailoring_eligibility(
    result_row: pd.Series | dict,
    job_description: str,
    minimum_match: float = 50.0,
    minimum_skill_coverage: float = 50.0,
) -> tuple[bool, str]:
    """Require overall JD match, skill coverage, and experience-range alignment."""
    match_score = float(result_row.get("Match %", result_row["Semantic %"]))
    semantic_score = float(result_row["Semantic %"])
    skill_score = float(result_row["Skill coverage %"])
    if match_score < minimum_match or skill_score < minimum_skill_coverage:
        return False, (
            f"Requires at least {minimum_match:.0f}% overall match and "
            f"{minimum_skill_coverage:.0f}% skill coverage; detected "
            f"{match_score:.1f}% and {skill_score:.1f}%."
        )

    minimum_years, maximum_years = extract_required_experience_range(job_description)
    if minimum_years is None:
        return False, "The job description does not state a clear experience requirement."
    candidate_years = result_row["Experience"]
    if not isinstance(candidate_years, (int, float, np.integer, np.floating)):
        return False, "The resume does not state a clear number of years of experience."
    candidate_years = int(candidate_years)
    if candidate_years < minimum_years:
        return False, f"Requires at least {minimum_years} years; detected {candidate_years}."
    if maximum_years is not None and candidate_years > maximum_years:
        return False, (
            f"Requires {minimum_years}–{maximum_years} years; detected {candidate_years}."
        )
    requirement = f"{minimum_years}+" if maximum_years is None else (
        str(minimum_years) if minimum_years == maximum_years else f"{minimum_years}–{maximum_years}"
    )
    return True, (
        f"Eligible: overall match {match_score:.1f}%, job description match {semantic_score:.1f}%, "
        f"skills {skill_score:.1f}%, "
        f"experience {candidate_years} years within the {requirement}-year requirement."
    )
