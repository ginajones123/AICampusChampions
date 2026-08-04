import io
import json
import os
import re
from dataclasses import dataclass
from datetime import date
from typing import Optional

import docx
import PyPDF2
import requests
import streamlit as st

def as_dict(x):
    return x if isinstance(x, dict) else {}

def as_list(x):
    return x if isinstance(x, list) else []

st.set_page_config(
    page_title="U of A HR MQ, Relevant Experience & Preliminary Salary Range Helper",
    page_icon="🎓",
    layout="wide",
)

if "raw_structured_data" not in st.session_state:
    st.session_state["raw_structured_data"] = {}

if "structured_data" not in st.session_state:
    st.session_state["structured_data"] = {}

if "final_output" not in st.session_state:
    st.session_state["final_output"] = ""

if "last_json_text" not in st.session_state:
    st.session_state["last_json_text"] = ""

if "last_structured_parse" not in st.session_state:
    st.session_state["last_structured_parse"] = None

st.title("🎓 U of A HR MQ, Relevant Experience & Preliminary Salary Range Helper")
st.write("Upload one candidate or incumbent resume and one University job description, then click Run HR review.")

SYSTEM_PROMPT = r'''You are the U of A HR MQ, Relevant Experience & Preliminary Salary Helper.

You assist Human Resources with a decision-support workflow only. Compare the attached candidate resume to the attached University job description, apply the embedded equivalency rules, and produce a short structured HR review plus the required machine-readable JSON.

Use only the uploaded files and these instructions. Do not search the web or use outside sources.

Always produce the short version unless the user explicitly asks for details.

Visible output must begin directly with:
## HR MQ, Relevant Experience & Preliminary Salary Range Review

Print these exact sections in this exact order:
## 1. MINIMUM QUALIFICATIONS SNAPSHOT
## 2. EXPERIENCE CALCULATION
## 3. RELEVANT EXPERIENCE ALIGNMENT
## 4. FLAGS FOR HR REVIEWER
## 5. PRELIMINARY SALARY RANGE
## 6. PRELIMINARY DISPOSITION

Section 2 visible table must use exactly these columns:
Role | Dates | Duration | Relevance | School Overlap | Allowed Relevance

In Section 2:
• Relevance = raw duty alignment before school-time adjustment.
• Allowed Relevance = school-time-adjusted relevance actually used in totaling.
• School Overlap must be Yes, No, Partial, or Unclear.
• End with exactly: Relevant experience total: X.X years ~ Y+ years

Machine-readable JSON requirement:
Append valid JSON inside this exact wrapper:
<<<STRUCTURED_DATA_START>>>
{ ... }
<<<STRUCTURED_DATA_END>>>

Rules:
• Output valid JSON only inside the wrapper.
• Do not omit keys; use null only if truly unclear.
• The JSON values must match the visible output.
• Return the JSON object using these top-level keys when possible: minimum_qualifications, education_mq, experience_mq, relevant_experience_alignment, flags, preliminary_disposition.
• For education_mq, include a `degrees` array where each entry uses exactly these fields: degree, school, completion_date, completed, years.
• For education_mq, use `degrees` as the primary array name. If there is no degree data, return an empty array rather than omitting the key.
• For experience_mq, include a `roles` array and, when available, a `principal_responsibilities` array of strings containing only the job description's principal responsibilities.
• For relevant_experience_alignment, include an array of objects where each entry uses exactly these fields: jd_responsibility, resume_evidence, alignment.
• If the model cannot identify a responsibility or evidence, still return the object with the required keys and use empty strings rather than omitting the field.
• For non-OC target roles, overlapping full-time school work counts at 0%.
• For OC1-OC4 target roles, overlapping full-time school work counts at 25% if relevant.
• allowed_contribution_years must equal duration_years × allowed_relevance_pct.
• For the visible Section 1 Relevant Experience row, base the status on allowed counted experience only, not on degree equivalency.
• If allowed counted experience is below the stated experience requirement, do not mark the Section 1 Relevant Experience row as Exceeds.
• Equivalency may still permit overall MQ status of MET or EXCEEDS in later sections.
• Review the resume for actual education credentials, degree completion status, completion date, and school name. Do not infer degree information from context or keyword guessing.
• If the resume mentions apprenticeship, technical training, vocational training, or similar non-degree preparation that is relevant to the target role, capture that as `training_evidence` in the education payload so it can support the Section 1 evidence summary.
• Only count a degree if it is explicitly shown as completed on the resume and the resume lists the school name next to the degree.
• If more than one degree is present, only the highest completed degree should count.
• If a degree was not completed, it must not be counted even if the resume mentions the degree title.
• When returning education data, use a structured array of education entries with fields such as degree, school, completion_date, completed, and years. Use the highest completed degree entry for the summary fields.
• If no degree is indicated, use a fallback degree label of "Not stated" and a 0-year value rather than guessing from unrelated education or training language.
• For the job description, extract principal responsibilities strictly from the section titled "Principal Responsibilities" (or the closest equivalent section). Do not mix in requirements from other sections such as Education, Minimum Qualifications, Preferred Qualifications, Knowledge, Skills, or Special Instructions when building the Section 3 job-requirement rows.
• In Section 3, include a markdown table with one row per principal responsibility from the job description whenever possible.
• In Section 4, output a proper markdown table with columns: # | Flag | Detail.
'''

STRUCTURED_SYSTEM_PROMPT = SYSTEM_PROMPT + """
Machine-readable JSON requirement (JSON-only pass):
Append valid JSON inside this exact wrapper (no prose before or after):
<<<STRUCTURED_DATA_START>>>
{ ... }
<<<STRUCTURED_DATA_END>>>

Return JSON only inside the wrapper, no markdown, no commentary.
"""
DEBUG = False
DEFAULT_MODEL = "ai2s-external-claude-sonnet-4-6"
API_BASE = "https://llm-api.cyverse.ai/v1"
CHAT_COMPLETIONS_URL = f"{API_BASE}/chat/completions"

GRADE_SUFFIX_MAP = {2: "GR2", 3: "GR3-8", 4: "GR3-8", 5: "GR3-8", 6: "GR3-8", 7: "GR3-8", 8: "GR3-8", 9: "GR9-16", 10: "GR9-16", 11: "GR9-16", 12: "GR9-16", 13: "GR9-16", 14: "GR9-16", 15: "GR9-16", 16: "GR9-16", 17: "GR17-20", 18: "GR17-20", 19: "GR17-20", 20: "GR17-20"}
CO_RATIO_TABLE = {"GR2": {1: {"min": 0.88, "max": 0.927143, "quartile": "First"}, 2: {"min": 0.884286, "max": 0.944286, "quartile": "First"}, 3: {"min": 0.901429, "max": 0.961429, "quartile": "First"}, 4: {"min": 0.918571, "max": 0.978571, "quartile": "Second"}, 5: {"min": 0.935714, "max": 0.995714, "quartile": "Second"}, 6: {"min": 0.952857, "max": 1.012857, "quartile": "Second"}, 7: {"min": 0.97, "max": 1.03, "quartile": "Third"}, 8: {"min": 0.995385, "max": 1.035385, "quartile": "Third"}, 9: {"min": 1.010769, "max": 1.050769, "quartile": "Third"}, 10: {"min": 1.026154, "max": 1.066154, "quartile": "Third"}, 11: {"min": 1.041538, "max": 1.081538, "quartile": "Third"}, 12: {"min": 1.056923, "max": 1.096923, "quartile": "Third"}, 13: {"min": 1.072308, "max": 1.112308, "quartile": "Fourth"}, 14: {"min": 1.087692, "max": 1.127692, "quartile": "Fourth"}, 15: {"min": 1.103077, "max": 1.143077, "quartile": "Fourth"}, 16: {"min": 1.118462, "max": 1.158462, "quartile": "Fourth"}, 17: {"min": 1.133846, "max": 1.173846, "quartile": "Fourth"}, 18: {"min": 1.149231, "max": 1.189231, "quartile": "Fourth"}, 19: {"min": 1.164615, "max": 1.2, "quartile": "Fourth"}}, "GR3-8": {1: {"min": 0.80003, "max": 0.858597, "quartile": "First"}, 2: {"min": 0.827164, "max": 0.887164, "quartile": "First"}, 3: {"min": 0.855731, "max": 0.915731, "quartile": "First"}, 4: {"min": 0.884299, "max": 0.944299, "quartile": "Second"}, 5: {"min": 0.912866, "max": 0.972866, "quartile": "Second"}, 6: {"min": 0.941433, "max": 1.001433, "quartile": "Second"}, 7: {"min": 0.97, "max": 1.03, "quartile": "Third"}, 8: {"min": 0.995384, "max": 1.035384, "quartile": "Third"}, 9: {"min": 1.010767, "max": 1.050767, "quartile": "Third"}, 10: {"min": 1.026151, "max": 1.066151, "quartile": "Third"}, 11: {"min": 1.041535, "max": 1.081535, "quartile": "Third"}, 12: {"min": 1.056918, "max": 1.096918, "quartile": "Third"}, 13: {"min": 1.072302, "max": 1.112302, "quartile": "Fourth"}, 14: {"min": 1.087686, "max": 1.127686, "quartile": "Fourth"}, 15: {"min": 1.103069, "max": 1.143069, "quartile": "Fourth"}, 16: {"min": 1.118453, "max": 1.158453, "quartile": "Fourth"}, 17: {"min": 1.133837, "max": 1.173837, "quartile": "Fourth"}, 18: {"min": 1.149221, "max": 1.189221, "quartile": "Fourth"}, 19: {"min": 1.164604, "max": 1.199988, "quartile": "Fourth"}}, "GR9-16": {1: {"min": 0.769239, "max": 0.832205, "quartile": "First"}, 2: {"min": 0.805171, "max": 0.865171, "quartile": "First"}, 3: {"min": 0.838136, "max": 0.898136, "quartile": "First"}, 4: {"min": 0.871102, "max": 0.931102, "quartile": "Second"}, 5: {"min": 0.904068, "max": 0.964068, "quartile": "Second"}, 6: {"min": 0.937034, "max": 0.997034, "quartile": "Second"}, 7: {"min": 0.97, "max": 1.03, "quartile": "Third"}, 8: {"min": 0.997751, "max": 1.037751, "quartile": "Third"}, 9: {"min": 1.015503, "max": 1.055503, "quartile": "Third"}, 10: {"min": 1.033254, "max": 1.073254, "quartile": "Third"}, 11: {"min": 1.051005, "max": 1.091005, "quartile": "Third"}, 12: {"min": 1.068757, "max": 1.108757, "quartile": "Third"}, 13: {"min": 1.086508, "max": 1.126508, "quartile": "Fourth"}, 14: {"min": 1.104259, "max": 1.144259, "quartile": "Fourth"}, 15: {"min": 1.122011, "max": 1.162011, "quartile": "Fourth"}, 16: {"min": 1.139762, "max": 1.179762, "quartile": "Fourth"}, 17: {"min": 1.157513, "max": 1.197513, "quartile": "Fourth"}, 18: {"min": 1.175264, "max": 1.215264, "quartile": "Fourth"}, 19: {"min": 1.193016, "max": 1.230767, "quartile": "Fourth"}}, "GR17-20": {1: {"min": 0.754718, "max": 0.819758, "quartile": "First"}, 2: {"min": 0.794799, "max": 0.854799, "quartile": "First"}, 3: {"min": 0.829839, "max": 0.889839, "quartile": "First"}, 4: {"min": 0.864879, "max": 0.924879, "quartile": "Second"}, 5: {"min": 0.899919, "max": 0.959919, "quartile": "Second"}, 6: {"min": 0.93496, "max": 0.99496, "quartile": "Second"}, 7: {"min": 0.97, "max": 1.03, "quartile": "Third"}, 8: {"min": 0.998868, "max": 1.038868, "quartile": "Third"}, 9: {"min": 1.017736, "max": 1.057736, "quartile": "Third"}, 10: {"min": 1.036604, "max": 1.076604, "quartile": "Third"}, 11: {"min": 1.055471, "max": 1.095471, "quartile": "Third"}, 12: {"min": 1.074339, "max": 1.114339, "quartile": "Third"}, 13: {"min": 1.093207, "max": 1.133207, "quartile": "Fourth"}, 14: {"min": 1.112075, "max": 1.152075, "quartile": "Fourth"}, 15: {"min": 1.130943, "max": 1.170943, "quartile": "Fourth"}, 16: {"min": 1.149811, "max": 1.189811, "quartile": "Fourth"}, 17: {"min": 1.168679, "max": 1.208679, "quartile": "Fourth"}, 18: {"min": 1.187546, "max": 1.227546, "quartile": "Fourth"}, 19: {"min": 1.206414, "max": 1.245282, "quartile": "Fourth"}}}

QUARTILE_CHARACTERISTICS = {
    "First": "Meets minimum qualification of the job. However, may be fairly new to the job or field. Building both skills and knowledge as well as the ability to handle the full breadth of job duties and responsibilities. Employee is working towards proficiency in the job.",
    "Second": "Possesses all/most of the knowledge/skill requirements, but may need to build upon them through experience. Performs job responsibilities with increasing effectiveness. May still be learning some aspects of the job or developing expertise to handle the job more independently and effectively.",
    "Third": "Significant relevant experience and possesses all required knowledge and skills. Seasoned and proficient; consistently high-level of proficiency in all aspects of job over an extended period of time. Has broad and deep knowledge of own area as well as related areas.",
    "Fourth": "Expert in all job criteria; depth and breadth of experience, specialized skills, adds significant value to the University. Serves as expert resource and/or role model/mentor to others. This represents a premium on market salaries; typically reserved for employees with exceptional expertise or who have consistently demonstrated the highest levels of sustained contribution.",
}

PRELIM_NOTE = "This is a preliminary salary range and is intended only as a guide when considering internal equity and departmental budget constraints."
DISPOSITION_NOTE = "This artifact is decision-support only. Final qualification determination rests with the HR reviewer."
DISCLOSURE = "This analysis was assisted by the U of A HR MQ, Relevant Experience & Preliminary Salary Range Helper in AI VERDI. The HR reviewer must review and confirm all job architecture, equivalency, relevance, and preliminary salary calculations, as the final MQ status and salary range reflect HR judgment."

@dataclass
class SalaryResult:
    above_below_years: int
    quartile: str
    range_text: str
    quartile_descriptor: str
    quartile_characteristics: str
    employee_characteristics: str

@dataclass
class JobDescriptionParsed:
    job_code: str
    job_title: str
    job_level: str
    pay_grade: Optional[int]
    min_salary: Optional[float]
    midpoint: Optional[float]
    max_salary: Optional[float]
    required_degree: str
    education_requirement_text: str
    required_degree_years: int
    required_experience_years: int
    equivalency_allowed: bool
    family_stream_responsibilities: list[str]
    job_responsibilities: list[str]
    

def get_api_key() -> str:
    if "AIVERDE_API_KEY" in st.secrets:
        return st.secrets["AIVERDE_API_KEY"]
    return os.environ.get("AIVERDE_API_KEY", "")


def read_text_file(uploaded_file):
    name = uploaded_file.name.lower()
    raw = uploaded_file.read()
    uploaded_file.seek(0)
    if name.endswith(".pdf"):
        reader = PyPDF2.PdfReader(io.BytesIO(raw))
        pages = [(page.extract_text() or "") for page in reader.pages]
        return "\n".join(pages), {"type": "pdf", "pages": len(pages), "nonempty_pages": sum(1 for x in pages if x.strip()), "chars": sum(len(x) for x in pages)}
    if name.endswith(".docx"):
        doc = docx.Document(io.BytesIO(raw))
        paras = [p.text for p in doc.paragraphs]
        joined = "\n".join(paras)
        return joined, {"type": "docx", "paragraphs": len(paras), "nonempty_paragraphs": sum(1 for x in paras if x.strip()), "chars": len(joined)}
    if name.endswith(".txt"):
        joined = raw.decode("utf-8", errors="ignore")
        return joined, {"type": "txt", "chars": len(joined)}
    return None, {"type": "unknown", "chars": 0}

def extract_labeled_text(text: str, label: str) -> str:
    pattern = rf"{re.escape(label)}\s*[:\-]?\s*(.+)"
    match = re.search(pattern, text, re.IGNORECASE)
    if not match:
        return ""
    value = match.group(1).strip()
    value = value.split("\n")[0].strip()
    return value

def extract_money_field(text: str, label: str) -> Optional[float]:
    pattern = rf"{re.escape(label)}\s*\$?\s*([0-9,]+)"
    match = re.search(pattern, text, re.IGNORECASE)
    if not match:
        return None
    return float(match.group(1).replace(",", ""))

def extract_section_block(text: str, start_label: str, end_labels: list[str]) -> str:
    txt = text.replace("\r", "")
    start_match = re.search(re.escape(start_label), txt, re.IGNORECASE)
    if not start_match:
        return ""
    start = start_match.end()
    end = len(txt)
    for end_label in end_labels:
        m = re.search(re.escape(end_label), txt[start:], re.IGNORECASE)
        if m:
            end = min(end, start + m.start())
    return txt[start:end].strip()

def extract_bullets_from_block(block: str) -> list[str]:
    items = []
    current = ""
    for raw_line in block.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("•"):
            if current:
                items.append(current.strip())
            current = line.lstrip("•").strip()
        else:
            if current:
                current += " " + line
    if current:
        items.append(current.strip())

    cleaned = []
    seen = set()
    for item in items:
        key = item.lower()
        if key not in seen:
            seen.add(key)
            cleaned.append(item)
    return cleaned

def parse_jd_degree_years(label: str) -> int:
    s = (label or "").lower()
    if "high school" in s or "ged" in s:
        return 0
    if "doctor" in s or "ph.d" in s or "phd" in s:
        return 8
    if "master" in s:
        return 6
    if "bachelor" in s:
        return 4
    if "associate" in s:
        return 2
    return 0

def parse_jd_experience_years(experience_text: str) -> int:
    text = (experience_text or "").lower()
    match = re.search(r"minimum of\s+(\d+)\s+years", text)
    if match:
        return int(match.group(1))
    match = re.search(r"(\d+)\s+years", text)
    if match:
        return int(match.group(1))
    return 0

def parse_job_description_structured(job_description_text: str) -> JobDescriptionParsed:
    education_block = extract_section_block(
        job_description_text,
        "Education",
        ["Experience", "Family/Stream Principal Responsibilities", "Principal Responsibilities for the Job"]
    )

    experience_block = extract_section_block(
        job_description_text,
        "Experience",
        ["Family/Stream Principal Responsibilities", "Principal Responsibilities for the Job"]
    )

    family_block = extract_section_block(
        job_description_text,
        "Family/Stream Principal Responsibilities",
        ["Principal Responsibilities for the Job"]
    )

    job_block = extract_section_block(
        job_description_text,
        "Principal Responsibilities for the Job",
        []
    )

    required_degree = "Not stated"
    edu_line = " ".join([x.strip() for x in education_block.splitlines() if x.strip()])
    edu_lower = edu_line.lower()

    if "high school diploma" in edu_lower or "high school diploma equivalency" in edu_lower or "ged" in edu_lower:
        required_degree = "High school diploma"
    elif "associate" in edu_lower:
        required_degree = "Associate's degree"
    elif "bachelor" in edu_lower:
        required_degree = "Bachelor's degree"
    elif "master" in edu_lower:
        required_degree = "Master's degree"
    elif "doctor" in edu_lower or "ph.d" in edu_lower or "phd" in edu_lower:
        required_degree = "Doctoral degree"

    education_requirement_text = required_degree
    if "apprenticeship" in edu_lower or "technical training" in edu_lower:
        extras = []
        if "apprenticeship" in edu_lower:
            extras.append("apprenticeship program")
        if "technical training" in edu_lower:
            extras.append("technical training")
        if required_degree != "Not stated":
            education_requirement_text = f"{required_degree} + {' and '.join(extras)}"
        else:
            education_requirement_text = " and ".join(extras)

    exp_line = " ".join([x.strip() for x in experience_block.splitlines() if x.strip()])

    return JobDescriptionParsed(
        job_code=extract_labeled_text(job_description_text, "Job Code") or infer_job_code(job_description_text),
        job_title=extract_labeled_text(job_description_text, "Job Description Title") or infer_job_title(job_description_text),
        job_level=extract_labeled_text(job_description_text, "Job Level") or infer_job_level(job_description_text),
        pay_grade=safe_int(extract_labeled_text(job_description_text, "Pay Grade"), None),
        min_salary=extract_money_field(job_description_text, "Min"),
        midpoint=extract_money_field(job_description_text, "Mid"),
        max_salary=extract_money_field(job_description_text, "Max"),
        required_degree=required_degree,
        education_requirement_text=education_requirement_text,
        required_degree_years=parse_jd_degree_years(required_degree),
        required_experience_years=parse_jd_experience_years(exp_line),
        equivalency_allowed=(
            "equivalent combination of education and work experience" in exp_line.lower()
            or "equivalent advanced learning" in edu_line.lower()
        ),
        family_stream_responsibilities=extract_bullets_from_block(family_block),
        job_responsibilities=extract_bullets_from_block(job_block),
    )
    
def validate_extracted_text(label: str, text_value: str, meta: dict) -> Optional[str]:
    if not text_value or not text_value.strip():
        return f"Could not read the {label}. Please upload a readable PDF, DOCX, or TXT version."
    chars = meta.get("chars", 0)
    if label == "resume":
        if chars < 300:
            return "The resume may not have been read completely. Please upload a readable text-based version."
    else:
        if chars < 200:
            return "The job description may not have been read completely. Please upload a readable text-based version."
    if meta.get("type") == "pdf":
        pages = meta.get("pages", 0)
        nonempty_pages = meta.get("nonempty_pages", 0)
        if pages >= 3 and nonempty_pages == 0:
            return f"Could not read the {label} PDF. Please upload a readable text-based PDF or DOCX version."
        if pages >= 4 and nonempty_pages / max(pages, 1) < 0.5:
            return f"The {label} PDF appears substantially unreadable. Please upload a readable text-based PDF or DOCX version."
    if meta.get("type") == "docx" and meta.get("nonempty_paragraphs", 0) == 0:
        return f"The {label} DOCX appears unreadable. Please upload another version."
    return None


def call_aiverde(api_key: str, model_name: str, system_prompt: str, user_message: str) -> str:
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {"model": model_name, "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": user_message}], "temperature": 0}
    response = requests.post(CHAT_COMPLETIONS_URL, headers=headers, json=payload, timeout=300)
    response.raise_for_status()
    return response.json()["choices"][0]["message"]["content"]


def extract_structured_data(result_text: str) -> Optional[dict]:
    match = re.search(
        r"(?:<<>>|<<<STRUCTURED_DATA_START>>>)\s*(\{.*?\})\s*(?:<<>>|<<<STRUCTURED_DATA_END>>>)",
        result_text,
        re.DOTALL,
    )
    if not match:
        return None
    try:
        data = json.loads(match.group(1))
    except json.JSONDecodeError:
        return None
    if not isinstance(data, dict):
        return None
    return data

def strip_structured_data(result_text: str) -> str:
    return re.sub(r"\n?<<<STRUCTURED_DATA_START>>>.*?<<<STRUCTURED_DATA_END>>>\n?", "", result_text, flags=re.DOTALL).strip()

def infer_job_code(job_description_text: str) -> str:
    text = job_description_text or ""
    patterns = [
        r"\bjob\s*code\s*[:#-]?\s*([A-Z]\d{4,6})\b",
        r"\bjob\s*code\s*[:#-]?\s*([A-Z0-9-]{4,20})\b",
        r"\bposition\s*code\s*[:#-]?\s*([A-Z0-9-]{4,20})\b",
        r"\bworking\s*title.*?\b([A-Z]\d{4,6})\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return ""

def infer_job_title(job_description_text: str) -> str:
    text = (job_description_text or "").replace("\r", "")

    patterns = [
        r"\bjob\s*description\s*title\s*[:#-]?\s*([^\n]+)",
        r"\bworking\s*title\s*[:#-]?\s*([^\n]+)",
        r"\bclassification\s*title\s*[:#-]?\s*([^\n]+)",
        r"\btitle\s*[:#-]?\s*([^\n]+)",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            value = match.group(1).strip()
            value = re.split(r"\s{2,}|Job Code|Pay Grade|FLSA|Organization Level|Job Level", value, maxsplit=1)[0]
            value = value.strip(" |:-")
            if value and value.lower() not in {"job description", "not found"}:
                return value

    lines = [line.strip() for line in text.splitlines() if line.strip()]

    for i, line in enumerate(lines[:25]):
        if re.fullmatch(r"job\s*description", line, re.IGNORECASE):
            if i > 0:
                candidate = lines[i - 1].strip(" |:-")
                if candidate and len(candidate) <= 80:
                    return candidate

    return ""

def infer_job_level(job_description_text: str) -> str:
    text = job_description_text or ""
    patterns = [
        r"\bjob\s*level\s*[:#-]?\s*([A-Za-z0-9 /_-]+)",
        r"\blevel\s*[:#-]?\s*([A-Za-z0-9 /_-]+)",
        r"\bgrade\s*[:#-]?\s*([A-Za-z0-9 /_-]+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            value = match.group(1).strip()
            value = re.split(r"\s{2,}|\n|Pay Grade|FLSA|Job Code|Title", value, maxsplit=1)[0].strip(" |:-")
            if value:
                return value
    return ""

def build_review_metadata_line(data: dict, job_description_text: str) -> str:
    meta = as_dict(data.get("review_metadata"))
    jd = parse_job_description_structured(job_description_text)

    job_code = str(meta.get("job_code") or "").strip() or jd.job_code or "Not found"
    job_title = str(meta.get("job_title") or "").strip() or jd.job_title or "Not found"
    job_level = (
        str(meta.get("job_level") or "").strip()
        or jd.job_level
        or (f"Grade {jd.pay_grade}" if jd.pay_grade is not None else "")
        or "Not found"
    )

    return (
        f"Job Code: {job_code} | "
        f"Job Description Title: {job_title} | "
        f"Job Level: {job_level}"
    )

def infer_pay_grade(job_description_text: str) -> Optional[int]:
    text = job_description_text or ""
    patterns = [
        r"pay\s*grade\s*[:#-]?\s*(\d{1,2})",
        r"\bgrade\s*[:#-]?\s*(\d{1,2})\b",
        r"\bpay\s*level\s*[:#-]?\s*(\d{1,2})\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return int(match.group(1))
    return None


def infer_midpoint(job_description_text: str) -> Optional[float]:
    text = job_description_text or ""
    patterns = [
        r"mid(?:point)?\s*[:#-]?\s*\$?([0-9,]+)",
        r"salary\s*range\s*\$?([0-9,]+)",
        r"\bmid\s*\$?([0-9,]+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return float(match.group(1).replace(",", ""))
    return None


def normalize_followup_prompt(text: str) -> str:
    return re.sub(r"\n*Need details on any section, role, calculation, or assumption\?\s*", "\n", text, flags=re.IGNORECASE).strip()


def remove_model_section5_and_6(text: str) -> str:
    text = re.sub(r"\n---\n\n## 5\. PRELIMINARY SALARY RANGE.*?(?=\n---\n\n## 6\. PRELIMINARY DISPOSITION|$)", "", text, flags=re.DOTALL)
    text = re.sub(r"\n---\n\n## 6\. PRELIMINARY DISPOSITION.*$", "", text, flags=re.DOTALL)
    return text.strip()

def pct_to_whole_number(v, default=0):
    try:
        x = safe_float(v)
        if 0 <= x <= 1:
            x *= 100
        return safe_int(round(x))
    except Exception:
        return default


def safe_float(v, default=0.0):
    if v is None:
        return default
    if isinstance(v, str):
        text = v.strip().replace(",", "").replace("%", "")
        if not text or text.lower() in {"n/a", "na", "none", "null"}:
            return default
        try:
            return float(text)
        except Exception:
            return default
    try:
        return float(v)
    except Exception:
        return default


def safe_int(v, default=0):
    if v is None:
        return default
    if isinstance(v, str):
        text = v.strip().replace(",", "").replace("%", "")
        if not text or text.lower() in {"n/a", "na", "none", "null"}:
            return default
        try:
            return int(float(text))
        except Exception:
            return default
    try:
        return int(v)
    except Exception:
        return default

def as_dict(x):
    return x if isinstance(x, dict) else {}

def as_list(x):
    return x if isinstance(x, list) else []


def _collect_role_like_candidates(value):
    found = []
    if isinstance(value, dict):
        role_keys = {"role", "role_name", "title", "job_title", "position", "role_title", "name"}
        duration_keys = {"duration_years", "duration", "years", "experience_years", "years_of_experience", "tenure_years"}
        relevance_keys = {"relevance_pct", "relevance", "relevance_percentage", "allowed_relevance_pct", "allowed_relevance", "allowed_relevance_percentage", "allowed_pct", "adjusted_relevance_pct"}
        key_names = {str(k).lower() for k in value.keys()}
        has_role_key = any(k in role_keys for k in key_names)
        has_duration_key = any(k in duration_keys for k in key_names)
        has_relevance_key = any(k in relevance_keys for k in key_names)
        if has_role_key and (has_duration_key or has_relevance_key):
            found.append(value)
        for child in value.values():
            found.extend(_collect_role_like_candidates(child))
    elif isinstance(value, list):
        for item in value:
            found.extend(_collect_role_like_candidates(item))
    return found


def extract_role_candidates(data: dict) -> list[dict]:
    candidates = []
    for payload in [
        as_dict(data.get("experience_mq")),
        as_dict(data.get("experience_calculation")),
        as_dict(data),
    ]:
        roles = as_list(payload.get("roles") or payload.get("role_details") or payload.get("experience_roles") or payload.get("role_candidates"))
        if roles:
            return roles

        for key in ["roles", "role_details", "experience_roles", "role_candidates"]:
            if isinstance(payload.get(key), dict):
                nested_roles = as_list(payload.get(key).get("items") or payload.get(key).get("entries"))
                if nested_roles:
                    return nested_roles

        recursive_roles = _collect_role_like_candidates(payload)
        if recursive_roles:
            return recursive_roles

    return candidates


def compute_role_allowed_contribution(role: dict) -> float:
    explicit = safe_float(role.get("allowed_contribution_years"), None)
    if explicit is not None and explicit != 0.0:
        return round(explicit, 1)

    duration_years = safe_float(role.get("duration_years", 0.0))
    allowed_pct = safe_int(role.get("allowed_relevance_pct", None), safe_int(role.get("relevance_pct", 0), 0))
    return round(duration_years * (allowed_pct / 100.0), 1)


def compute_totals_from_roles(roles: list[dict]) -> tuple[float, int]:
    total = round(
        sum(
            compute_role_allowed_contribution(r)
            for r in roles
            if r.get("count_in_total", True)
        ),
        1,
    )
    rounded = 0 if total == 0 else safe_int(total) if safe_float(total).is_integer() else safe_int(total) + 1
    return total, rounded


def get_effective_counted_experience(data: dict) -> float:
    mq_summary = as_dict(data.get("mq_summary"))
    counted_value = safe_float(mq_summary.get("counted_experience_years"), None)
    if counted_value is None or counted_value <= 0:
        roles_total_precise, _ = compute_totals_from_roles(data.get("roles", []))
        return roles_total_precise
    return counted_value


def round_experience_years(value: float) -> int:
    total = safe_float(value, 0.0)
    if total <= 0:
        return 0
    return safe_int(total) if safe_float(total).is_integer() else safe_int(total) + 1


def degree_years_from_label(label: str) -> int:
    s = (label or "").lower()
    if "ph.d" in s or "doctor" in s:
        return 8
    if "master" in s or "m.s" in s or "m.a" in s or "mba" in s:
        return 6
    if "bachelor" in s or "b.s" in s or "b.a" in s or "bs " in s or "ba " in s:
        return 4
    if "associate" in s:
        return 2
    return 0


def infer_required_experience_years(job_description_text: str) -> Optional[int]:
    text = job_description_text or ""
    number_words = {
        "one": 1,
        "two": 2,
        "three": 3,
        "four": 4,
        "five": 5,
        "six": 6,
        "seven": 7,
        "eight": 8,
        "nine": 9,
        "ten": 10,
        "eleven": 11,
        "twelve": 12,
    }

    patterns = [
        r"(?:minimum of|at least|minimum|required)\s+(?P<value>\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve)\s*(?:\(\s*(?P<paren>\d+)\s*\))?\s+years?\b",
        r"\b(?P<value>\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve)\s*(?:\(\s*(?P<paren>\d+)\s*\))?\s+years?\b",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if not match:
            continue
        raw_value = match.group("value") or match.group("paren") or ""
        if not raw_value:
            continue
        if raw_value.isdigit():
            return int(raw_value)
        return number_words.get(raw_value.lower())

    if re.search(r"years?\s+(?:of|in|with).*(?:experience|experience in)", text, re.IGNORECASE):
        match = re.search(r"\b(?P<value>\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve)\b", text, re.IGNORECASE)
        if match:
            raw_value = match.group("value")
            if raw_value.isdigit():
                return int(raw_value)
            return number_words.get(raw_value.lower())

    return None


def infer_required_degree_from_job_description(job_description_text: str) -> tuple[str, int]:
    text = (job_description_text or "").lower()
    patterns = [
        ("doctor", "Doctoral degree", 8),
        ("ph.d", "Doctoral degree", 8),
        ("phd", "Doctoral degree", 8),
        ("master", "Master's degree", 6),
        ("bachelor", "Bachelor's degree", 4),
        ("associate", "Associate's degree", 2),
        ("high school diploma", "High school diploma", 0),
        ("ged", "High school diploma", 0),
        ("high school equivalency", "High school diploma", 0),
    ]
    for token, label, years in patterns:
        if token in text:
            return label, years
    return "Not stated", 0


def infer_candidate_degree_from_resume_text(resume_text: str) -> tuple[str, int]:
    text = (resume_text or "").lower()
    patterns = [
        ("ph.d", "PhD", 8),
        ("phd", "PhD", 8),
        ("doctorate", "Doctoral degree", 8),
        ("master's", "Master's degree", 6),
        ("master of", "Master's degree", 6),
        ("m.s.", "M.S.", 6),
        ("m.a.", "M.A.", 6),
        ("mba", "MBA", 6),
        ("bachelor's", "Bachelor's degree", 4),
        ("bachelors", "Bachelor's degree", 4),
        ("b.s.", "B.S.", 4),
        ("b.a.", "B.A.", 4),
        ("associate's", "Associate's degree", 2),
        ("associate degree", "Associate's degree", 2),
        ("high school diploma", "High school diploma", 0),
        ("ged", "High school diploma", 0),
    ]
    for token, label, years in patterns:
        if token in text:
            return label, years
    return "Not stated", 0


def candidate_degree_years(data: dict) -> int:
    mq_summary = data.get("mq_summary", {})
    degree_label = str(mq_summary.get("highest_completed_degree") or "").strip()
    if not degree_label or degree_label.lower() in {"not stated", "unclear", "n/a", "na", "none", "null"}:
        return 0
    explicit = safe_int(mq_summary.get("highest_completed_degree_years", 0))
    mapped = degree_years_from_label(degree_label)
    if mapped > 0:
        if explicit <= 0:
            return mapped
        if explicit != mapped:
            return mapped
        return explicit
    return explicit if explicit > 0 else mapped


def render_section1_from_data(data: dict) -> str:
    target_role = data.get("target_role", {})
    mq_summary = data.get("mq_summary", {})
    roles_total_precise, roles_total_rounded = compute_totals_from_roles(data.get("roles", []))
    total_precise = get_effective_counted_experience(data)
    total_rounded = round_experience_years(total_precise)
    required_degree = mq_summary.get("required_degree") or "Not stated"
    required_degree_years = safe_int(mq_summary.get("required_degree_years"), 0)
    highest_degree = mq_summary.get("highest_completed_degree") or "Not stated"
    education_evidence = mq_summary.get("education_evidence_text") or highest_degree
    highest_degree_years = safe_int(candidate_degree_years(data),0)
    required_exp = safe_int(mq_summary.get("required_experience_years", 0))
    required_combined = safe_int(required_degree_years, 0) + safe_int(required_exp, 0)
    candidate_combined = safe_int(highest_degree_years, 0) + safe_int(total_rounded, 0)
    if required_degree_years <= 0:
        degree_status = "ℹ️ N/A"
    elif highest_degree_years <= 0:
        degree_status = "⚠️ Does not clearly meet"
    elif highest_degree_years > required_degree_years:
        degree_status = "✅ Exceeds"
    else:
        degree_status = "✅ Met"
    exp_status = "✅ Met" if total_rounded >= required_exp and required_exp > 0 else "⚠️ Does not clearly meet"
    equiv_needed = "Yes" if total_rounded < required_exp else "No"
    equiv_status = "✅ Applied" if mq_summary.get("mq_final_status") in {"MET", "EXCEEDS"} and equiv_needed == "Yes" else "ℹ️ N/A"
    combined_status = "✅ Met" if candidate_combined >= required_combined else "⚠️ Does not clearly meet"
    oc_text = "Yes" if target_role.get("oc_level_applies") is True else "No" if target_role.get("oc_level_applies") is False else "Unclear"
    overlap_rule = "25%" if target_role.get("oc_level_applies") is True else "0%" if target_role.get("oc_level_applies") is False else "unclear"
    equiv_allowed = "Yes" if target_role.get("equivalency_allowed") else "No"
 
    return (
        "## 1. MINIMUM QUALIFICATIONS SNAPSHOT\n\n"
        "| MQ Element | Requirement | Candidate Evidence | Status |\n"
        "|---|---|---|---|\n"
        f"| Education | {required_degree} | {education_evidence} | {degree_status} |\n"
        f"| Relevant Experience | {required_exp} years (or equiv. combo) | See Section 2 ({total_precise:.1f} years ~ {total_rounded}+ years allowed) | {exp_status} |\n"
        f"| Combined Education + Relevant Experience | {required_combined} total years | {candidate_combined} total years ({highest_degree_years} education + {total_rounded} relevant experience) | {combined_status} |\n"
        f"| Equivalency Needed | {equiv_needed} | Equivalency allowed per JD: {equiv_allowed} | {equiv_status} |\n\n"
        f"- Equivalency allowed per JD: {equiv_allowed} (equivalent combination of education and work experience)\n"
        f"- OC-level role: {oc_text} → overlapping full-time school work counts at {overlap_rule}"
    )


def render_section2_from_roles(data: dict) -> str:
    roles = data.get("roles", [])
    notes = data.get("section2_notes", []) or []
    total_precise, total_rounded = compute_totals_from_roles(roles)
    lines = [
        "## 2. EXPERIENCE CALCULATION",
        "",
        "| Role | Dates | Duration | Relevance | School Overlap | Allowed Relevance |",
        "|---|---|---|---|---|---|",
    ]
    for r in roles:
        role_name = r.get("role_name") or r.get("role") or r.get("title") or r.get("job_title") or r.get("position") or ""
        dates_text = r.get("dates_text") or r.get("dates") or r.get("date_range") or ""
        duration_years = safe_float(r.get("duration_years") or r.get("years") or r.get("duration") or r.get("experience_years"), 0.0)
        relevance_pct = safe_int(r.get("relevance_pct") or r.get("relevance") or r.get("relevance_percentage") or r.get("alignment_pct"), 0)
        school_overlap_text = r.get("school_overlap_text") or ("Yes" if r.get("school_overlap") is True else "No" if r.get("school_overlap") is False else "Unclear")
        allowed_pct = safe_int(r.get("allowed_relevance_pct") or r.get("allowed_relevance") or r.get("allowed_relevance_percentage") or r.get("allowed_pct") or r.get("adjusted_relevance_pct"), relevance_pct)
        allowed_years = compute_role_allowed_contribution(r)
        lines.append(f"| {role_name} | {dates_text} | ~{duration_years:.1f} yrs | {relevance_pct}% | {school_overlap_text} | {allowed_pct}% → {allowed_years:.1f} yrs |")
    lines.append("")
    for note in [n for n in notes if n][:1]:
        lines.append(note)
    lines.append("")
    lines.append(f"Relevant experience total: {total_precise:.1f} years ~ {total_rounded}+ years")
    return "\n".join(lines).strip()


def parse_principal_responsibilities(job_description_text: str) -> list[str]:
    jd = parse_job_description_structured(job_description_text)
    return jd.job_responsibilities


def build_candidate_evidence_for_req(req: str, resume_text: str) -> tuple[str, str]:
    sentences = re.split(r"(?<=[.!?])\s+|\n+", resume_text)
    keywords = [w.lower() for w in re.findall(r"[A-Za-z]{4,}", req) if w.lower() not in {"with", "that", "this", "from", "will", "have", "more", "such", "their", "using", "within", "shared", "resources", "environment"}]
    scored = []
    for s in sentences:
        ss = s.strip()
        if len(ss) < 25:
            continue
        score = sum(1 for k in keywords if k in ss.lower())
        if score:
            scored.append((score, ss))
    scored.sort(key=lambda x: (-x[0], len(x[1])))
    if scored:
        evidence = scored[0][1]
        score = scored[0][0]
        match = "High" if score >= 3 else "Moderate" if score >= 2 else "Possible"
        return evidence[:220], match
    return "No direct resume evidence clearly identified.", "Unclear"

def mq_allows_salary(data: dict) -> bool:
    mq_summary = as_dict(data.get("mq_summary"))
    section6 = as_dict(data.get("section6_summary"))
    prelim = as_dict(data.get("preliminary_disposition"))

    mq_status = str(mq_summary.get("mq_final_status") or "").upper()
    edu_text = str(section6.get("mq_education_text") or "").upper()
    exp_text = str(section6.get("mq_experience_text") or "").upper()
    prelim_text = str(prelim.get("overall_disposition") or "").upper()

    if mq_status in {"EXCEEDS", "MET"}:
        return True

    required_degree_years = safe_int(mq_summary.get("required_degree_years"), 0)
    required_experience_years = safe_int(mq_summary.get("required_experience_years"), 0)
    candidate_degree_years_value = candidate_degree_years(data)
    counted_experience_years = get_effective_counted_experience(data)
    rounded_experience_years = round_experience_years(counted_experience_years)

    required_total = required_degree_years + required_experience_years
    candidate_total = candidate_degree_years_value + rounded_experience_years

    if required_total > 0 and candidate_total >= required_total:
        return True

    if ("EXCEEDS" in edu_text or "MET" in edu_text) and (
        "PENDING" in exp_text or "CONDITIONAL" in prelim_text or "UNCLEAR" in prelim_text
    ):
        return True

    if "MET" in exp_text or "MEETS" in exp_text:
        return True

    return False

def format_quartile_descriptor(quartile: str, above_below_years: int) -> str:
    if quartile == "First":
        return "1st Quartile (at or near minimum) — meets MQ at threshold; limited verified post-degree experience above requirement" if above_below_years <= 0 else "1st Quartile — modestly above minimum; early progression within range"
    if quartile == "Second":
        return "2nd Quartile — developing proficiency with experience above threshold"
    if quartile == "Third":
        return "3rd Quartile — strong credentials and specialized experience"
    return "4th Quartile — exceptional depth and sustained contribution"


def build_employee_characteristics(mq_summary: dict, quartile: str) -> str:
    degree_name = mq_summary.get("highest_completed_degree") or "advanced degree"
    if quartile == "First":
        return f"New to role at this level; highly credentialed ({degree_name}) with strong technical/research depth; limited verified independent post-degree experience; strong publication and presentation record"
    if quartile == "Third":
        return f"Highly credentialed ({degree_name}); strong specialized technical/research depth; substantial alignment to core duties"
    return QUARTILE_CHARACTERISTICS[quartile]


def compute_salary_result(pay_grade: int, midpoint: float, above_below_years: int, mq_summary: dict) -> Optional[SalaryResult]:
    suffix = GRADE_SUFFIX_MAP.get(pay_grade)
    if not suffix:
        return None
    bucket = min(max(int(above_below_years), 1), 19)
    row = CO_RATIO_TABLE.get(suffix, {}).get(bucket)
    if not row:
        return None
    quartile = row["quartile"]
    min_salary = row["min"] * midpoint
    max_salary = row["max"] * midpoint
    return SalaryResult(
        above_below_years=above_below_years,
        quartile=quartile,
        range_text=f"{min_salary:,.0f} – {max_salary:,.0f}",
        quartile_descriptor=format_quartile_descriptor(quartile, above_below_years),
        quartile_characteristics=QUARTILE_CHARACTERISTICS[quartile],
        employee_characteristics=build_employee_characteristics(mq_summary, quartile),
    )

    if DEBUG:
        with st.expander("Debug: Section 5 gate"):
            st.write("data keys", list(data.keys()))
            st.write("mq_summary", data.get("mq_summary"))
            st.write("section6_summary", data.get("section6_summary"))
            st.write("preliminary_disposition", data.get("preliminary_disposition"))

def build_salary_section(data: dict, job_description_text: str) -> tuple[str, str]:
    mq_summary = data.get("mq_summary", {})

    if DEBUG:
        with st.expander("Debug: Section 5 gate"):
            st.write("mq_summary:", mq_summary)
            st.write("section6_summary:", data.get("section6_summary", {}))
            st.write("preliminary_disposition:", data.get("preliminary_disposition", {}))
            st.write("mq_final_status:", mq_summary.get("mq_final_status"))
            st.write("mq_allows_salary:", mq_allows_salary(data))

    if not mq_allows_salary(data):
        return "## 5. PRELIMINARY SALARY RANGE\n\nN/A — candidate does not clearly meet MQ or MQ is unclear.", "N/A"

    jd = parse_job_description_structured(job_description_text)
    pay_grade = jd.pay_grade
    midpoint = jd.midpoint
    if pay_grade is None:
        pay_grade = safe_int(data.get("pay_grade"), None)
    if midpoint is None:
        midpoint = safe_float(data.get("midpoint"), None)
    required_degree_years = mq_summary.get("required_degree_years")
    required_experience_years = mq_summary.get("required_experience_years")
    candidate_degree = candidate_degree_years(data)
    total_precise, total_rounded = compute_totals_from_roles(data.get("roles", []))
    degree_name = mq_summary.get("highest_completed_degree", "Degree")

    if None in (pay_grade, midpoint, required_degree_years, required_experience_years):
        return "## 5. PRELIMINARY SALARY RANGE\n\nN/A — preliminary salary range could not be determined.", "N/A"

    if candidate_degree == 0 and safe_int(required_degree_years, 0) > 0:
        candidate_degree = safe_int(required_degree_years, 0)

    required_total = safe_int(required_degree_years) + safe_int(required_experience_years)
    candidate_total = safe_int(candidate_degree) + safe_int(total_rounded)
    above_below = candidate_total - required_total

    salary = compute_salary_result(pay_grade, midpoint, above_below, mq_summary)
    if not salary:
        return "## 5. PRELIMINARY SALARY RANGE\n\nN/A — salary lookup failed.", "N/A"

    if above_below == 0:
        above_text = f"Meets requirement at threshold via equivalency ({degree_name} = {candidate_degree} yrs + {total_rounded} yr relevant exp = {candidate_total} yrs; requirement = {required_total} yrs)."
    elif above_below < 0:
        above_text = f"Candidate was determined to meet MQ via equivalency judgment; numeric combination is approximately {abs(above_below)} year(s) below the standard total threshold."
    else:
        above_text = f"{degree_name} ({candidate_degree} yrs equiv.) + ~{total_precise:.1f} yrs relevant experience = ~{candidate_degree + total_precise:.1f} yrs combined vs. required {required_total} yrs. Candidate is approximately {above_below}+ year(s) above the combined minimum."

    range_suffix = " (Min to ~1st Quartile)" if salary.quartile == "First" else ""

    section = (
        "## 5. PRELIMINARY SALARY RANGE\n\n"
        "| | |\n|---|---|\n"
        f"| **Education & Experience Above (Below) Requirement** | {above_text} |\n"
        f"| **Pay Grade** | Grade {pay_grade} |\n"
        f"| **Midpoint** | ${midpoint:,.0f} |\n"
        f"| **Quartile** | {salary.quartile_descriptor} |\n"
        f"| **Quartile Characteristics** | {salary.quartile_characteristics} |\n"
        f"| **Preliminary Salary Range** | {salary.range_text}{range_suffix} |\n"
        f"> ⚠️ *{PRELIM_NOTE}*"
    )
    return section, salary.range_text


def render_section6_from_data(data: dict, salary_range_text: str) -> str:
    mq_summary = data.get("mq_summary", {})
    roles_total_precise, roles_total_rounded = compute_totals_from_roles(data.get("roles", []))
    total_precise = safe_float(mq_summary.get("counted_experience_years"), roles_total_precise)
    total_rounded = round_experience_years(total_precise)

    mq_education = mq_summary.get("required_degree") or "Not stated"
    required_exp = safe_int(mq_summary.get("required_experience_years"), 0)
    mq_experience = f"{required_exp} years (or equiv. combo)" if required_exp > 0 else "Not stated"

    highest_degree = mq_summary.get("highest_completed_degree", "") or "Not stated"
    highest_degree_years = safe_int(candidate_degree_years(data), 0)
    required_degree_years = safe_int(mq_summary.get("required_degree_years"), 0)
    required_combined = required_degree_years + required_exp
    candidate_combined = highest_degree_years + total_rounded
    combined_evidence = f"{candidate_combined} total years ({highest_degree_years} education + {total_rounded} relevant experience); requirement: {required_combined} total years"
    alignment_text = "See Section 3"

    return (
        "## 6. PRELIMINARY DISPOSITION\n\n"
        "| | |\n|---|---|\n"
        f"| **MQ Education** | {mq_education} |\n"
        f"| **MQ Experience** | {mq_experience} |\n"
        f"| **Highest Completed Degree** | {highest_degree} |\n"
        f"| **Relevant Experience Total** | {total_precise:.1f} years ~ {total_rounded}+ years |\n"
        f"| **Combined Education + Relevant Experience** | {combined_evidence} |\n"
        f"| **Relevant Experience Alignment** | {alignment_text} |\n"
        f"| **Preliminary Salary Range** | {salary_range_text} |\n\n"
        f"> ⚠️ *{DISPOSITION_NOTE}*\n\n{DISCLOSURE}"
    )

def render_section3_from_data(data: dict, job_description_text: str, resume_text: str) -> str:
    alignment_block = (
        as_list(data.get("relevant_experience_alignment"))
        or as_list(data.get("responsibility_alignment"))
        or as_list(data.get("responsibilities"))
        or as_list(data.get("principal_responsibilities"))
        or as_list(as_dict(data.get("experience_mq")).get("relevant_experience_alignment"))
        or as_list(as_dict(data.get("experience_mq")).get("principal_responsibilities"))
        or as_list(as_dict(data.get("experience_mq")).get("responsibilities"))
        or []
    )
    rows = []

    jd_principals = parse_principal_responsibilities(job_description_text)
    jd_iter = iter(jd_principals)

    for item in alignment_block:
        if isinstance(item, str):
            req = item.strip()
            evidence = ""
            match = ""
        elif isinstance(item, dict):
            req = str(
                item.get("jd_responsibility")
                or item.get("responsibility")
                or item.get("job_requirement")
                or item.get("requirement")
                or ""
            ).strip()

            evidence = str(
                item.get("resume_evidence")
                or item.get("evidence")
                or item.get("candidate_evidence")
                or ""
            ).strip()

            match = str(
                item.get("alignment")
                or item.get("match")
                or ""
            ).strip()
        else:
            continue

        if not req:
            try:
                req = next(jd_iter)
            except StopIteration:
                req = "Principal responsibility (see JD)"

        if req or evidence or match:
            rows.append(
                f"| {req} | {evidence or 'No direct resume evidence clearly identified.'} | {match or 'Unclear'} |"
            )

    if not rows and jd_principals:
        for req in jd_principals:
            ev, match = build_candidate_evidence_for_req(req, resume_text)
            rows.append(f"| {req} | {ev} | {match} |")
    elif not rows:
        rows.append(
            "| See job description principal responsibilities | See candidate experience summarized by model | Review narrative below |"
        )

    lines = [
        "## 3. RELEVANT EXPERIENCE ALIGNMENT",
        "",
        "| Job Requirement | Candidate Evidence | Match |",
        "|---|---|---|",
        *rows,
    ]
    return "\n".join(lines)


def render_section4_from_data(data: dict) -> str:
    flags = data.get("flags", []) or []
    lines = [
        "## 4. FLAGS FOR HR REVIEWER",
        "",
        "| # | Flag | Detail |",
        "|---|---|---|",
    ]

    if flags:
        for i, f in enumerate(flags, start=1):
            if not isinstance(f, dict):
                continue
            flag_label = str(f.get("flag") or f"Flag {i}").strip()
            detail = str(f.get("detail") or "").strip()
            lines.append(f"| {i} | {flag_label} | {detail} |")
    else:
        lines.append("| 1 | Review output | HR reviewer should verify degree relevance, chronology, and equivalency judgment. |")

    return "\n".join(lines)

def run_analysis(followup_text: str = ""):
    missing = []
    api_key = get_api_key()
    if not api_key:
        missing.append("AI-VERDE API key in .streamlit/secrets.toml or AIVERDE_API_KEY environment variable")
    if not st.session_state.get("resume_file"):
        missing.append("Candidate/Incumbent resume")
    if not st.session_state.get("job_description_file"):
        missing.append("University job description")
    if missing:
        st.error("Please resolve the following before running:\n\n" + "\n".join(f"• {m}" for m in missing))
        return

    resume_text, resume_meta = read_text_file(st.session_state["resume_file"])
    job_description_text, jd_meta = read_text_file(st.session_state["job_description_file"])

    resume_issue = validate_extracted_text("resume", resume_text, resume_meta)
    jd_issue = validate_extracted_text("job description", job_description_text, jd_meta)
    if resume_issue or jd_issue:
        problems = [x for x in [resume_issue, jd_issue] if x]
        st.error("Analysis not completed.\n\n" + "\n".join(f"• {p}" for p in problems))
        return

    with st.spinner("Running HR review through AI-VERDE (JSON pass)..."):
        jd = parse_job_description_structured(job_description_text)

        structured_jd_payload = {
            "job_code": jd.job_code,
            "job_title": jd.job_title,
            "job_level": jd.job_level,
            "pay_grade": jd.pay_grade,
            "midpoint": jd.midpoint,
            "required_degree": jd.required_degree,
            "required_degree_years": jd.required_degree_years,
            "required_experience_years": jd.required_experience_years,
            "equivalency_allowed": jd.equivalency_allowed,
            "principal_responsibilities": jd.job_responsibilities,
            "family_stream_responsibilities": jd.family_stream_responsibilities,
        }

        json_user_message = f"""
HR reviewer instruction: {st.session_state['reviewer_instruction']}

Use this structured job description data as source of truth for job requirements and responsibilities:

{json.dumps(structured_jd_payload, indent=2)}

Candidate resume text:

{resume_text}

Return only the required structured JSON wrapper, no prose.
"""

        json_text = call_aiverde(get_api_key(), DEFAULT_MODEL, STRUCTURED_SYSTEM_PROMPT, json_user_message)
        data = extract_structured_data(json_text)

        if not data and not followup_text.strip():
            repair_message = json_user_message + "\nEarlier response did not include valid JSON. Return valid JSON only inside the wrapper.\n"
            json_text = call_aiverde(get_api_key(), DEFAULT_MODEL, STRUCTURED_SYSTEM_PROMPT, repair_message)
            data = extract_structured_data(json_text)

    st.session_state["last_json_text"] = json_text
    st.session_state["last_structured_parse"] = data
    st.session_state["raw_structured_data"] = data

    if not data:
        st.error("The analysis could not be completed because the model did not return valid structured data even after a retry. Please rerun the analysis.")
        return

    if not isinstance(data, dict):
        if DEBUG:
            with st.expander("Debug: parsed structured data type"):
                st.write(type(data).__name__)
                st.write(data)
        st.error("Parsed structured data was not a dictionary.")
        return

    def as_dict(x):
        return x if isinstance(x, dict) else {}

    def as_list(x):
        return x if isinstance(x, list) else []

    minq = as_dict(data.get("minimum_qualifications"))
    exp_calc = as_dict(data.get("experience_calculation"))
    prelim = as_dict(data.get("preliminary_disposition"))
    meta = as_dict(data.get("review_metadata"))

    edu_block = as_dict(data.get("education_mq") or {})
    exp_block = as_dict(data.get("experience_mq") or {})
    roles_raw = extract_role_candidates(data)
    if not roles_raw:
        roles_raw = as_list(exp_block.get("roles") or [])
    
    jd = parse_job_description_structured(job_description_text)

    required_degree = jd.education_requirement_text or jd.required_degree or "Not stated"
    required_degree_years = safe_int(
        edu_block.get("required_degree_years") or edu_block.get("years_required") or edu_block.get("degree_years"),
        jd.required_degree_years,
    )
    if required_degree_years <= 0:
        required_degree_years = jd.required_degree_years  

    candidate_degree_entries = as_list(
        edu_block.get("degrees")
        or edu_block.get("education_entries")
        or edu_block.get("education")
        or edu_block.get("candidate_degrees")
        or data.get("education_entries")
        or data.get("degrees")
        or []
    )
    completed_degree_entries = []
    for entry in candidate_degree_entries:
        if not isinstance(entry, dict):
            continue
        degree_name = str(entry.get("degree") or entry.get("name") or entry.get("degree_name") or "").strip()
        completed = entry.get("completed")
        if not degree_name:
            continue
        if completed is False:
            continue
        school_name = str(entry.get("school") or entry.get("school_name") or entry.get("institution") or "").strip()
        if not school_name:
            continue
        explicit_years = safe_int(entry.get("years") or entry.get("years_required") or entry.get("degree_years"), 0)
        mapped_years = degree_years_from_label(degree_name)
        if mapped_years > 0 and (explicit_years <= 0 or explicit_years != mapped_years):
            explicit_years = mapped_years
        completed_degree_entries.append((degree_name, explicit_years))

    if completed_degree_entries:
        highest_completed_degree, highest_completed_degree_years = max(completed_degree_entries, key=lambda item: (item[1], item[0]))
    else:
        highest_completed_degree = str(edu_block.get("highest_completed_degree") or edu_block.get("candidate_degree") or "").strip() or ""
        highest_completed_degree_years = safe_int(
            edu_block.get("highest_completed_degree_years")
            or edu_block.get("candidate_degree_years")
            or edu_block.get("degree_years")
            or 0,
            0,
        )
    if not highest_completed_degree:
        highest_completed_degree = ""
        highest_completed_degree_years = 0

    training_evidence_candidates = [
        edu_block.get("training_evidence"),
        edu_block.get("other_training"),
        edu_block.get("non_degree_training"),
        edu_block.get("training"),
    ]
    training_evidence_text = ""
    for candidate in training_evidence_candidates:
        if isinstance(candidate, str) and candidate.strip():
            training_evidence_text = candidate.strip()
            break
        if isinstance(candidate, list):
            values = [str(x).strip() for x in candidate if str(x).strip()]
            if values:
                training_evidence_text = "; ".join(values)
                break
    education_evidence_text = highest_completed_degree or "Not stated"
    if training_evidence_text:
        education_evidence_text = f"{education_evidence_text}; {training_evidence_text}" if education_evidence_text and education_evidence_text != "Not stated" else training_evidence_text
    
    required_experience_years = safe_float(
    exp_block.get("requirement_years"),
    float(jd.required_experience_years),
)

    # Model’s own counted experience in Section 1
    direct_allowed_years = safe_float(
        exp_block.get("relevant_experience_total_years"),
        0.0,
    )
    computed_role_total = 0.0
    if roles_raw:
        computed_role_total = compute_totals_from_roles([
            {
                "duration_years": safe_float(r.get("duration_years"), 0.0),
                "allowed_relevance_pct": pct_to_whole_number(r.get("allowed_relevance_pct"), pct_to_whole_number(r.get("relevance_pct"), 0)),
                "relevance_pct": pct_to_whole_number(r.get("relevance_pct"), 0),
                "allowed_contribution_years": safe_float(r.get("allowed_contribution_years"), 0.0),
            }
            for r in roles_raw
            if isinstance(r, dict)
        ])[0]
    if direct_allowed_years <= 0 and computed_role_total > 0:
        direct_allowed_years = computed_role_total
    degree_equivalency_years = safe_float(
        exp_block.get("equivalency_total_years"),
        0.0,
    )

    combined_total_years = direct_allowed_years + degree_equivalency_years
    total_counted_years = direct_allowed_years

    mq_final_raw = str(
        minq.get("overall_mq_status")
        or prelim.get("overall_disposition")
        or ""
    ).upper()

    if "EXCEEDS" in mq_final_raw:
        mq_final_status = "EXCEEDS"
    elif "MET" in mq_final_raw:
        mq_final_status = "MET"
    elif "CONDITIONAL" in mq_final_raw or "MARGINAL" in mq_final_raw or "PENDING" in mq_final_raw:
        mq_final_status = "UNCLEAR"
    elif "NOT CLEARLY MEET" in mq_final_raw or "NOT MET" in mq_final_raw:
        mq_final_status = "DOES NOT CLEARLY MEET"
    else:
        mq_final_status = "UNCLEAR"

    roles = []
    bad_roles = []

    is_oc_role = bool((jd.job_code or "").upper().startswith("OC")) or bool(
        re.match(r"^OC[1-4]\b", (jd.job_level or "").upper())
    )

    for i, r in enumerate(roles_raw):
        if not isinstance(r, dict):
            bad_roles.append({"index": i, "type": type(r).__name__, "value": r})
            continue

        relevance_pct = r.get("relevance_pct")
        raw_relevance_pct = pct_to_whole_number(relevance_pct, 0)

        overlap_raw = str(
            r.get("school_overlap_text")
            or r.get("school_overlap")
            or ""
        ).strip().lower()

        if overlap_raw in {"yes", "y", "true"}:
            overlap_text = "Yes"
        elif overlap_raw in {"no", "n", "false"}:
            overlap_text = "No"
        elif overlap_raw == "partial":
            overlap_text = "Partial"
        else:
            overlap_text = "Unclear"

        duration_years = safe_float(
            r.get("duration_years") or r.get("years") or r.get("duration") or r.get("experience_years"),
            0.0,
        )

        explicit_allowed_pct = r.get("allowed_relevance_pct")
        explicit_allowed_years = r.get("allowed_contribution_years")

        if explicit_allowed_pct not in (None, "", "null"):
            allowed_relevance_pct = pct_to_whole_number(explicit_allowed_pct, raw_relevance_pct)
        elif overlap_text == "Yes":
            allowed_relevance_pct = round(raw_relevance_pct * 0.25) if is_oc_role else 0
        elif overlap_text == "Partial":
            allowed_relevance_pct = raw_relevance_pct
        else:
            allowed_relevance_pct = raw_relevance_pct

        if explicit_allowed_years not in (None, "", "null"):
            allowed_contribution_years = round(safe_float(explicit_allowed_years, 0.0), 1)
            if duration_years > 0 and allowed_relevance_pct == raw_relevance_pct:
                allowed_relevance_pct = safe_int(round((allowed_contribution_years / duration_years) * 100), raw_relevance_pct)
        else:
            allowed_contribution_years = round(
                duration_years * (allowed_relevance_pct / 100.0),
                1,
            )

            roles.append({
                "role_name": r.get("role_name") or r.get("role") or r.get("job_title") or r.get("title") or r.get("position") or "",
                "dates_text": r.get("dates_text") or r.get("dates") or r.get("date_range") or "",
                "duration_years": duration_years,
                "relevance_pct": raw_relevance_pct,
                "school_overlap_text": overlap_text,
                "allowed_relevance_pct": allowed_relevance_pct,
                "allowed_contribution_years": allowed_contribution_years,
                "count_in_total": True,
            })

        roles.append({
            "role_name": r.get("role_name") or r.get("role") or r.get("job_title") or r.get("title") or r.get("position") or "",
            "dates_text": r.get("dates_text") or r.get("dates") or r.get("date_range") or "",
            "duration_years": duration_years,
            "relevance_pct": raw_relevance_pct,
            "school_overlap_text": overlap_text,
            "allowed_relevance_pct": allowed_relevance_pct,
            "allowed_contribution_years": allowed_contribution_years,
            "count_in_total": True,
        })

    effective_counted_experience = compute_totals_from_roles(roles)[0] if roles else direct_allowed_years

    adapted = {
        "mq_summary": {
            "required_degree": jd.education_requirement_text,
            "required_degree_years": safe_int(required_degree_years, 0),
            "highest_completed_degree": highest_completed_degree,
            "highest_completed_degree_years": safe_int(highest_completed_degree_years, 0),
            "education_evidence_text": education_evidence_text,
            "required_experience_years": safe_int(required_experience_years, 0),
            "counted_experience_years": safe_float(effective_counted_experience, 0.0),
            "combined_experience_years": safe_float(effective_counted_experience + degree_equivalency_years, 0.0),
            "degree_equivalency_years": safe_float(degree_equivalency_years, 0.0),
            "mq_final_status": mq_final_status,
        },
        "roles": roles,
        "target_role": {
            "equivalency_allowed": jd.equivalency_allowed,
            "oc_level_applies": is_oc_role,
        },
        "section2_notes": [
            x for x in [
                exp_block.get("note"),
                exp_calc.get("summary_label"),
            ] if x
        ],
        "section6_summary": {
            "mq_education_text": edu_block.get("status") or "",
            "mq_experience_text": exp_block.get("status") or "",
            "highest_degree_text": highest_completed_degree or "",
            "employee_characteristics_text": prelim.get("employee_characteristics") or "",
            "alignment_text": "See Section 3",
        },
        "preliminary_disposition": prelim,
        "relevant_experience_alignment": as_list(data.get("relevant_experience_alignment")),
        "flags": as_list(data.get("flags")),
    }

    if DEBUG:
        with st.expander("Debug: adapter inputs"):
            st.write("required_degree", required_degree)
            st.write("required_experience_years", required_experience_years)
            st.write("total_counted_years", total_counted_years)
            st.write("edu_block", edu_block)
            st.write("exp_block", exp_block)

    st.session_state["structured_data"] = adapted
    data = adapted

    if DEBUG:
        with st.expander("Debug: adapter check"):
            st.write("minimum_qualifications:", minq)
            st.write("experience_calculation:", exp_calc)
            st.write("review_metadata:", meta)
            st.write("roles_raw count:", len(roles_raw))
            st.write("bad roles skipped:", bad_roles)
            st.write("adapted mq_summary:", adapted["mq_summary"])
            st.write("adapted roles count:", len(adapted["roles"]))
          
    metadata_line = build_review_metadata_line(data, job_description_text)

    visible = "## HR MQ, Relevant Experience & Preliminary Salary Range Review"
    visible += "\n\n" + metadata_line
    visible += "\n\n" + render_section1_from_data(data)
    visible += "\n\n" + render_section2_from_roles(data)

    section3 = render_section3_from_data(data, job_description_text, resume_text)
    section4 = render_section4_from_data(data)
    salary_section, salary_range_text = build_salary_section(data, job_description_text)
    section6 = render_section6_from_data(data, salary_range_text)

    full_output = "\n\n---\n\n".join([visible, section3, section4, salary_section, section6]).strip()
    st.session_state["final_output"] = full_output
    st.success("✅ Review complete")

with st.sidebar:
    st.header("Settings")
    st.caption(f"Analysis date: {date.today().isoformat()}")
    st.text_input("AI-VERDE model", value=DEFAULT_MODEL, disabled=True)
    st.session_state.setdefault(
    "reviewer_instruction",
    "Use provided structured job description data as source of truth. Analyze resume only for education evidence, role chronology, relevance, alignment to listed responsibilities, school overlap, and flags. Do not infer job requirements beyond provided JD data."
)
    st.session_state["reviewer_instruction"] = st.text_area("HR reviewer instruction", value=st.session_state["reviewer_instruction"], height=300)
    if get_api_key():
        st.success("AI-VERDE API key loaded from secrets/environment.")
    else:
        st.error("No AI-VERDE API key found yet.")

st.subheader("1. Upload resume")
st.session_state["resume_file"] = st.file_uploader("Candidate/Incumbent resume (PDF, DOCX, or TXT)", type=["pdf", "docx", "txt"], accept_multiple_files=False, key="resume_uploader")

st.subheader("2. Upload job description")
st.session_state["job_description_file"] = st.file_uploader("University job description (PDF, DOCX, or TXT)", type=["pdf", "docx", "txt"], accept_multiple_files=False, key="jd_uploader")

if st.button("▶ Run HR review", type="primary"):
    try:
        run_analysis()
    except requests.HTTPError as e:
        details = e.response.text if e.response is not None else str(e)
        st.error(f"API request failed: {details}")
    except Exception as e:
        st.error(f"Something went wrong: {e}")

if st.session_state.get("last_json_text"):
    if DEBUG:
        with st.expander("Debug: raw model JSON response"):
            st.text(st.session_state["last_json_text"])

if "last_structured_parse" in st.session_state:
    if DEBUG:
        with st.expander("Debug: parsed structured data"):
            st.write(st.session_state["last_structured_parse"])

if st.session_state.get("final_output"):
    st.markdown(st.session_state["final_output"])
    st.download_button(label="⬇ Download results as .txt", data=st.session_state["final_output"], file_name="HR_MQ_Review.txt", mime="text/plain")

    if st.session_state.get("structured_data"):
        if DEBUG:
            with st.expander("Debug: structured_data"):
                st.json(st.session_state["structured_data"])

    st.subheader("Additional Details")
    additional_followup = st.text_input("Request details on a section, role, calculation, or assumption.", placeholder="e.g., Give details on Section 2 only", key="additional_followup")
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("Additional Details"):
            try:
                run_analysis(additional_followup)
            except requests.HTTPError as e:
                details = e.response.text if e.response is not None else str(e)
                st.error(f"API request failed: {details}")
            except Exception as e:
                st.error(f"Something went wrong: {e}")

