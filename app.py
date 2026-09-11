import os
import hmac
import sqlite3
import hashlib
import secrets
import re
from datetime import datetime

import gradio as gr
import spacy
from pypdf import PdfReader
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


# ============================================================
# CV STORAGE
# ============================================================
# ============================================================


# ============================================================
# CVFIX-SA OWNER AUTHENTICATION
# ============================================================

# Owner credentials are intentionally NOT stored in app.py.
# They come from Render environment variables.
#
# Required environment variables:
# CVFIX_OWNER_USERNAME
# CVFIX_OWNER_PASSWORD

def cvfix_owner_auth(username, password):

    owner_username = os.environ.get(
        "CVFIX_OWNER_USERNAME",
        ""
    )

    owner_password = os.environ.get(
        "CVFIX_OWNER_PASSWORD",
        ""
    )

    if not owner_username or not owner_password:
        return False

    username_match = hmac.compare_digest(
        str(username),
        str(owner_username)
    )

    password_match = hmac.compare_digest(
        str(password),
        str(owner_password)
    )

    return username_match and password_match


# CVFIX-SA SMART ENGINE
# ============================================================

nlp = spacy.load("en_core_web_sm")


# ------------------------------------------------------------
# Generic job-advertisement words
# ------------------------------------------------------------

GENERIC_JOB_WORDS = {
    "candidate", "company", "employee", "role",
    "work", "working", "worked", "job",
    "responsibility", "responsibilities",
    "requirement", "requirements", "position",
    "successful", "successfully", "join",
    "provide", "providing", "previous",
    "right", "look", "looking", "opportunity",
    "experience", "advantage", "preferred",
    "prefer", "include", "including",
    "ensure", "maintain", "support",
    "supporting", "ability", "able",
    "good", "strong", "excellent",
    "best", "new", "multiple", "various",
    "related", "relevant", "need",
    "needed", "required", "develop",
    "developing", "member"
}


# ------------------------------------------------------------
# Related concepts
# ------------------------------------------------------------

TERM_GROUPS = {

    "communication": {
        "communicate",
        "communication",
        "communicative",
        "speaking",
        "verbal"
    },

    "customer_service": {
        "customer",
        "customers",
        "client",
        "clients",
        "service",
        "serving",
        "enquiry",
        "enquiries"
    },

    "teamwork": {
        "teamwork",
        "team",
        "collaborate",
        "collaboration",
        "colleague",
        "colleagues"
    },

    "problem_solving": {
        "problem",
        "problems",
        "solving",
        "solve",
        "critical",
        "thinking"
    },

    "computer_literacy": {
        "computer",
        "computers",
        "literacy",
        "technology",
        "software"
    },

    "excel": {
        "excel",
        "spreadsheet",
        "spreadsheets"
    },

    "microsoft_office": {
        "microsoft",
        "office",
        "word",
        "excel"
    },

    "time_management": {
        "time",
        "management",
        "manage",
        "managing",
        "punctual"
    },

    "adaptability": {
        "adaptable",
        "adaptability",
        "flexible",
        "flexibility"
    },

    "retail_sales": {
        "retail",
        "sales",
        "sale",
        "selling",
        "sell",
        "salesperson"
    },

    "professionalism": {
        "professional",
        "professionally",
        "professionalism"
    },

    "pressure": {
        "pressure",
        "stress",
        "deadline",
        "deadlines"
    },

    "learning": {
        "learn",
        "learning",
        "learner",
        "training",
        "train"
    },

    "task_management": {
        "task",
        "tasks",
        "manage",
        "managing",
        "multiple"
    }
}


def get_lemmas(text):

    doc = nlp(text.lower())

    lemmas = set()

    for token in doc:

        if token.is_alpha and not token.is_stop:

            lemma = token.lemma_.lower().strip()

            if len(lemma) >= 3:

                lemmas.add(lemma)

    return lemmas


def get_meaningful_terms(text):

    doc = nlp(text.lower())

    terms = set()

    for token in doc:

        if not token.is_alpha:
            continue

        if token.is_stop:
            continue

        lemma = token.lemma_.lower().strip()

        if len(lemma) < 3:
            continue

        if lemma in GENERIC_JOB_WORDS:
            continue

        terms.add(lemma)

    return terms


def find_term_groups(terms):

    groups = set()

    for group_name, group_terms in TERM_GROUPS.items():

        if terms.intersection(group_terms):

            groups.add(group_name)

    return groups


def smart_cv_job_match(cv_text, job_description):

    if not cv_text.strip():

        return {
            "score": 0,
            "direct_matches": [],
            "related_matches": [],
            "missing": []
        }

    if not job_description.strip():

        return {
            "score": 0,
            "direct_matches": [],
            "related_matches": [],
            "missing": []
        }

    cv_terms = get_meaningful_terms(cv_text)
    job_terms = get_meaningful_terms(job_description)

    cv_groups = find_term_groups(cv_terms)
    job_groups = find_term_groups(job_terms)

    direct_matches = set()
    related_matches = set()
    missing = set()

    # --------------------------------------------------------
    # 1. Direct matches
    # --------------------------------------------------------

    for term in job_terms:

        if term in cv_terms:

            direct_matches.add(term)

        else:

            missing.add(term)

    # --------------------------------------------------------
    # 2. Related / transferable matches
    # --------------------------------------------------------

    for group in job_groups:

        if group in cv_groups:

            group_terms = TERM_GROUPS[group]

            job_group_terms = job_terms.intersection(group_terms)

            for term in job_group_terms:

                if term not in direct_matches:

                    related_matches.add(term)

                if term in missing:

                    missing.remove(term)

    # --------------------------------------------------------
    # 3. Identify genuinely missing concepts
    # --------------------------------------------------------

    missing_groups = []

    for group in job_groups:

        if group not in cv_groups:

            missing_groups.append(group)

    # --------------------------------------------------------
    # 4. Calculate balanced score
    # --------------------------------------------------------

    total_requirements = (
        len(direct_matches)
        + len(related_matches)
        + len(missing)
    )

    if total_requirements == 0:

        score = 0

    else:

        score = round(
            (
                len(direct_matches)
                + (len(related_matches) * 0.75)
            )
            / total_requirements
            * 100
        )

    return {
        "score": score,
        "direct_matches": sorted(direct_matches),
        "related_matches": sorted(related_matches),
        "missing": sorted(missing),
        "missing_groups": sorted(missing_groups)
    }


def calculate_smart_scores(cv_text, job_description):

    result = smart_cv_job_match(
        cv_text,
        job_description
    )

    return {
        "match_score": result["score"],
        "direct_matches": result["direct_matches"],
        "related_matches": result["related_matches"],
        "missing_keywords": result["missing"],
        "missing_groups": result["missing_groups"],
        "matched_count": (
            len(result["direct_matches"])
            + len(result["related_matches"])
        ),
        "missing_count": len(result["missing"])
    }


def smart_cv_improvement(cv_text, job_description):

    result = smart_cv_job_match(
        cv_text,
        job_description
    )

    suggestions = []

    for group in result["missing_groups"]:

        suggestions.append(
            f"Your CV may benefit from adding truthful "
            f"information related to '{group.replace('_', ' ')}'."
        )

    for keyword in result["missing"]:

        suggestions.append(
            f"Consider addressing '{keyword}' if it genuinely "
            f"applies to your experience."
        )

    return {
        "match_score": result["score"],
        "direct_matches": result["direct_matches"],
        "related_matches": result["related_matches"],
        "missing_keywords": result["missing"],
        "suggestions": suggestions
    }
website_cv_text = ""

improved_cv_text = ""
current_cv_text = ""

# Store the original uploaded CV file for format preservation
original_cv_file_path = None


# ============================================================
# CV UPLOAD
# ============================================================

def upload_cv(file):

    global original_cv_file_path

    global website_cv_text

    if file is not None:
        original_cv_file_path = file

    if file is None:
        return "⚠️ Please upload your CV."

    filename = file.name

    try:

        if filename.lower().endswith(".pdf"):

            reader = PdfReader(filename)

            text = ""

            for page in reader.pages:

                page_text = page.extract_text()

                if page_text:
                    text += page_text + "\n"

        elif filename.lower().endswith(".docx"):

            document = Document(filename)

            text = ""

            for paragraph in document.paragraphs:

                text += paragraph.text + "\n"

        else:

            return "⚠️ Please upload a PDF or DOCX file."

        if not text.strip():

            return "⚠️ No readable text was found."

        website_cv_text = text

        return (
            "## ✅ CV Uploaded Successfully\n\n"
            f"**File:** {filename.split('/')[-1]}\n\n"
            f"**Characters extracted:** {len(text)}"
        )

    except Exception as error:

        return (
            "⚠️ **Could not read the CV**\n\n"
            f"`{error}`"
        )


# ============================================================
# SMART CV ANALYSIS
# ============================================================

def analyse_cv(job_description):

    if not website_cv_text.strip():

        return (
            "⚠️ Please upload your CV first.",
            "",
            "",
            "",
            ""
        )

    if not job_description.strip():

        return (
            "⚠️ Please enter a job description.",
            "",
            "",
            "",
            ""
        )

    # --------------------------------------------------------
    # SMART JOB MATCH
    # --------------------------------------------------------

    match_result = smart_cv_job_match(
        website_cv_text,
        job_description
    )

    match_score = match_result["score"]

    direct_matches = match_result["direct_matches"]
    related_matches = match_result["related_matches"]
    missing = match_result["missing"]

    # --------------------------------------------------------
    # CV STRUCTURE
    # --------------------------------------------------------

    word_count = len(
        website_cv_text.split()
    )

    if word_count >= 300:

        structure_score = 90

    elif word_count >= 150:

        structure_score = 75

    elif word_count >= 75:

        structure_score = 60

    else:

        structure_score = 40

    # --------------------------------------------------------
    # SKILLS
    # --------------------------------------------------------

    cv_lower = website_cv_text.lower()

    skill_groups = {

        "Communication": {
            "communication",
            "communicate",
            "speaking"
        },

        "Customer Service": {
            "customer",
            "customers",
            "client",
            "service"
        },

        "Teamwork": {
            "teamwork",
            "team",
            "colleague"
        },

        "Problem Solving": {
            "problem solving",
            "problem-solving",
            "critical thinking"
        },

        "Computer Literacy": {
            "computer",
            "computer literacy"
        },

        "Microsoft Excel": {
            "excel"
        },

        "Time Management": {
            "time management",
            "punctual"
        },

        "Adaptability": {
            "adaptable",
            "adaptability"
        },

        "Leadership": {
            "leadership",
            "leader"
        },

        "Project Management": {
            "project management"
        },

        "Data Analysis": {
            "data analysis"
        },

        "Python": {
            "python"
        },

        "SQL": {
            "sql"
        },

        "Power BI": {
            "power bi"
        }
    }

    detected_skills = []

    for skill_name, terms in skill_groups.items():

        if any(term in cv_lower for term in terms):

            detected_skills.append(skill_name)

    skill_score = min(
        100,
        len(detected_skills) * 10
    )

    # --------------------------------------------------------
    # OVERALL SCORE
    # --------------------------------------------------------

    overall = int(
        structure_score * 0.30
        + match_score * 0.50
        + skill_score * 0.20
    )

    # --------------------------------------------------------
    # DISPLAY MATCHES
    # --------------------------------------------------------

    direct_text = (
        "\n".join(
            f"✓ {word}"
            for word in direct_matches[:30]
        )
        if direct_matches
        else "None detected."
    )

    related_text = (
        "\n".join(
            f"🔵 {word}"
            for word in related_matches[:30]
        )
        if related_matches
        else "None detected."
    )

    requirement_labels = {
        "retail_sales": "Retail / sales experience",
        "customer_service": "Customer service",
        "communication": "Professional communication",
        "teamwork": "Teamwork and collaboration",
        "problem_solving": "Problem solving and critical thinking",
        "computer_literacy": "Computer literacy",
        "excel": "Microsoft Excel",
        "microsoft_office": "Microsoft Office",
        "time_management": "Time management",
        "adaptability": "Adaptability",
        "professionalism": "Professionalism",
        "pressure": "Working under pressure",
        "learning": "Training and willingness to learn",
        "task_management": "Task management"
    }

    missing_text = (
        "\n".join(
            f"⚠ {requirement_labels.get(group, group.replace('_', ' ').title())}"
            for group in match_result["missing_groups"]
        )
        if match_result["missing_groups"]
        else "No major gaps detected."
    )

    skills_text = (
        ", ".join(detected_skills)
        if detected_skills
        else "None detected."
    )

    # --------------------------------------------------------
    # REPORT
    # --------------------------------------------------------

    report = f"""
# 📊 CVFix-SA Smart Analysis

## Overall Score: {overall}/100

| Category | Score |
|---|---:|
| 📄 CV Structure | {structure_score}/100 |
| 💼 Smart Job Match | {match_score}/100 |
| 🛠 Skills | {skill_score}/100 |

---

## 🟢 Direct Matches

These are skills or requirements directly supported
by wording in your CV.

{direct_text}

---

## 🔵 Related / Transferable Matches

These are relevant concepts where your CV shows
related abilities, even when the exact job wording
is different.

{related_text}

---

## 🟠 Potential Gaps

These are requirements that were not clearly detected
in your CV.

{missing_text}

---

## 🛠 Detected Skills

{skills_text}

---

## 💡 Recommendation

CVFix-SA compares the job description with your CV
using meaningful terms and related skill concepts.

A potential gap does not automatically mean you are
unqualified. It may simply mean that the relevant
experience or skill is not clearly stated in your CV.

Only add information that is truthful and supported
by your actual experience.
"""

    return (
        report,
        f"{overall}/100",
        f"{structure_score}/100",
        f"{match_score}/100",
        f"{skill_score}/100"
    )

# ============================================================
# CV IMPROVEMENT
# ============================================================
def improve_cv(job_description):

    global improved_cv_text

    if not website_cv_text.strip():

        return "⚠️ Please upload your CV first."

    if not job_description.strip():

        return "⚠️ Please enter the job description first."

    replacements = {
        "responsible for": "managed and delivered",
        "worked on": "contributed to",
        "helped with": "supported",
        "helped": "supported",
        "used": "utilised",
        "worked with": "collaborated with",
        "involved in": "contributed to"
    }

    improved = website_cv_text
    recommendations = []

    for weak, strong in replacements.items():

        pattern = re.compile(
            re.escape(weak),
            re.IGNORECASE
        )

        if pattern.search(improved):

            recommendations.append(
                f"Consider replacing '{weak}' "
                f"with '{strong}'."
            )

            improved = pattern.sub(
                strong,
                improved
            )

    # Save the improved CV so the download function
    # can use the same improved wording.
    improved_cv_text = improved

    if not recommendations:

        recommendations.append(
            "Your CV wording has a reasonable foundation. "
            "Consider adding measurable achievements."
        )

    result = "# ✨ CVFix-SA Improvement Report\n\n"

    result += "## 💡 Recommendations\n\n"

    for item in recommendations:

        result += "- " + item + "\n"

    result += "\n---\n\n"

    result += "## 📄 Improved CV\n\n"

    result += improved

    result += (
        "\n\n---\n\n"
        "⚠️ **Important:** Verify all information before "
        "using the improved CV."
    )

    return result




# ============================================================
# CREATE WORD DOCUMENT
# ============================================================



# ============================================================
# CVFix-SA — STRUCTURED PROFESSIONAL CV GENERATOR
# ============================================================

def _cvfix_extract_sections(text):
    """
    Convert improved CV text into recognisable CV sections.
    """
    sections = {
        "about": [],
        "education": [],
        "experience": [],
        "skills": [],
        "other": []
    }

    current = "other"

    heading_map = {
        "about": "about",
        "about me": "about",
        "profile": "about",
        "professional profile": "about",
        "professional summary": "about",
        "summary": "about",
        "objective": "about",
        "career objective": "about",

        "education": "education",
        "qualifications": "education",
        "academic qualifications": "education",

        "experience": "experience",
        "work experience": "experience",
        "employment": "experience",
        "employment history": "experience",
        "work history": "experience",

        "skills": "skills",
        "technical skills": "skills",
        "key skills": "skills",
        "core skills": "skills",

        "references": "other",
        "certifications": "other",
        "achievements": "other",
        "projects": "other",
        "languages": "other"
    }

    for raw in text.splitlines():

        line = raw.strip()

        if not line:
            continue

        clean = line.lower().strip(": ").strip()

        if clean in heading_map:
            current = heading_map[clean]
            continue

        sections[current].append(line)

    return sections


def _cvfix_first_lines(text, count=5):
    """
    Extract likely header/contact information.
    """
    result = []

    for line in text.splitlines():

        line = line.strip()

        if not line:
            continue

        result.append(line)

        if len(result) >= count:
            break

    return result


def _cvfix_section_heading(document, title):
    """
    Create the light-grey CV section bar.
    """

    paragraph = document.add_paragraph()

    paragraph.paragraph_format.space_before = Pt(7)
    paragraph.paragraph_format.space_after = Pt(4)

    # Shading
    pPr = paragraph._p.get_or_add_pPr()

    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), "EDEDED")

    pPr.append(shading)

    paragraph.paragraph_format.left_indent = Inches(0)
    paragraph.paragraph_format.right_indent = Inches(0)

    run = paragraph.add_run("  " + title.upper())

    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(11)

    return paragraph


def _cvfix_add_bullet(document, text):
    """
    Add a clean bullet item.
    """

    paragraph = document.add_paragraph(style="List Bullet")

    paragraph.paragraph_format.space_after = Pt(2)
    paragraph.paragraph_format.left_indent = Inches(0.18)

    run = paragraph.add_run(text)

    run.font.name = "Arial"
    run.font.size = Pt(9.5)

    return paragraph


def _cvfix_add_normal(document, text):
    """
    Add normal CV body text.
    """

    paragraph = document.add_paragraph()

    paragraph.paragraph_format.space_after = Pt(3)
    paragraph.paragraph_format.line_spacing = 1.0

    run = paragraph.add_run(text)

    run.font.name = "Arial"
    run.font.size = Pt(9.5)

    return paragraph



def create_cvfix_professional_cv():
    """
    CVFix-SA Smart Professional CV Generator.

    Creates a structured professional CV instead of simply
    placing extracted CV text line-by-line into a document.

    Structure:
        Header
        Contact row
        About Me
        Education
        Work Experience
        Skills
        Certifications
        Achievements
        Languages
        Projects
        References

    The parser intelligently recognises:
        - names
        - professional titles
        - phone numbers
        - email addresses
        - locations
        - section headings
        - dates
        - date ranges
        - bullets
        - skills
    """

    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    from docx.enum.section import WD_SECTION
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    import os

    # ========================================================
    # SOURCE CV
    # ========================================================

    raw_text = globals().get("website_cv_text", "")

    if not raw_text or not str(raw_text).strip():

        raw_text = globals().get("current_cv_text", "")

    if not raw_text or not str(raw_text).strip():

        return None

    raw_text = str(raw_text).replace("\r\n", "\n").replace("\r", "\n")

    lines = [
        re.sub(r"[ \t]+", " ", line).strip()
        for line in raw_text.split("\n")
    ]

    lines = [line for line in lines if line]

    # ========================================================
    # SMART REGEX PATTERNS
    # ========================================================

    email_pattern = re.compile(
        r"\b[A-Za-z0-9._%+-]+@"
        r"[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b"
    )

    phone_pattern = re.compile(
        r"(?:(?:\+27|0027)\s*"
        r"(?:\(?\d{2}\)?)[\s.-]*"
        r"\d{3}[\s.-]*\d{4})"
        r"|(?:0\d{2})[\s.-]*\d{3}[\s.-]*\d{4}"
    )

    date_pattern = re.compile(
        r"\b(?:"
        r"(?:19|20)\d{2}"
        r"|"
        r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|"
        r"Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|"
        r"Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|"
        r"Dec(?:ember)?)"
        r"(?:\s+\d{4})?"
        r")\b",
        re.I
    )

    date_range_pattern = re.compile(
        r"("
        r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|"
        r"Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|"
        r"Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|"
        r"Dec(?:ember)?)[\s-]+\d{4}"
        r"|"
        r"\d{1,2}[/-]\d{4}"
        r"|"
        r"(?:19|20)\d{2}"
        r")"
        r"\s*(?:-|–|—|to|until)\s*"
        r"("
        r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|"
        r"Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|"
        r"Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|"
        r"Dec(?:ember)?)[\s-]+\d{4}"
        r"|"
        r"\d{1,2}[/-]\d{4}"
        r"|"
        r"(?:19|20)\d{2}"
        r"|Present|Current|Now"
        r")",
        re.I
    )

    # ========================================================
    # SECTION RECOGNITION
    # ========================================================

    section_aliases = {
        "ABOUT ME": {
            "about me",
            "about",
            "profile",
            "personal profile",
            "professional profile",
            "professional summary",
            "summary",
            "career summary",
            "objective",
            "career objective",
            "personal statement",
            "professional statement",
        },

        "EDUCATION": {
            "education",
            "educational background",
            "academic background",
            "qualifications",
            "academic qualifications",
        },

        "WORK EXPERIENCE": {
            "work experience",
            "experience",
            "employment history",
            "work history",
            "professional experience",
            "employment",
        },

        "SKILLS": {
            "skills",
            "key skills",
            "core skills",
            "technical skills",
            "professional skills",
            "competencies",
            "core competencies",
        },

        "CERTIFICATIONS": {
            "certifications",
            "certificates",
            "certification",
            "professional certifications",
            "licenses",
            "licences",
        },

        "ACHIEVEMENTS": {
            "achievements",
            "accomplishments",
            "awards",
            "honours",
            "honors",
        },

        "PROJECTS": {
            "projects",
            "academic projects",
            "personal projects",
            "key projects",
        },

        "LANGUAGES": {
            "languages",
            "language skills",
        },

        "REFERENCES": {
            "references",
            "referees",
            "professional references",
        },

        "CONTACT": {
            "contact",
            "contact details",
            "personal details",
        },
    }

    def normalise_heading(text):
        cleaned = re.sub(r"[:\-–—]+$", "", text.strip())
        cleaned = re.sub(r"\s+", " ", cleaned)
        return cleaned.lower()

    def detect_section(text):
        normalized = normalise_heading(text)

        # Remove common numbering such as:
        # 1. EDUCATION
        # 02 - EXPERIENCE
        normalized = re.sub(
            r"^\d+\s*[\.\):-]\s*",
            "",
            normalized
        )

        for canonical, aliases in section_aliases.items():

            if normalized in aliases:
                return canonical

        return None

    # ========================================================
    # SMART DATE FUNCTIONS
    # ========================================================

    def clean_date_text(text):
        """
        Convert common date representations into a consistent
        professional form without changing the actual dates.
        """

        text = text.strip()

        text = re.sub(
            r"\bto\b",
            "–",
            text,
            flags=re.I
        )

        text = re.sub(
            r"\s*[-–—]\s*",
            " – ",
            text
        )

        text = re.sub(
            r"\s+",
            " ",
            text
        )

        text = re.sub(
            r"\bCurrent\b",
            "Present",
            text,
            flags=re.I
        )

        text = re.sub(
            r"\bNow\b",
            "Present",
            text,
            flags=re.I
        )

        # Convert numeric month/year:
        # 01/2024 -> Jan 2024
        month_names = {
            "01": "Jan",
            "02": "Feb",
            "03": "Mar",
            "04": "Apr",
            "05": "May",
            "06": "Jun",
            "07": "Jul",
            "08": "Aug",
            "09": "Sep",
            "10": "Oct",
            "11": "Nov",
            "12": "Dec",
        }

        def convert_month(match):
            month = match.group(1)
            year = match.group(2)

            return (
                f"{month_names.get(month, month)} "
                f"{year}"
            )

        text = re.sub(
            r"\b(0[1-9]|1[0-2])[/.-](19\d{2}|20\d{2})\b",
            convert_month,
            text
        )

        return text

    def extract_date(text):

        if not text:
            return None

        match = date_range_pattern.search(text)

        if match:
            return clean_date_text(match.group(0))

        # Handle "2024 to Present" where the range regex
        # may not catch an unusual separator.
        year_range = re.search(
            r"\b((?:19|20)\d{2})\s*"
            r"(?:-|–|—|to)\s*"
            r"(Present|Current|Now|(?:19|20)\d{2})\b",
            text,
            re.I
        )

        if year_range:
            return clean_date_text(year_range.group(0))

        return None

    def remove_date_from_line(text, date_text):

        if not date_text:
            return text.strip()

        result = text.replace(date_text, "")

        result = re.sub(
            r"\s{2,}",
            " ",
            result
        )

        result = re.sub(
            r"[\s,|]+$",
            "",
            result
        )

        return result.strip()

    # ========================================================
    # SMART CONTACT DETECTION
    # ========================================================

    email = None
    phone = None
    # -----------------------------------------------------------------------
    # SMART LOCATION DETECTION
    # -----------------------------------------------------------------------
    # Never treat an education institution, employer, project title or
    # qualification as a person's location.
    
    location = ""
    
    location_candidates = []
    
    # Common South African / international location patterns.
    location_patterns = [
        r'\bBellville South\s*,?\s*Cape Town\b',
        r'\bBellville\s*,?\s*Cape Town\b',
        r'\bCape Town\s*,?\s*Western Cape\b',
        r'\bCape Town\b',
        r'\bJohannesburg\b',
        r'\bPretoria\b',
        r'\bDurban\b',
        r'\bGqeberha\b',
        r'\bPort Elizabeth\b',
        r'\bStellenbosch\b',
        r'\bSomerset West\b',
        r'\bKhayelitsha\b',
        r'\bMitchells Plain\b',
        r'\bGoodwood\b',
        r'\bParow\b',
        r'\bBellville\b',
        r'\bSouth Africa\b'
    ]
    
    # Search the original CV text first.
    location_search_text = current_cv_text or website_cv_text or ""
    
    for pattern in location_patterns:
        match = re.search(pattern, location_search_text, re.IGNORECASE)
        if match:
            candidate = match.group(0).strip()
    
            # Reject country-only matches when a more useful location exists.
            if candidate.lower() == "south africa":
                continue
    
            location_candidates.append(candidate)
    
    if location_candidates:
        # Prefer the most specific location.
        location = max(location_candidates, key=len)
    
    # -----------------------------------------------------------------------
    # SAFETY FILTER
    # -----------------------------------------------------------------------
    # These are institutions/organisations and must NEVER become a location.
    
    invalid_location_terms = [
        "cape peninsula university of technology",
        "university",
        "college",
        "school",
        "academy",
        "company",
        "corporation",
        "pty ltd",
        "project",
        "department",
        "faculty"
    ]
    
    if any(
        bad_term in location.lower()
        for bad_term in invalid_location_terms
    ):
        location = ""
    
    # If location was not detected, inspect individual CV lines.
    if not location:
        for raw_line in location_search_text.splitlines():
            line = raw_line.strip()
    
            if not line:
                continue
    
            low = line.lower()
    
            # Reject obvious non-location lines.
            if any(term in low for term in invalid_location_terms):
                continue
    
            # Reject email / phone / headings.
            if "@" in line:
                continue
    
            if re.search(r'\+?\d[\d\s().-]{7,}', line):
                continue
    
            if low in {
                "about me",
                "education",
                "work experience",
                "experience",
                "skills",
                "certifications",
                "achievements",
                "references",
                "projects",
                "languages"
            }:
                continue
    
            # Location-like line.
            if re.search(
                r'\b(cape town|bellville|south africa|johannesburg|pretoria|durban|'
                r'gqeberha|port elizabeth|stellenbosch|somerset west|khayelitsha|'
                r'mitchells plain|goodwood|parow)\b',
                line,
                re.IGNORECASE
            ):
                location = line
                break

    email_index = None
    phone_index = None

    for index, line in enumerate(lines):

        if email is None:

            match = email_pattern.search(line)

            if match:
                email = match.group(0)
                email_index = index

        if phone is None:

            match = phone_pattern.search(line)

            if match:
                phone = match.group(0)
                phone_index = index

    # Look for explicit location labels first.
    for line in lines[:15]:

        lower = line.lower()

        if (
            "location:" in lower
            or "address:" in lower
            or "based in:" in lower
            or "residence:" in lower
        ):

            parts = re.split(
                r":",
                line,
                maxsplit=1
            )

            if len(parts) == 2:

                possible = parts[1].strip()

                if possible:
                    location = possible
                    break

    # Location heuristics if there was no explicit label.
    if location is None:

        location_candidates = []

        for line in lines[:15]:

            lower = line.lower()

            if (
                "location" in lower
                or "address" in lower
                or "cape town" in lower
                or "bellville" in lower
                or "south africa" in lower
                or "cape" in lower
            ):

                if (
                    email_pattern.search(line)
                    or phone_pattern.search(line)
                ):
                    continue

                if detect_section(line):
                    continue

                location_candidates.append(line)

        if location_candidates:
            location = location_candidates[-1]

    # ========================================================
    # SMART HEADER DETECTION
    # ========================================================

    contact_indices = set()

    for index, line in enumerate(lines[:15]):

        if email_pattern.search(line):
            contact_indices.add(index)

        if phone_pattern.search(line):
            contact_indices.add(index)

        lower = line.lower()

        if (
            "location:" in lower
            or "address:" in lower
            or "based in:" in lower
        ):
            contact_indices.add(index)

    # Name is normally the first meaningful line that is
    # not contact information or a section heading.
    name = None

    for index, line in enumerate(lines[:10]):

        if index in contact_indices:
            continue

        if detect_section(line):
            continue

        if email_pattern.search(line):
            continue

        if phone_pattern.search(line):
            continue

        # Don't accept obvious date-only lines.
        if extract_date(line) and len(line) < 35:
            continue

        name = line
        break

    # Professional title is normally the next meaningful
    # non-contact line.
    professional_title = None

    if name:

        name_index = lines.index(name)

        for line in lines[name_index + 1:name_index + 7]:

            if line in contact_indices:
                continue

            if email_pattern.search(line):
                continue

            if phone_pattern.search(line):
                continue

            if detect_section(line):
                continue

            if extract_date(line) and len(line) < 35:
                continue

            professional_title = line
            break

    # ========================================================
    # DETECT CONTACT-ONLY LINES
    # ========================================================

    def is_contact_line(line):

        if email_pattern.search(line):
            return True

        if phone_pattern.search(line):
            return True

        lower = line.lower()

        if (
            lower.startswith("location:")
            or lower.startswith("address:")
            or lower.startswith("based in:")
        ):
            return True

        # A short line containing only contact fields.
        if (
            (
                "location" in lower
                or "address" in lower
            )
            and len(line) < 100
        ):
            return True

        return False

    # ========================================================
    # PARSE CV INTO SECTIONS
    # ========================================================

    sections = {}
    current_section = None

    header_indices = set()

    if name and name in lines:
        header_indices.add(lines.index(name))

    if professional_title and professional_title in lines:
        header_indices.add(lines.index(professional_title))

    # Everything immediately before the first recognised
    # section is treated as header/contact information.
    first_section_index = None

    for i, line in enumerate(lines):

        if detect_section(line):
            first_section_index = i
            break

    if first_section_index is None:
        first_section_index = len(lines)

    for i in range(first_section_index):

        header_indices.add(i)

    for line in lines[first_section_index:]:

        detected = detect_section(line)

        if detected:

            current_section = detected

            if current_section not in sections:
                sections[current_section] = []

            continue

        if current_section is not None:

            # Remove accidental duplicate header/contact lines
            # from later sections.
            if line == name:
                continue

            if (
                professional_title
                and line == professional_title
            ):
                continue

            if email and email in line:
                continue

            if phone and phone in line:
                continue

            if location and line == location:
                continue

            sections[current_section].append(line)

    # ========================================================
    # RECOVER UNLABELLED INFORMATION
    # ========================================================

    # Some CVs don't use section headings consistently.
    # Don't lose their content.

    if not sections:

        sections["ABOUT ME"] = [
            line
            for line in lines
            if line not in header_indices
        ]

    # ========================================================
    # WORD DOCUMENT
    # ========================================================

    document = Document()

    section = document.sections[0]

    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.65)
    section.right_margin = Inches(0.65)

    # ========================================================
    # GLOBAL NORMAL STYLE
    # ========================================================

    normal_style = document.styles["Normal"]

    normal_style.font.name = "Arial"
    normal_style.font.size = Pt(10)

    # ========================================================
    # HELPER: CELL SHADING
    # ========================================================

    def shade_cell(cell, fill="E7E7E7"):

        tcPr = cell._tc.get_or_add_tcPr()

        shd = OxmlElement("w:shd")

        shd.set(
            qn("w:fill"),
            fill
        )

        tcPr.append(shd)

    # ========================================================
    # HELPER: SECTION HEADING
    # ========================================================

    def add_section_heading(title):

        table = document.add_table(
            rows=1,
            cols=1
        )

        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        table.autofit = True

        cell = table.cell(0, 0)

        shade_cell(cell, "E6E6E6")

        paragraph = cell.paragraphs[0]

        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        paragraph.paragraph_format.space_before = Pt(3)
        paragraph.paragraph_format.space_after = Pt(3)

        run = paragraph.add_run(title.upper())

        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(11)

        return table

    # ========================================================
    # HELPER: NORMAL PARAGRAPH
    # ========================================================

    def add_body_text(text):

        if not text:
            return

        paragraph = document.add_paragraph()

        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

        paragraph.paragraph_format.space_before = Pt(1)
        paragraph.paragraph_format.space_after = Pt(3)
        paragraph.paragraph_format.line_spacing = 1.05

        run = paragraph.add_run(text)

        run.font.name = "Arial"
        run.font.size = Pt(9.5)

    # ========================================================
    # HELPER: BULLET
    # ========================================================

    def add_bullet(text):

        paragraph = document.add_paragraph(
            style="List Bullet"
        )

        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(2)

        run = paragraph.add_run(text)

        run.font.name = "Arial"
        run.font.size = Pt(9.5)

    # ========================================================
    # HEADER
    # ========================================================

    if name:

        paragraph = document.add_paragraph()

        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        paragraph.paragraph_format.space_after = Pt(1)

        run = paragraph.add_run(name)

        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(22)

    if professional_title:

        paragraph = document.add_paragraph()

        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        paragraph.paragraph_format.space_after = Pt(5)

        run = paragraph.add_run(
            professional_title
        )

        run.font.name = "Arial"
        run.font.size = Pt(11)
        run.italic = True

    # ========================================================
    # CONTACT ROW
    # ========================================================

    contact_items = []

    if phone:
        phone_clean = re.sub(
            r"\s+",
            " ",
            phone
        ).strip()

        contact_items.append(
            f"☎ {phone_clean}"
        )

    if email:
        contact_items.append(
            f"✉ {email}"
        )

    if location:
        contact_items.append(
            f"📍 {location}"
        )

    if contact_items:

        paragraph = document.add_paragraph()

        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        paragraph.paragraph_format.space_after = Pt(8)

        run = paragraph.add_run(
            "   |   ".join(contact_items)
        )

        run.font.name = "Arial"
        run.font.size = Pt(9.5)

    # ========================================================
    # ABOUT ME
    # ========================================================

    about_lines = sections.get("ABOUT ME", [])

    if about_lines:

        add_section_heading("ABOUT ME")

        for line in about_lines:

            if line.strip():

                if line.startswith(
                    ("•", "-", "*")
                ):

                    add_bullet(
                        re.sub(
                            r"^[•\-*]\s*",
                            "",
                            line
                        )
                    )

                else:

                    add_body_text(line)

    # ========================================================
    # EDUCATION
    # ========================================================

    education_lines = sections.get(
        "EDUCATION",
        []
    )

    if education_lines:

        add_section_heading("EDUCATION")

        i = 0

        while i < len(education_lines):

            line = education_lines[i]

            date = extract_date(line)

            if date:

                # Standalone date belongs to previous entry.
                if i > 0:

                    previous = education_lines[i - 1]

                    if date == previous:

                        i += 1
                        continue

                i += 1
                continue

            # Look ahead for a date.
            date = None
            date_index = None

            for j in range(
                i,
                min(i + 4, len(education_lines))
            ):

                possible_date = extract_date(
                    education_lines[j]
                )

                if possible_date:

                    date = possible_date
                    date_index = j
                    break

            # Institution/qualification line
            table = document.add_table(
                rows=1,
                cols=2
            )

            table.alignment = WD_TABLE_ALIGNMENT.CENTER

            left = table.cell(0, 0)
            right = table.cell(0, 1)

            left.vertical_alignment = (
                WD_CELL_VERTICAL_ALIGNMENT.CENTER
            )

            right.vertical_alignment = (
                WD_CELL_VERTICAL_ALIGNMENT.CENTER
            )

            p_left = left.paragraphs[0]
            p_left.alignment = WD_ALIGN_PARAGRAPH.LEFT

            r_left = p_left.add_run(line)

            r_left.bold = True
            r_left.font.name = "Arial"
            r_left.font.size = Pt(10)

            if date:

                p_right = right.paragraphs[0]
                p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT

                r_right = p_right.add_run(
                    date
                )

                r_right.bold = True
                r_right.font.name = "Arial"
                r_right.font.size = Pt(9)

            i += 1

            # Qualification underneath
            if (
                i < len(education_lines)
                and (
                    date_index is None
                    or i != date_index
                )
            ):

                next_line = education_lines[i]

                if (
                    not extract_date(next_line)
                    and not detect_section(next_line)
                ):

                    add_body_text(next_line)

                    i += 1

            # Description until next likely entry.
            while i < len(education_lines):

                next_line = education_lines[i]

                if extract_date(next_line):
                    i += 1
                    continue

                if (
                    i + 1 < len(education_lines)
                    and detect_section(
                        education_lines[i]
                    )
                ):
                    break

                # A short title-like line following a
                # description may indicate a new entry.
                if (
                    len(next_line) < 90
                    and i > 0
                    and not next_line.endswith(".")
                ):

                    # Keep ordinary description text.
                    add_body_text(next_line)
                    i += 1

                else:

                    add_body_text(next_line)
                    i += 1

    # ========================================================
    # WORK EXPERIENCE
    # ========================================================

    experience_lines = sections.get(
        "WORK EXPERIENCE",
        []
    )

    if experience_lines:

        add_section_heading(
            "WORK EXPERIENCE"
        )

        i = 0

        while i < len(experience_lines):

            line = experience_lines[i]

            if extract_date(line):

                i += 1
                continue

            date = None
            date_index = None

            for j in range(
                i,
                min(i + 5, len(experience_lines))
            ):

                possible_date = extract_date(
                    experience_lines[j]
                )

                if possible_date:

                    date = possible_date
                    date_index = j
                    break

            table = document.add_table(
                rows=1,
                cols=2
            )

            table.alignment = WD_TABLE_ALIGNMENT.CENTER

            left = table.cell(0, 0)
            right = table.cell(0, 1)

            p_left = left.paragraphs[0]
            p_left.alignment = WD_ALIGN_PARAGRAPH.LEFT

            r_left = p_left.add_run(line)

            r_left.bold = True
            r_left.font.name = "Arial"
            r_left.font.size = Pt(10)

            if date:

                p_right = right.paragraphs[0]
                p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT

                r_right = p_right.add_run(
                    date
                )

                r_right.bold = True
                r_right.font.name = "Arial"
                r_right.font.size = Pt(9)

            i += 1

            # Organisation/company
            if i < len(experience_lines):

                possible_company = (
                    experience_lines[i]
                )

                if (
                    not extract_date(
                        possible_company
                    )
                    and not possible_company.startswith(
                        ("•", "-", "*")
                    )
                ):

                    add_body_text(
                        possible_company
                    )

                    i += 1

            # Description
            while i < len(experience_lines):

                next_line = experience_lines[i]

                if extract_date(next_line):

                    i += 1
                    continue

                # Stop if this looks like a new
                # position/company entry.
                if (
                    i + 1 < len(experience_lines)
                    and extract_date(
                        experience_lines[i + 1]
                    )
                    and len(next_line) < 100
                ):
                    break

                if next_line.startswith(
                    ("•", "-", "*")
                ):

                    add_bullet(
                        re.sub(
                            r"^[•\-*]\s*",
                            "",
                            next_line
                        )
                    )

                else:

                    add_body_text(
                        next_line
                    )

                i += 1

    # ========================================================
    # SKILLS — THREE COLUMN LAYOUT
    # ========================================================

    skills_lines = sections.get(
        "SKILLS",
        []
    )

    if skills_lines:

        add_section_heading("SKILLS")

        skills = []

        for line in skills_lines:

            cleaned = re.sub(
                r"^[•\-*]\s*",
                "",
                line
            ).strip()

            # Split comma/semicolon separated skills.
            parts = re.split(
                r"[,;|]",
                cleaned
            )

            for part in parts:

                skill = part.strip()

                if skill:
                    skills.append(skill)

        # Remove duplicates while preserving order.
        unique_skills = []

        seen = set()

        for skill in skills:

            key = skill.lower()

            if key not in seen:

                seen.add(key)
                unique_skills.append(skill)

        columns = 3

        rows = max(
            1,
            (len(unique_skills) + columns - 1)
            // columns
        )

        table = document.add_table(
            rows=rows,
            cols=columns
        )

        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        index = 0

        for row in table.rows:

            for cell in row.cells:

                cell.vertical_alignment = (
                    WD_CELL_VERTICAL_ALIGNMENT.CENTER
                )

                paragraph = cell.paragraphs[0]

                paragraph.alignment = (
                    WD_ALIGN_PARAGRAPH.LEFT
                )

                if index < len(unique_skills):

                    run = paragraph.add_run(
                        "• " + unique_skills[index]
                    )

                    run.font.name = "Arial"
                    run.font.size = Pt(9.5)

                    index += 1

    # ========================================================
    # SIMPLE SECTION RENDERER
    # ========================================================

    def render_simple_section(
        canonical_name,
        aliases=None,
        bullets=False
    ):

        content = sections.get(
            canonical_name,
            []
        )

        if not content:
            return

        add_section_heading(
            canonical_name
        )

        for line in content:

            cleaned = re.sub(
                r"^[•\-*]\s*",
                "",
                line
            ).strip()

            if not cleaned:
                continue

            if bullets:
                add_bullet(cleaned)
            else:
                add_body_text(cleaned)

    # ========================================================
    # CERTIFICATIONS
    # ========================================================

    render_simple_section(
        "CERTIFICATIONS",
        bullets=True
    )

    # ========================================================
    # ACHIEVEMENTS
    # ========================================================

    render_simple_section(
        "ACHIEVEMENTS",
        bullets=True
    )

    # ========================================================
    # PROJECTS
    # ========================================================

    render_simple_section(
        "PROJECTS",
        bullets=False
    )

    # ========================================================
    # LANGUAGES
    # ========================================================

    render_simple_section(
        "LANGUAGES",
        bullets=True
    )

    # ========================================================
    # REFERENCES
    # ========================================================

    render_simple_section(
        "REFERENCES",
        bullets=False
    )

    # ========================================================
    # FALLBACK FOR UNRECOGNISED INFORMATION
    # ========================================================

    known_sections = {
        "ABOUT ME",
        "EDUCATION",
        "WORK EXPERIENCE",
        "SKILLS",
        "CERTIFICATIONS",
        "ACHIEVEMENTS",
        "PROJECTS",
        "LANGUAGES",
        "REFERENCES",
        "CONTACT",
    }

    for section_name, content in sections.items():

        if section_name in known_sections:
            continue

        if not content:
            continue

        add_section_heading(
            section_name
        )

        for line in content:

            if line.strip():
                add_body_text(line)

    # ========================================================
    # FOOTER
    # ========================================================

    for doc_section in document.sections:

        footer = doc_section.footer

        paragraph = footer.paragraphs[0]

        paragraph.alignment = (
            WD_ALIGN_PARAGRAPH.CENTER
        )

        run = paragraph.add_run(
            "CVFix-SA | Your CV. Your Career."
        )

        run.font.name = "Arial"
        run.font.size = Pt(8)

    # ========================================================
    # SAVE
    # ========================================================

    filename = (
        "CVFix-SA_Professional_CV.docx"
    )

    document.save(filename)

    return filename


# ============================================================
# KEEP ORIGINAL FORMAT
# ============================================================
def create_original_format_cv():

    global website_cv_text
    global improved_cv_text
    global original_cv_file_path

    if not website_cv_text.strip():
        return None

    filename = "CVFix-SA_Original_Format_CV.docx"

    # --------------------------------------------------------
    # DOCX:
    # Preserve the user's original document structure,
    # formatting, tables and layout while improving wording.
    # --------------------------------------------------------

    if (
        "original_cv_file_path" in globals()
        and original_cv_file_path
        and str(original_cv_file_path).lower().endswith(".docx")
        and os.path.exists(original_cv_file_path)
    ):

        document = Document(original_cv_file_path)

        replacements = {
            "responsible for": "managed and delivered",
            "worked on": "contributed to",
            "helped with": "supported",
            "helped": "supported",
            "used": "utilised",
            "worked with": "collaborated with",
            "involved in": "contributed to"
        }

        def improve_paragraph(paragraph):

            if not paragraph.text.strip():
                return

            original_text = paragraph.text
            improved_text = original_text

            for weak, strong in replacements.items():

                pattern = re.compile(
                    re.escape(weak),
                    re.IGNORECASE
                )

                improved_text = pattern.sub(
                    strong,
                    improved_text
                )

            if improved_text != original_text:

                if paragraph.runs:

                    paragraph.runs[0].text = improved_text

                    for run in paragraph.runs[1:]:
                        run.text = ""

                else:

                    paragraph.text = improved_text

        # Normal paragraphs
        for paragraph in document.paragraphs:
            improve_paragraph(paragraph)

        # Tables
        for table in document.tables:

            for row in table.rows:

                for cell in row.cells:

                    for paragraph in cell.paragraphs:
                        improve_paragraph(paragraph)

        document.save(filename)

        return filename


    # --------------------------------------------------------
    # PDF fallback
    # --------------------------------------------------------

    source_text = improved_cv_text.strip()

    if not source_text:
        source_text = website_cv_text.strip()

    document = Document()

    section = document.sections[0]

    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    styles = document.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(10)

    for line in source_text.splitlines():

        line = line.strip()

        if line:

            p = document.add_paragraph()

            p.paragraph_format.space_after = Pt(3)

            r = p.add_run(line)

            r.font.name = "Arial"
            r.font.size = Pt(10)

    document.save(filename)

    return filename




# ============================================================
# FORMAT DISPATCHER
# ============================================================
def create_download(format_choice):

    global website_cv_text
    global improved_cv_text

    if not website_cv_text.strip():
        return None

    # --------------------------------------------------------
    # Use the improved wording if the user has run
    # "Improve My CV".
    # --------------------------------------------------------

    if improved_cv_text.strip():

        original_text = website_cv_text

        website_cv_text = improved_cv_text

        try:

            if format_choice == "Keep My Original Format":

                return create_original_format_cv()

            return create_cvfix_professional_cv()

        finally:

            website_cv_text = original_text

    # --------------------------------------------------------
    # If the user has not run improvement yet, use the
    # original uploaded CV content.
    # --------------------------------------------------------

    if format_choice == "Keep My Original Format":

        return create_original_format_cv()

    return create_cvfix_professional_cv()




# ============================================================
# CSS
# ============================================================
# CSS
# ============================================================

css = """

.gradio-container {
    max-width: 1150px !important;
    margin: auto !important;
}

.cvfix-hero {
    display: block !important;
    width: 100% !important;
    box-sizing: border-box !important;
    text-align: center !important;
    padding: 60px 25px !important;
    margin: 0 0 35px 0 !important;
    border-radius: 25px !important;
    background: linear-gradient(
        135deg,
        #12355B 0%,
        #1F6AA5 100%
    ) !important;
    color: #ffffff !important;
    min-height: 220px !important;
    visibility: visible !important;
    opacity: 1 !important;
}

.cvfix-how-it-works {
    width: 100% !important;
    box-sizing: border-box !important;
    padding: 32px 25px !important;
    margin: 0 0 35px 0 !important;
    border-radius: 22px !important;
    background: linear-gradient(
        135deg,
        #12355B 0%,
        #1F6AA5 100%
    ) !important;
    color: #ffffff !important;
    text-align: center !important;
    visibility: visible !important;
    opacity: 1 !important;
}

.cvfix-how-title {
    color: #ffffff !important;
    font-size: 28px !important;
    font-weight: 800 !important;
    margin-bottom: 25px !important;
}

.cvfix-how-steps {
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    flex-wrap: wrap !important;
    gap: 12px !important;
}

.cvfix-step {
    display: flex !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center !important;
    min-width: 95px !important;
}

.cvfix-step-icon {
    font-size: 30px !important;
    margin-bottom: 6px !important;
}

.cvfix-step-text {
    color: #ffffff !important;
    font-size: 16px !important;
    font-weight: 700 !important;
}

.cvfix-arrow {
    color: #ffffff !important;
    font-size: 25px !important;
    font-weight: 700 !important;
}


.cvfix-logo {
    display: block !important;
    color: #ffffff !important;
    font-size: 52px;
    font-weight: 800;
}

.cvfix-tagline {
    display: block !important;
    color: #ffffff !important;
    font-size: 25px;
    margin-top: 10px;
    font-weight: 600;
}

.cvfix-description {
    display: block !important;
    color: #ffffff !important;
    max-width: 700px;
    margin: 18px auto 0 auto;
    font-size: 17px;
    line-height: 1.6;
}

"""


# ============================================================
# WEBSITE
# ============================================================


# ============================================================
# CVFIX-SA REGULAR USER ACCOUNT FOUNDATION
# ============================================================

CVFIX_ACCOUNT_DB = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "cvfix_accounts_dev.db"
)


def cvfix_account_connection():
    connection = sqlite3.connect(CVFIX_ACCOUNT_DB)
    connection.row_factory = sqlite3.Row
    return connection


def cvfix_init_account_database():

    connection = cvfix_account_connection()

    connection.execute(
        "CREATE TABLE IF NOT EXISTS users ("
        "id INTEGER PRIMARY KEY AUTOINCREMENT, "
        "username TEXT NOT NULL UNIQUE, "
        "email TEXT NOT NULL UNIQUE, "
        "password_hash TEXT NOT NULL, "
        "password_salt TEXT NOT NULL, "
        "role TEXT NOT NULL DEFAULT 'user', "
        "created_at TEXT NOT NULL, "
        "last_login TEXT)"
    )

    connection.commit()
    connection.close()


def cvfix_hash_password(password, salt=None):

    if salt is None:
        salt = secrets.token_hex(16)

    password_hash = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("utf-8"),
        310000
    ).hex()

    return password_hash, salt


def cvfix_verify_password(
    password,
    stored_hash,
    stored_salt
):

    password_hash, _ = cvfix_hash_password(
        password,
        stored_salt
    )

    return secrets.compare_digest(
        password_hash,
        stored_hash
    )


cvfix_init_account_database()

# ============================================================
# END ACCOUNT DATABASE FOUNDATION
# ============================================================



# ============================================================
# CVFIX-SA REGULAR USER REGISTRATION + LOGIN
# ============================================================


def cvfix_validate_username(username):

    username = str(username).strip()

    if not username:
        return False, "Username is required."

    if len(username) < 3:
        return False, "Username must be at least 3 characters."

    if len(username) > 30:
        return False, "Username must not exceed 30 characters."

    if not re.fullmatch(
        r"[A-Za-z0-9_.-]+",
        username
    ):
        return False, (
            "Username may contain only letters, numbers, "
            "underscores, dots and hyphens."
        )

    return True, ""


def cvfix_validate_email(email):

    email = str(email).strip().lower()

    if not email:
        return False, "Email is required."

    if not re.fullmatch(
        r"[^@\s]+@[^@\s]+\.[^@\s]+",
        email
    ):
        return False, "Please enter a valid email address."

    return True, ""


def cvfix_register_user(
    username,
    email,
    password
):

    username = str(username).strip()
    email = str(email).strip().lower()
    password = str(password)

    valid_username, message = (
        cvfix_validate_username(username)
    )

    if not valid_username:
        return False, message

    valid_email, message = (
        cvfix_validate_email(email)
    )

    if not valid_email:
        return False, message

    if len(password) < 8:
        return False, (
            "Password must be at least 8 characters."
        )

    connection = cvfix_account_connection()

    try:

        existing_username = connection.execute(
            "SELECT id FROM users WHERE username = ?",
            (username,)
        ).fetchone()

        if existing_username:
            return False, "Username is already registered."

        existing_email = connection.execute(
            "SELECT id FROM users WHERE email = ?",
            (email,)
        ).fetchone()

        if existing_email:
            return False, "Email is already registered."

        password_hash, password_salt = (
            cvfix_hash_password(password)
        )

        created_at = datetime.utcnow().isoformat()

        connection.execute(
            "INSERT INTO users "
            "(username, email, password_hash, "
            "password_salt, role, created_at, last_login) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (
                username,
                email,
                password_hash,
                password_salt,
                "user",
                created_at,
                None
            )
        )

        connection.commit()

        return True, "Account created successfully."

    except Exception as error:

        connection.rollback()

        return False, (
            "Could not create account: "
            + str(error)
        )

    finally:

        connection.close()


def cvfix_login_user(
    username_or_email,
    password
):

    username_or_email = (
        str(username_or_email).strip()
    )

    password = str(password)

    # --------------------------------------------------------
    # EXISTING OWNER LOGIN
    # --------------------------------------------------------
    #
    # This uses the existing environment variables and secure
    # comparison. It is intentionally kept in this function
    # only so owner and regular-user login can share the UI.
    #
    # OWNER CREDENTIALS ARE NOT STORED IN SQLITE.
    # --------------------------------------------------------

    owner_username = os.environ.get(
        "CVFIX_OWNER_USERNAME",
        ""
    )

    owner_password = os.environ.get(
        "CVFIX_OWNER_PASSWORD",
        ""
    )

    if owner_username and owner_password:

        username_match = hmac.compare_digest(
            username_or_email,
            str(owner_username)
        )

        password_match = hmac.compare_digest(
            password,
            str(owner_password)
        )

        if username_match and password_match:

            return {
                "id": "OWNER",
                "username": str(owner_username),
                "email": "",
                "role": "owner",
                "created_at": "",
                "last_login": ""
            }

    # --------------------------------------------------------
    # REGULAR USER LOGIN
    # --------------------------------------------------------

    connection = cvfix_account_connection()

    row = connection.execute(
        "SELECT id, username, email, password_hash, "
        "password_salt, role, created_at, last_login "
        "FROM users "
        "WHERE username = ? OR email = ? "
        "LIMIT 1",
        (
            username_or_email,
            username_or_email.lower()
        )
    ).fetchone()

    if row is None:

        connection.close()

        return None

    valid_password = cvfix_verify_password(
        password,
        row["password_hash"],
        row["password_salt"]
    )

    if not valid_password:

        connection.close()

        return None

    login_time = datetime.utcnow().isoformat()

    connection.execute(
        "UPDATE users "
        "SET last_login = ? "
        "WHERE id = ?",
        (
            login_time,
            row["id"]
        )
    )

    connection.commit()
    connection.close()

    return {
        "id": row["id"],
        "username": row["username"],
        "email": row["email"],
        "role": row["role"],
        "created_at": row["created_at"],
        "last_login": login_time
    }


def cvfix_get_user_by_id(user_id):

    if user_id == "OWNER":

        owner_username = os.environ.get(
            "CVFIX_OWNER_USERNAME",
            ""
        )

        return {
            "id": "OWNER",
            "username": owner_username,
            "email": "",
            "role": "owner",
            "created_at": "",
            "last_login": ""
        }

    connection = cvfix_account_connection()

    row = connection.execute(
        "SELECT id, username, email, role, "
        "created_at, last_login "
        "FROM users "
        "WHERE id = ? "
        "LIMIT 1",
        (user_id,)
    ).fetchone()

    connection.close()

    if row is None:
        return None

    return {
        "id": row["id"],
        "username": row["username"],
        "email": row["email"],
        "role": row["role"],
        "created_at": row["created_at"],
        "last_login": row["last_login"]
    }


# ============================================================
# END REGULAR USER REGISTRATION + LOGIN
# ============================================================




# ============================================================
# CVFIX-SA ACCOUNT LOGIN UI WRAPPER
# ============================================================

def cvfix_account_login_ui(
    username_or_email,
    password
):

    user = cvfix_login_user(
        username_or_email,
        password
    )

    if user is None:

        return (
            None,
            "### ❌ Login failed\n\n"
            "Incorrect username/email or password.",
            gr.update(visible=False),
            gr.update(visible=False),
            ""
        )

    return (
        user,
        (
            "### ✅ Login successful\n\n"
            f"Welcome, **{user['username']}**."
        ),
        gr.update(visible=True),
        gr.update(visible=True),
        cvfix_account_status_ui(user)
    )


# CVFIX-SA ACCOUNT STATUS DISPLAY
def cvfix_account_status_ui(user):

    if not isinstance(user, dict) or not user.get("id"):
        return ""

    username = str(user.get("username", "")).strip()
    email = str(user.get("email", "")).strip()

    if email:
        return (
            "### 👤 Account\n\n"
            f"**Welcome, {username}**\n\n"
            f"**Email:** {email}"
        )

    return (
        "### 👤 Account\n\n"
        f"**Welcome, {username}**"
    )


# CVFIX-SA BROWSER SESSION RESTORE
def cvfix_restore_session_ui(user):

    logged_in = (
        isinstance(user, dict)
        and user.get("id")
    )

    return (
        gr.update(visible=bool(logged_in)),
        gr.update(visible=bool(logged_in)),
        cvfix_account_status_ui(user)
    )


# CVFIX-SA ACCOUNT LOGOUT
def cvfix_account_logout_ui():

    return (
        None,
        "### 👋 Logged out successfully\n\n"
        "You have been logged out of CVFix-SA.",
        gr.update(visible=False),
        gr.update(visible=False),
        ""
    )


# CVFIX-SA REGISTRATION UI WRAPPER
def cvfix_account_register_ui(
    username,
    email,
    password
):
    success, message = cvfix_register_user(
        username,
        email,
        password
    )

    if success:
        return (
            "### ✅ Account created successfully\n\n"
            "Your CVFix-SA account has been created. "
            "You can now use the **🔐 Login** tab to sign in."
        )

    return (
        "### ❌ Error\n\n"
        + str(message)
    )


with gr.Blocks(
    title="CVFix-SA"
) as app:

    gr.HTML(
        """
        <div class="cvfix-hero">

            <div class="cvfix-logo">
                CVFix-SA
            </div>

            <div class="cvfix-tagline">
                Your CV. Your Career.
            </div>

            <div class="cvfix-description">
                Analyse your CV, match it with a job,
                improve your application and create
                a stronger CV.
            </div>

        </div>
        """
    )

    gr.HTML(
        """
        <div class="cvfix-how-it-works">

            <div class="cvfix-how-title">
                🚀 How CVFix-SA Works
            </div>

            <div class="cvfix-how-steps">
                <div class="cvfix-step">
                    <div class="cvfix-step-icon">📄</div>
                    <div class="cvfix-step-text">Upload</div>
                </div>

                <div class="cvfix-arrow">→</div>

                <div class="cvfix-step">
                    <div class="cvfix-step-icon">💼</div>
                    <div class="cvfix-step-text">Add Job</div>
                </div>

                <div class="cvfix-arrow">→</div>

                <div class="cvfix-step">
                    <div class="cvfix-step-icon">🔍</div>
                    <div class="cvfix-step-text">Analyse</div>
                </div>

                <div class="cvfix-arrow">→</div>

                <div class="cvfix-step">
                    <div class="cvfix-step-icon">✨</div>
                    <div class="cvfix-step-text">Improve</div>
                </div>

                <div class="cvfix-arrow">→</div>

                <div class="cvfix-step">
                    <div class="cvfix-step-icon">⬇️</div>
                    <div class="cvfix-step-text">Download</div>
                </div>
            </div>

        </div>
        """
    )


    # ========================================================
    # CVFIX-SA REGULAR USER ACCOUNT UI
    # ========================================================

    account_user = gr.BrowserState(
        default_value=None,
        storage_key="cvfix_sa_account_user",
        secret=os.environ.get(
            "CVFIX_BROWSER_STATE_SECRET",
            "CVFix-SA-development-browser-state-secret"
        )
    )

    with gr.Group():

        gr.Markdown(
            """
            ## 👤 CVFix-SA Account

            Create an account to use CVFix-SA, or log in if you
            already have an account.
            """
        )

        with gr.Tabs():

            # ------------------------------------------------
            # REGISTER
            # ------------------------------------------------

            with gr.Tab("📝 Create Account"):

                register_username = gr.Textbox(
                    label="Username",
                    placeholder="Choose a username",
                    max_lines=1
                )

                register_email = gr.Textbox(
                    label="Email",
                    placeholder="Enter your email address",
                    max_lines=1
                )

                register_password = gr.Textbox(
                    label="Password",
                    placeholder="Minimum 8 characters",
                    type="password",
                    max_lines=1
                )

                register_button = gr.Button(
                    "📝 Create My Account",
                    variant="primary"
                )

                register_message = gr.Markdown()

                register_button.click(
                    fn=cvfix_account_register_ui,
                    inputs=[
                        register_username,
                        register_email,
                        register_password
                    ],
                    outputs=register_message,
                    queue=False
                )

            # ------------------------------------------------
            # LOGIN
            # ------------------------------------------------

            with gr.Tab("🔐 Login"):

                login_username = gr.Textbox(
                    label="Username or Email",
                    placeholder="Enter your username or email",
                    max_lines=1
                )

                login_password = gr.Textbox(
                    label="Password",
                    placeholder="Enter your password",
                    type="password",
                    max_lines=1
                )

                login_button = gr.Button(
                    "🔐 Log In",
                    variant="primary"
                )

                login_message = gr.Markdown()

                logout_button = gr.Button(
                    "🚪 Log Out",
                    variant="secondary",
                    visible=False
                )

                account_status = gr.Markdown()


        gr.Markdown(
            """
            **Your account is separate from your CV content.**
            Your password is securely hashed before it is stored.
            """
        )

    gr.Markdown("---")

    with gr.Group(visible=False) as cv_workspace:

        with gr.Tabs():

            with gr.Tab("📄 Upload CV"):

                gr.Markdown(
                    "# Upload Your CV\n\n"
                    "Supported formats: **PDF** and **DOCX**"
                )

                cv_file = gr.File(
                    label="Choose your CV",
                    file_types=[".pdf", ".docx"],
                    type="filepath"
                )

                upload_result = gr.Markdown()

                cv_file.change(
                    fn=upload_cv,
                    inputs=cv_file,
                    outputs=upload_result
                )

            with gr.Tab("🔍 Analyse"):

                gr.Markdown(
                    "# 🔍 Analyse Your CV"
                )

                job_description = gr.Textbox(
                    label="💼 Job Description",
                    placeholder=(
                        "Paste the complete job description here..."
                    ),
                    lines=15
                )

                analyse_button = gr.Button(
                    "🔍 Analyse My CV",
                    variant="primary"
                )

                analysis_result = gr.Markdown()

                with gr.Row():

                    overall = gr.Textbox(
                        label="🏆 Overall",
                        interactive=False
                    )

                    structure = gr.Textbox(
                        label="📄 Structure",
                        interactive=False
                    )

                    match = gr.Textbox(
                        label="💼 Job Match",
                        interactive=False
                    )

                    skills = gr.Textbox(
                        label="🛠 Skills",
                        interactive=False
                    )

                analyse_button.click(
                    fn=analyse_cv,
                    inputs=job_description,
                    outputs=[
                        analysis_result,
                        overall,
                        structure,
                        match,
                        skills
                    ]
                )

            with gr.Tab("✨ Improve"):

                gr.Markdown(
                    "# ✨ Improve Your CV"
                )

                improve_button = gr.Button(
                    "✨ Improve My CV",
                    variant="primary"
                )

                improvement_result = gr.Markdown()

                improve_button.click(
                    fn=improve_cv,
                    inputs=job_description,
                    outputs=improvement_result
                )

        
            with gr.Tab("⬇️ Download"):

                gr.Markdown(
                    "# ⬇️ Download Your CV\n\n"
                    "Choose how you want CVFix-SA to format your CV."
                )

                cv_format = gr.Radio(
                    choices=[
                        "CVFix-SA Professional Format",
                        "Keep My Original Format"
                    ],
                    value="CVFix-SA Professional Format",
                    label="📄 Choose Your CV Format",
                    info=(
                        "CVFix-SA Professional Format uses the "
                        "CVFix-SA structured professional layout. "
                        "Keep My Original Format preserves your "
                        "original Word document formatting when possible."
                    )
                )

                download_button = gr.Button(
                    "📄 Generate My CV",
                    variant="primary"
                )

                download_file = gr.File(
                    label="Your CV"
                )

                download_button.click(
                    fn=create_download,
                    inputs=cv_format,
                    outputs=download_file
                )




    
    # ========================================================
    # CVFIX-SA BROWSER SESSION RESTORE
    # Restore workspace visibility after page refresh.
    # ========================================================
    app.load(
        fn=cvfix_restore_session_ui,
        inputs=account_user,
        outputs=[
            cv_workspace,
            logout_button,
            account_status
        ]
    )


    # ========================================================
    # CVFIX-SA LOGIN EVENT
    # Registered after cv_workspace exists.
    # ========================================================
    login_button.click(
        fn=cvfix_account_login_ui,
        inputs=[
            login_username,
            login_password
        ],
        outputs=[
            account_user,
            login_message,
            cv_workspace,
            logout_button,
            account_status
        ]
    )


    # ========================================================
    # CVFIX-SA LOGOUT EVENT
    # ========================================================
    logout_button.click(
        fn=cvfix_account_logout_ui,
        inputs=[],
        outputs=[
            account_user,
            login_message,
            cv_workspace,
            logout_button,
            account_status
        ]
    )

    gr.Markdown(
        """
        ---

        ### CVFix-SA

        **Your CV. Your Career.**

        © 2026 CVFix-SA
        """
    )


    # ========================================================
    # FOUNDER
    # ========================================================

    gr.HTML(
        """
        <div style="
            margin-top: 35px;
            padding: 30px 25px;
            text-align: center;
            border-radius: 20px;
            background: #f5f7fa;
        ">

            <h2>👨🏽‍💻 Founder</h2>

            <h3>
                Junior Software Developer Mr BF Manikela
            </h3>

            <p>
                Creator and developer of CVFix-SA.
            </p>

        </div>
        """
    )

    # ========================================================
    # CONTACT US
    # ========================================================

    gr.HTML(
        """
        <div style="
            margin-top: 25px;
            padding: 30px 25px;
            text-align: center;
            border-radius: 20px;
            border: 1px solid #e5e7eb;
        ">

            <h2>📞 Contact Us</h2>

            <p>
                Have a question or need assistance with CVFix-SA?
            </p>

            <p>
                <strong>Phone:</strong> 079 852 5726
            </p>

            <p>
                <strong>Monday – Friday</strong><br>
                9:00 AM – 9:00 PM
            </p>

        </div>
        """
    )


cvfix_seo = """
<meta name="google-site-verification" content="Nyz8jJ9FEv8NCKHYVmXrzWqCksVPm6uZrJKQkgdR_vo" />
<meta name="description" content="CVFix-SA — Professional CV creation and career tools by Junior Software Developer Mr BF Manikela.">
<meta name="author" content="Mr BF Manikela">
<meta name="robots" content="index, follow">
<title>CVFix-SA | Professional CV Builder</title>
"""

app.launch(
    server_name="0.0.0.0",
    server_port=int(os.environ.get("PORT", 7860)),
    share=False,
    css=css,
    head=cvfix_seo,
)
