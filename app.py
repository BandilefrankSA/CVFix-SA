import os

import gradio as gr
from pypdf import PdfReader
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


# ============================================================
# CV STORAGE
# ============================================================

website_cv_text = ""


# ============================================================
# CV UPLOAD
# ============================================================

def upload_cv(file):

    global website_cv_text

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
# BASIC CV ANALYSIS
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

    cv_lower = website_cv_text.lower()
    job_lower = job_description.lower()

    words = set(
        word.strip(
            ".,!?;:()[]{}"
        )
        for word in job_lower.split()
    )

    words = {
        word for word in words
        if len(word) > 2
    }

    matched = [
        word for word in words
        if word in cv_lower
    ]

    if words:

        match_score = int(
            len(matched) / len(words) * 100
        )

    else:

        match_score = 0

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

    skills = [
        "python",
        "sql",
        "excel",
        "power bi",
        "tableau",
        "statistics",
        "data analysis",
        "machine learning",
        "communication",
        "leadership",
        "project management"
    ]

    detected_skills = [
        skill for skill in skills
        if skill in cv_lower
    ]

    skill_score = min(
        100,
        len(detected_skills) * 10
    )

    overall = int(
        structure_score * 0.30
        + match_score * 0.50
        + skill_score * 0.20
    )

    missing = [
        word for word in words
        if word not in cv_lower
    ]

    matching_text = (
        "\n".join(
            f"✓ {word}"
            for word in sorted(matched)[:30]
        )
        if matched
        else "None detected."
    )

    missing_text = (
        "\n".join(
            f"⚠ {word}"
            for word in sorted(missing)[:30]
        )
        if missing
        else "None detected."
    )

    report = f"""
# 📊 CVFix-SA Analysis

## Overall Score: {overall}/100

| Category | Score |
|---|---:|
| 📄 CV Structure | {structure_score}/100 |
| 💼 Job Match | {match_score}/100 |
| 🛠 Skills | {skill_score}/100 |

---

## ✅ Matching Keywords

{matching_text}

---

## ⚠️ Potential Missing Keywords

{missing_text}

---

## 💡 Recommendation

Tailor your CV to the requirements of the job,
but only include skills and experience that are
genuinely true about you.
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

        if weak in improved.lower():

            recommendations.append(
                f"Consider replacing '{weak}' "
                f"with '{strong}'."
            )

            improved = improved.replace(
                weak,
                strong
            )

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

def create_download():

    if not website_cv_text.strip():

        return None

    filename = "CVFix-SA_Improved_CV.docx"

    document = Document()

    title = document.add_paragraph()

    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = title.add_run(
        "CVFix-SA Improved CV"
    )

    run.bold = True
    run.font.size = Pt(20)

    for line in website_cv_text.splitlines():

        line = line.strip()

        if line:

            document.add_paragraph(line)

    document.save(filename)

    return filename


# ============================================================
# CSS
# ============================================================

css = """

.gradio-container {
    max-width: 1150px !important;
    margin: auto !important;
}

.cvfix-hero {
    text-align: center;
    padding: 60px 25px;
    margin-bottom: 35px;
    border-radius: 25px;
    background: linear-gradient(
        135deg,
        #12355B,
        #1F6AA5
    );
    color: white;
}

.cvfix-logo {
    font-size: 52px;
    font-weight: 800;
}

.cvfix-tagline {
    font-size: 25px;
    margin-top: 10px;
    font-weight: 600;
}

.cvfix-description {
    max-width: 700px;
    margin: 18px auto 0 auto;
    font-size: 17px;
    line-height: 1.6;
}

"""


# ============================================================
# WEBSITE
# ============================================================

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

    gr.Markdown(
        """
        ## 🚀 How CVFix-SA Works

        **📄 Upload → 💼 Add Job → 🔍 Analyse → ✨ Improve → ⬇️ Download**
        """
    )

    gr.Markdown("---")

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
                "# ⬇️ Download Your CV"
            )

            download_button = gr.Button(
                "📄 Generate Improved CV",
                variant="primary"
            )

            download_file = gr.File(
                label="CVFix-SA CV"
            )

            download_button.click(
                fn=create_download,
                outputs=download_file
            )

    gr.Markdown(
        """
        ---

        ### CVFix-SA

        **Your CV. Your Career.**

        © 2026 CVFix-SA
        """
    )


cvfix_seo = """
<meta name="google-site-verification" content="Nyz8jJ9FEv8NCKHYVmXrzWqCksVPm6uZrJKQkgdR_vo">
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
    head=cvfix_seo
)
