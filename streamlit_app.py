import io
import os
import re
import sys
import subprocess
import streamlit as st
import zipfile
import xml.etree.ElementTree as ET

try:
    import openai
except ImportError:
    openai = None

try:
    from docx import Document
except ImportError:
    Document = None

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import pypdf
except ImportError:
    pypdf = None


def ensure_pdf_parser_available() -> None:
    global pdfplumber, pypdf
    if pdfplumber is not None or pypdf is not None:
        return

    try:
        import importlib
        pdfplumber = importlib.import_module("pdfplumber")
    except ImportError:
        pdfplumber = None

    try:
        import importlib
        pypdf = importlib.import_module("pypdf")
    except ImportError:
        pypdf = None

    if pdfplumber is not None or pypdf is not None:
        return

    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "pypdf", "pdfplumber"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        import importlib
        try:
            pdfplumber = importlib.import_module("pdfplumber")
        except ImportError:
            pdfplumber = None
        try:
            pypdf = importlib.import_module("pypdf")
        except ImportError:
            pypdf = None
    except Exception:
        pass

st.set_page_config(
    page_title="Career & CV Booster",
    page_icon="🚀",
    layout="wide",
)

PAGE_HOME = "home"
PAGE_RESUME = "resume_analyzer"
PAGE_LINKEDIN = "linkedin_score"
PAGE_PREMIUM = "premium_profiles"
PAGE_JOBTRACKER = "job_tracker"

if "page" not in st.session_state:
    st.session_state.page = PAGE_HOME


def calculate_ats_score(resume_text: str) -> int:
    keywords = [
        "leadership",
        "project",
        "results",
        "strategy",
        "communication",
        "analysis",
        "team",
        "managed",
        "delivered",
    ]
    score = 50
    text_lower = resume_text.lower()
    score += sum(5 for keyword in keywords if keyword in text_lower)
    score += min(20, len(resume_text.split()) // 20)
    return max(0, min(score, 100))


def estimate_linkedin_score(profile_text: str) -> int:
    keywords = [
        "experienced",
        "driven",
        "innovative",
        "collaborative",
        "results",
        "digital",
        "strategy",
        "growth",
        "leadership",
    ]
    score = 45
    text_lower = profile_text.lower()
    score += sum(4 for keyword in keywords if keyword in text_lower)
    score += min(15, len(profile_text.split()) // 25)
    return max(0, min(score, 100))


def reset_to_home():
    st.session_state.page = PAGE_HOME


def extract_text_from_docx(uploaded_file) -> str:
    uploaded_file.seek(0)
    if Document is not None:
        try:
            document = Document(uploaded_file)
            return "\n\n".join(p.text for p in document.paragraphs if p.text)
        except Exception:
            uploaded_file.seek(0)

    try:
        uploaded_file.seek(0)
        with zipfile.ZipFile(uploaded_file) as docx_zip:
            xml_content = docx_zip.read("word/document.xml")
        tree = ET.fromstring(xml_content)
        namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        paragraphs = []
        for paragraph in tree.findall('.//w:p', namespaces):
            texts = [node.text for node in paragraph.findall('.//w:t', namespaces) if node.text]
            if texts:
                paragraphs.append(''.join(texts))
        return "\n\n".join(paragraphs)
    except Exception as exc:
        raise RuntimeError(
            "Failed to extract text from DOCX. Install python-docx or upload a plain text resume file."
        ) from exc


def create_docx_from_text(text: str) -> io.BytesIO:
    buffer = io.BytesIO()
    if Document is not None:
        document = Document()
        for line in text.splitlines():
            if line.strip():
                document.add_paragraph(line)
        document.save(buffer)
        buffer.seek(0)
    else:
        buffer.write(text.encode("utf-8", errors="ignore"))
        buffer.seek(0)
    return buffer


def extract_text_from_pdf(uploaded_file) -> str:
    ensure_pdf_parser_available()
    uploaded_file.seek(0)
    extracted_text = ""
    if pdfplumber is not None:
        with pdfplumber.open(uploaded_file) as pdf:
            pages = [page.extract_text() or "" for page in pdf.pages]
        extracted_text = "\n\n".join(pages)
    elif pypdf is not None:
        uploaded_file.seek(0)
        reader = pypdf.PdfReader(uploaded_file)
        pages = [page.extract_text() or "" for page in reader.pages]
        extracted_text = "\n\n".join(pages)
    else:
        raise RuntimeError(
            "PDF extraction requires pdfplumber or pypdf. "
            "Install one with `pip install pdfplumber` or `pip install pypdf` and restart the app."
        )

    if not extracted_text.strip():
        return ""

    if Document is not None:
        docx_stream = create_docx_from_text(extracted_text)
        return extract_text_from_docx(docx_stream)

    return extracted_text


def extract_text_from_uploaded_file(uploaded_file) -> str:
    if uploaded_file is None:
        return ""
    file_type = uploaded_file.type.lower()
    filename = uploaded_file.name.lower()
    if file_type == "application/pdf" or filename.endswith(".pdf"):
        return extract_text_from_pdf(uploaded_file)
    if filename.endswith(".docx") or file_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        return extract_text_from_docx(uploaded_file)
    if filename.endswith(".doc") or file_type == "application/msword":
        uploaded_file.seek(0)
        try:
            return extract_text_from_docx(uploaded_file)
        except Exception:
            uploaded_file.seek(0)
            raw = uploaded_file.read()
            try:
                return raw.decode("utf-8")
            except Exception:
                return raw.decode("latin-1", errors="ignore")
    uploaded_file.seek(0)
    raw = uploaded_file.read()
    try:
        return raw.decode("utf-8")
    except Exception:
        return raw.decode("latin-1", errors="ignore")


def build_llm_prompt(resume_text: str, jd_text: str) -> str:
    return f"""
Purpose: “Am I even qualifying?”

Act as an ATS and recruiter screening system.

Compare the Job Description and my Resume.

Tell me:

1. Match score out of 100

2. Would this resume be shortlisted? (Yes/No)

3. Top 5 reasons for rejection or shortlisting

4. Top 5 missing or weak requirements

5. Keyword match quality (High/Medium/Low)

Be direct and specific.

### 10-Second Rejection
Purpose: “Why am I getting ignored?”

Act as a busy hiring manager scanning 200 resumes.

You have 10 seconds.

Tell me:

1. First impression in one line

2. Would you continue reading or reject? (Continue/Reject)

3. Why would you reject this profile instantly?

4. What is unclear or confusing?

5. Does this feel relevant to the role? Why or why not?

Be blunt and realistic.

### Final Diagnosis
Purpose: “What is the real problem?”

Based on the Job Description and Resume, give a final diagnosis.

Tell me:

1. The top 3 reasons this candidate is not getting interview calls

2. The biggest gap between the JD and the resume

3. The biggest perception problem in the resume

4. What stronger candidates are doing better

5. One single change that would improve shortlisting chances the most

Keep it simple, clear, and actionable. Do not exceed the brief requirements.

Job Description:
{jd_text}

Resume:
{resume_text}
"""


def build_linkedin_analysis_prompt(profile_text: str, job_title: str) -> str:
    return f"""
Analyze this LinkedIn profile for the given position and generate a score for this profile, separately for headline, bio, and experience.

Provide:

1. Headline score out of 100
2. Bio/summary score out of 100
3. Experience score out of 100
4. Missing keywords related to the desired job title
5. Practical suggestions to make the profile more discoverable and compelling

The desired job title is: {job_title}

Profile extract:
{profile_text}

Return the output as a concise professional AI report with clear headings and score blocks.
"""


def local_linkedin_profile_report(profile_text: str, job_title: str) -> str:
    clean_text = profile_text.lower()
    title_terms = [w for w in re.findall(r"[a-zA-Z]+", job_title.lower()) if len(w) > 3]
    missing_keywords = [term for term in title_terms if term not in clean_text]
    if not missing_keywords:
        missing_keywords = [
            "leadership",
            "strategy",
            "impact",
            "results",
            "executive",
        ]

    header_score = min(100, 50 + sum(8 for term in ["leader", "director", "head", "executive", "manager"] if term in clean_text))
    bio_score = min(100, 45 + sum(5 for term in ["growth", "strategy", "experience", "driven", "focused"] if term in clean_text))
    experience_score = min(100, 40 + sum(6 for term in ["delivered", "launched", "improved", "revenue", "managed"] if term in clean_text))

    return (
        "### LinkedIn AI Review Report\n"
        f"Desired job title: {job_title}\n\n"
        f"Headline score: {header_score}/100\n"
        f"Bio score: {bio_score}/100\n"
        f"Experience score: {experience_score}/100\n\n"
        "Missing discoverability keywords:\n"
        + "\n".join(f"- {keyword}" for keyword in missing_keywords)
        + "\n\n"
        "Suggestions:\n"
        "- Strengthen your headline with the exact role and a value statement.\n"
        "- Add industry and job-specific keywords to your summary.\n"
        "- Highlight measurable achievements and business impact in your experience section.\n"
        "- Use the target role wording consistently across headline, summary, and experience.\n"
        "\n" "Disclaimer: This is an AI-generated profile review. For human-reviewed premium LinkedIn profile services and tailored rewriting, upgrade to our premium offering."
    )


def generate_linkedin_profile_report(profile_text: str, job_title: str) -> str:
    profile_text = truncate_text(profile_text, 3200)
    if openai is not None and os.environ.get("OPENAI_API_KEY"):
        prompt = build_linkedin_analysis_prompt(profile_text, job_title)
        try:
            return call_llm(prompt)
        except Exception as exc:
            print(f"[Developer] LinkedIn OpenAI analysis failed: {exc}")
            return local_linkedin_profile_report(profile_text, job_title)
    print("[Developer] OpenAI is not configured or unavailable; using local LinkedIn fallback report.")
    return local_linkedin_profile_report(profile_text, job_title)


def truncate_text(text: str, limit: int = 3000) -> str:
    if len(text) > limit:
        return text[:limit] + "\n\n... [Truncated to fit prompt limits]"
    return text


def normalize_ats_score_in_report(score: int, report: str) -> str:
    if "Match score out of 100:" in report:
        return re.sub(r"Match score out of 100:\s*\d{1,3}", f"Match score out of 100: {score}", report, count=1)
    return f"Match score out of 100: {score}\n\n{report}"


def call_llm(prompt: str) -> str:
    if openai is None:
        raise RuntimeError("OpenAI client is unavailable. Install openai to enable ChatGPT integration.")
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY is not set. Set this environment variable to use ChatGPT.")
    openai.api_key = api_key
    model = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")
    response = openai.ChatCompletion.create(
        model=model,
        messages=[
            {"role": "system", "content": "You are a professional ATS and recruiter screening assistant."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.2,
        max_tokens=1100,
    )
    return response.choices[0].message.content.strip()


def local_resume_jd_analysis(resume_text: str, jd_text: str) -> str:
    score = calculate_ats_score(resume_text)
    shortlist = "Yes" if score >= 72 else "No"
    keywords = ["leadership", "strategy", "delivery", "team", "stakeholder", "product", "results", "communication"]
    missing = [kw for kw in keywords if kw not in resume_text.lower() and kw in jd_text.lower()]
    missing = missing[:5] or ["Additional role-specific keywords and metrics"]
    return (
        "### ATS + Recruiter Summary\n"
        f"Match score out of 100: {score}\n\n"
        f"Would this resume be shortlisted? {shortlist}\n\n"
        "Top 5 reasons for rejection or shortlisting:\n"
        "- Limited role-specific keyword coverage\n"
        "- Few measurable achievements in the resume\n"
        "- Resume wording is not tightly aligned with the target JD\n"
        "- Experience descriptions are generic rather than outcome-driven\n"
        "- Profile lacks clear senior-level impact language\n\n"
        "Top 5 missing or weak requirements:\n"
        + "\n".join(f"- {item}" for item in missing)
        + "\n\n"
        "Keyword match quality: Medium\n\n"
        "### 10-Second Rejection\n"
        "First impression in one line: Resume reads as competent but not sharply tailored to the role.\n"
        f"Would you continue reading or reject? {'Continue' if score >= 60 else 'Reject'}\n"
        "Why would you reject this profile instantly? Because the resume lacks specific achievements and does not surface the most relevant skills quickly.\n"
        "What is unclear or confusing? The career narrative and the main value proposition are not immediately obvious.\n"
        "Does this feel relevant to the role? Why or why not? It feels partly relevant, but not clearly aligned with the job description because key terms and impact statements are missing.\n\n"
        "### Final Diagnosis\n"
        "Top 3 reasons this candidate is not getting interview calls:\n"
        "- Resume is not tailored to the job description.\n"
        "- Weak use of metrics and outcomes in experience statements.\n"
        "- Core strengths are not positioned as clearly career-enhancing value.\n\n"
        "Biggest gap between the JD and the resume: The resume does not match the target role's required keywords and outcomes closely enough.\n"
        "Biggest perception problem in the resume: It reads as generic rather than purpose-built for the role.\n"
        "What stronger candidates are doing better: They show impact with specific metrics, use the target role language, and lead with outcomes.\n"
        "One single change that would improve shortlisting chances the most: Rewrite the experience bullets to include clear results, metrics, and the most relevant job keywords.\n"
    )


def analyze_resume_with_jd(resume_text: str, jd_text: str, score: int | None = None) -> str:
    resume_text = truncate_text(resume_text, 2800)
    jd_text = truncate_text(jd_text, 2800)
    if score is None:
        score = calculate_ats_score(resume_text)
    if openai is not None and os.environ.get("OPENAI_API_KEY"):
        prompt = build_llm_prompt(resume_text, jd_text)
        try:
            return normalize_ats_score_in_report(score, call_llm(prompt))
        except Exception as exc:
            print(f"[Developer] OpenAI analysis failed: {exc}")
            return normalize_ats_score_in_report(score, local_resume_jd_analysis(resume_text, jd_text))
    print("[Developer] OpenAI is not configured or unavailable; using local fallback report.")
    return normalize_ats_score_in_report(score, local_resume_jd_analysis(resume_text, jd_text))


st.markdown("""
<style>
:root {
    color-scheme: dark;
}
body, .stApp, .block-container, .css-1d391kg, .css-1v0mbdj, .element-container {
    color: #e2e8f0 !important;
    background: transparent !important;
}
body {
    background: radial-gradient(circle at top, #0f172a 0%, #020617 55%, #020617 100%);
}
.block-container {
    padding: 2rem 2rem 3rem;
}
.stMarkdown h1, .stMarkdown h2, .stMarkdown h3, .stMarkdown h4, .stMarkdown p, .stMarkdown li, .stMarkdown strong, .stMarkdown em, .stMarkdown a {
    color: #e2e8f0 !important;
}
.stMarkdown a {
    color: #60a5fa !important;
    text-decoration: none;
}
.stMarkdown a:hover {
    text-decoration: underline;
}
hr {
    border-top: 1px solid rgba(148, 163, 184, 0.24);
}
.section-card {
    border-radius: 28px;
    padding: 28px;
    background: rgba(15, 23, 42, 0.94);
    box-shadow: 0 32px 80px rgba(0, 0, 0, 0.35);
    border: 1px solid rgba(148, 163, 184, 0.18);
    min-height: 320px;
    margin-bottom: 20px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    box-sizing: border-box;
}
.section-card h3 {
    color: #f8fafc;
    margin-bottom: 1rem;
}
.section-card p {
    color: #cbd5e1;
    line-height: 1.75;
}
.hero-card {
    border-radius: 30px;
    padding: 44px;
    background: linear-gradient(145deg, rgba(15, 23, 42, 0.96) 0%, rgba(30, 41, 59, 0.92) 48%, rgba(59, 130, 246, 0.08) 100%);
    box-shadow: 0 35px 120px rgba(15, 23, 42, 0.58);
    border: 1px solid rgba(96, 165, 250, 0.18);
    position: relative;
    overflow: hidden;
}
.hero-card::before {
    content: '';
    position: absolute;
    top: -20%;
    right: -20%;
    width: 240px;
    height: 240px;
    background: rgba(99, 102, 241, 0.18);
    border-radius: 50%;
    filter: blur(40px);
}
.hero-card h1,
.hero-card p,
.hero-card .cta-button {
    color: #f8fafc !important;
}
.hero-card h1 {
    font-size: 3.2rem;
    line-height: 1.02;
    margin-bottom: 0.5rem;
}
.hero-card p {
    color: #cbd5e1;
    max-width: 840px;
    line-height: 1.8;
}
.hero-badges {
    display: flex;
    flex-wrap: wrap;
    gap: 12px;
    margin-top: 26px;
}
.hero-badge {
    padding: 14px 18px;
    border-radius: 18px;
    background: rgba(37, 99, 235, 0.12);
    color: #e0f2fe;
    font-weight: 600;
    border: 1px solid rgba(59, 130, 246, 0.18);
}
.cta-button {
    margin-top: 22px;
    display: inline-block;
    background: linear-gradient(135deg, #2563eb 0%, #4338ca 100%);
    color: #f8fafc;
    padding: 14px 28px;
    border-radius: 999px;
    text-decoration: none;
    border: 1px solid rgba(67, 56, 202, 0.35);
    font-weight: 600;
    box-shadow: 0 18px 30px rgba(37, 99, 235, 0.35), 0 0 0 1px rgba(96, 165, 250, 0.12);
    transition: transform 0.24s ease, box-shadow 0.24s ease, filter 0.24s ease;
}
.cta-button:hover {
    background: linear-gradient(135deg, #1d4ed8 0%, #4338ca 100%);
    transform: translateY(-1px);
    box-shadow: 0 22px 40px rgba(37, 99, 235, 0.45), 0 0 0 1px rgba(96, 165, 250, 0.18);
    filter: brightness(1.06);
}
.ats-score-card {
    margin-bottom: 1.2rem;
    padding: 18px 22px;
    border-radius: 24px;
    background: rgba(30, 41, 59, 0.92);
    border: 1px solid rgba(75, 85, 99, 0.24);
    display: inline-flex;
    align-items: center;
    gap: 1rem;
}
.ats-score-number {
    font-size: 3rem;
    font-weight: 800;
    line-height: 1;
}
.ats-score-label {
    color: #cbd5e1;
    font-size: 0.95rem;
}
.ats-score-subtext {
    color: #94a3b8;
    font-size: 0.88rem;
}
.star-badge {
    color: #fbbf24;
    font-size: 1.2rem;
}
.review-card {
    border-radius: 22px;
    padding: 24px;
    background: rgba(15, 23, 42, 0.96);
    box-shadow: 0 22px 44px rgba(0, 0, 0, 0.28);
    border: 1px solid rgba(148, 163, 184, 0.18);
}
.review-card p {
    color: #cbd5e1;
    line-height: 1.7;
}
.review-card img {
    border-radius: 100px;
    width: 64px;
    height: 64px;
    object-fit: cover;
    border: 2px solid #334155;
}
.review-card strong {
    color: #f8fafc;
}
.review-card span {
    color: #94a3b8;
}
.stButton>button,
.stDownloadButton>button,
.stFileUploader button,
.stFileUploader>div>button,
section[data-testid="stFileUploader"] button {
    background-color: #2563eb !important;
    color: #eef2ff !important;
    border: 1px solid rgba(147, 197, 253, 0.45) !important;
    box-shadow: 0 14px 26px rgba(37, 99, 235, 0.18) !important;
    font-weight: 600;
    transition: transform 0.24s ease, box-shadow 0.24s ease, filter 0.24s ease;
}
.stButton>button:hover,
.stDownloadButton>button:hover,
.stFileUploader button:hover,
.stFileUploader>div>button:hover,
section[data-testid="stFileUploader"] button:hover {
    background-color: #1d4ed8 !important;
    color: #f8fafc !important;
    transform: translateY(-1px);
    box-shadow: 0 18px 34px rgba(37, 99, 235, 0.32) !important;
    filter: brightness(1.04);
}
.stTextArea textarea,
.stTextInput input,
.stRadio > div,
.stSelectbox select,
.stTextArea textarea:focus,
.stTextInput input:focus {
    background: rgba(15, 23, 42, 0.98) !important;
    color: #e2e8f0 !important;
    border: 1px solid rgba(148, 163, 184, 0.35) !important;
}
.stTextArea label,
.stTextInput label,
.stRadio label,
.stSelectbox label,
.stFileUploader label,
.stFileUploader span,
section[data-testid="stFileUploader"] label,
section[data-testid="stFileUploader"] span {
    color: #e2e8f0 !important;
}
section[data-testid="stFileUploader"] button,
.stFileUploader>div>button {
    color: #eef2f0 !important;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero-card">
  <h1 style="margin-bottom: 0.35rem; color: #f8fafc;">Career & CV Booster</h1>
  <p style="font-size: 1.18rem; color: #cbd5e1; max-width: 860px; line-height: 1.75;">
    A premium career toolkit for executives and ambitious professionals. Analyze resumes, predict LinkedIn strength, and build premium resume or LinkedIn profiles with trusted business-grade guidance.
  </p>
  <div class="hero-badges">
    <span class="hero-badge">Instant ATS clarity</span>
    <span class="hero-badge">Premium resume insights</span>
    <span class="hero-badge">LinkedIn strength score</span>
  </div>
  <a class="cta-button" href="#main-actions">Start your professional review</a>
</div>
""", unsafe_allow_html=True)

if st.session_state.page == PAGE_HOME:
    st.markdown("<div id='main-actions'></div>", unsafe_allow_html=True)
    col1, col2, col3, col4 = st.columns([1, 1, 1, 1], gap="large")

    with col1:
        st.markdown("""
        <div class="section-card">
        <h3>Resume Analyzer</h3>
        <p>Evaluate your resume structure, spot ATS risks, and get an interpretable score for more competitive applications.</p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Resume Analyzer & ATS score interpreter", key="home_resume"):
            st.session_state.page = PAGE_RESUME

    with col2:
        st.markdown("""
        <div class="section-card">
        <h3>LinkedIn Score Prediction</h3>
        <p>Predict the strength of your LinkedIn profile with data-driven insight for a stronger professional presence.</p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("LinkedIn Profile Score Prediction", key="home_linkedin"):
            st.session_state.page = PAGE_LINKEDIN

    with col3:
        st.markdown("""
        <div class="section-card">
        <h3>Premium Profile Packages</h3>
        <p>Human-verified CV and LinkedIn enhancement for busy professionals, starting at <strong style='color: #facc15;'>INR 499</strong>.</p>
        <p>Get a more compelling career narrative, polished branding, and recruiter-ready profile support.</p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Create Premium Profiles", key="home_premium"):
            st.session_state.page = PAGE_PREMIUM

    with col4:
        st.markdown("""
        <div class="section-card">
        <h3>Job Tracker</h3>
        <p>Keep every application, follow-up, and interview stage organized in one polished career dashboard.</p>
        <p style='margin-top: 12px; color: #a5f3fc; font-weight: 600;'>Services available from <span style='color: #facc15;'>INR 599</span>.</p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Job tracker", key="home_jobtracker"):
            st.session_state.page = PAGE_JOBTRACKER

    st.markdown("---")
    st.subheader("Trusted by professionals")
    st.markdown("""
    <div class="section-card" style="padding: 32px; background: rgba(15, 23, 42, 0.98); display: flex; flex-wrap: wrap; gap: 24px; align-items: stretch;">
      <div style="flex: 1 1 520px; min-width: 280px; max-width: 520px; display: flex; flex-direction: column; gap: 18px;">
        <div style="display: flex; flex-wrap: nowrap; gap: 16px; align-items: center; overflow-x: auto; padding-bottom: 4px;">
          <img src="https://images.unsplash.com/photo-1494790108377-be9c29b29330?w=400&h=400&fit=crop" alt="Professional" style="width: 180px; height: 180px; border-radius: 24px; object-fit: cover; flex-shrink: 0;" />
          <img src="https://images.unsplash.com/photo-1507003211169-0a1dd7228f2d?w=400&h=400&fit=crop" alt="HR Expert" style="width: 180px; height: 180px; border-radius: 24px; object-fit: cover; flex-shrink: 0;" />
          <img src="https://images.unsplash.com/photo-1438761681033-6461ffad8d80?w=400&h=400&fit=crop" alt="Team Lead" style="width: 180px; height: 180px; border-radius: 24px; object-fit: cover; flex-shrink: 0;" />
          <img src="https://images.unsplash.com/photo-1500648767791-00dcc994a43e?w=400&h=400&fit=crop" alt="Career Coach" style="width: 180px; height: 180px; border-radius: 24px; object-fit: cover; flex-shrink: 0;" />
        </div>
        <div style="color: #94a3b8; font-size: 0.95rem; line-height: 1.6;">A professional HR and career expert team reviews every CV, LinkedIn profile, and job tracker entry with human attention.</div>
      </div>
      <div style="flex: 1 1 360px; min-width: 280px; max-width: 100%; color: #cbd5e1; display: flex; flex-direction: column; justify-content: space-between;">
        <div style="min-width: 0;">
          <div style="margin-bottom: 16px; color: #7dd3fc; font-weight: 700; letter-spacing: 0.14em; text-transform: uppercase; font-size: 0.88rem;">Highly rated career guidance</div>
          <h3 style="color: #f8fafc; margin-bottom: 16px; font-size: 1.7rem; line-height: 1.15; max-width: 100%; overflow-wrap: break-word; word-break: break-word;">Trusted by senior professionals, executives, and hiring experts.</h3>
          <p style="line-height: 1.7; color: #cbd5e1; font-size: 1rem; max-width: 100%; overflow-wrap: break-word; word-break: break-word;">We guarantee your CV is optimized to pass ATS screening and positioned to earn real interview calls. With our expert team reviewing every profile, your application becomes a stronger contender.</p>
          <p style="line-height: 1.7; color: #94a3b8; font-size: 0.95rem; max-width: 100%; overflow-wrap: break-word; word-break: break-word; margin-top: 8px;">This is premium, human-reviewed career support for professionals who want results — not just advice.</p>
        </div>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; margin-top: 28px; width: 100%;">
          <div style="padding: 20px; background: rgba(37, 99, 235, 0.08); border-radius: 20px; border: 1px solid rgba(59, 130, 246, 0.18); min-width: 0;">
            <div style="color: #f8fafc; font-weight: 700; font-size: 1.35rem;">4.9/5</div>
            <div style="color: #94a3b8; font-size: 0.95rem; margin-top: 8px;">Average customer rating</div>
          </div>
          <div style="padding: 20px; background: rgba(37, 99, 235, 0.08); border-radius: 20px; border: 1px solid rgba(59, 130, 246, 0.18); min-width: 0;">
            <div style="color: #f8fafc; font-weight: 700; font-size: 1.35rem;">1,200+</div>
            <div style="color: #94a3b8; font-size: 0.95rem; margin-top: 8px;">Career reviews completed</div>
          </div>
          <div style="padding: 20px; background: rgba(37, 99, 235, 0.08); border-radius: 20px; border: 1px solid rgba(59, 130, 246, 0.18); min-width: 0;">
            <div style="color: #f8fafc; font-weight: 700; font-size: 1.35rem;">Premium</div>
            <div style="color: #94a3b8; font-size: 0.95rem; margin-top: 8px;">Human-verified CV + LinkedIn support</div>
          </div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    review_cols = st.columns(3, gap="large")
    with review_cols[0]:
        st.markdown("""
        <div class="review-card">
          <div style="display:flex; align-items:center; gap: 14px; margin-bottom: 18px;">
            <img src="https://images.pexels.com/photos/3771114/pexels-photo-3771114.jpeg?auto=compress&cs=tinysrgb&w=120" alt="Customer" />
            <div>
              <strong>Samantha L.</strong><br />
              <span>VP of Operations</span>
            </div>
          </div>
          <div class="star-badge">★★★★★</div>
          <p style="margin-top: 14px;">"The resume analyzer delivered clear, professional feedback and the ATS insights were spot on. This is a highly trustworthy career platform."</p>
        </div>
        """, unsafe_allow_html=True)

    with review_cols[1]:
        st.markdown("""
        <div class="review-card">
          <div style="display:flex; align-items:center; gap: 14px; margin-bottom: 18px;">
            <img src="https://images.unsplash.com/photo-1506794778202-cad84cf45f1d?auto=format&fit=crop&w=120&q=80" alt="Customer" />
            <div>
              <strong>Avery K.</strong><br />
              <span>Product Leader</span>
            </div>
          </div>
          <div class="star-badge">★★★★★</div>
          <p style="margin-top: 14px;">"The premium profile options helped me create a high-impact LinkedIn summary that clearly reflects my leadership brand."</p>
        </div>
        """, unsafe_allow_html=True)

    with review_cols[2]:
        st.markdown("""
        <div class="review-card">
          <div style="display:flex; align-items:center; gap: 14px; margin-bottom: 18px;">
            <img src="https://images.pexels.com/photos/774909/pexels-photo-774909.jpeg?auto=compress&cs=tinysrgb&w=120" alt="Customer" />
            <div>
              <strong>Jordan P.</strong><br />
              <span>Strategy Consultant</span>
            </div>
          </div>
          <div class="star-badge">★★★★★</div>
          <p style="margin-top: 14px;">"Serious, polished, and easy to use. The website gives the right professional impression for top-level career planning."</p>
        </div>
        """, unsafe_allow_html=True)

    review_cols_2 = st.columns(2, gap="large")
    with review_cols_2[0]:
        st.markdown("""
        <div class="review-card">
          <div style="display:flex; align-items:center; gap: 14px; margin-bottom: 18px;">
            <img src="https://images.pexels.com/photos/1181685/pexels-photo-1181685.jpeg?auto=compress&cs=tinysrgb&w=120" alt="Customer" />
            <div>
              <strong>Priya S.</strong><br />
              <span>Head of Talent</span>
            </div>
          </div>
          <div class="star-badge">★★★★★</div>
          <p style="margin-top: 14px;">"The resume advice is sharp, actionable, and exactly what high-performing professionals need to improve their interview call rate."</p>
        </div>
        """, unsafe_allow_html=True)

    with review_cols_2[1]:
        st.markdown("""
        <div class="review-card">
          <div style="display:flex; align-items:center; gap: 14px; margin-bottom: 18px;">
            <img src="https://images.pexels.com/photos/2379004/pexels-photo-2379004.jpeg?auto=compress&cs=tinysrgb&w=120" alt="Customer" />
            <div>
              <strong>Anil K.</strong><br />
              <span>Founder & CEO</span>
            </div>
          </div>
          <div class="star-badge">★★★★★</div>
          <p style="margin-top: 14px;">"This toolkit made it easy to spot resume gaps and strengthen my executive brand before investor meetings."</p>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.subheader("Why this toolkit?")
    st.write(
        "Our professional toolkit helps job seekers and career builders move from good to exceptional by combining model-driven review with premium profile guidance."
    )
    st.write("- Resume optimization for ATS-friendly formatting")
    st.write("- LinkedIn profile score prediction with actionable insights")
    st.write("- Premium profile creation for a refined career narrative")

else:
    st.write("[← Back to Home](#)", unsafe_allow_html=True)
    if st.button("Back to Home", key="back_button"):
        reset_to_home()

    if st.session_state.page == PAGE_RESUME:
        st.header("Resume Analyzer + ATS Score Interpreter")
        st.write("Upload your resume file first, then upload the job description to receive a professional ATS and recruiter analysis.")
        st.markdown("#### Upload your resume")
        resume_file = st.file_uploader(
            "Upload your resume (PDF, DOCX, or DOC)",
            type=["pdf", "docx", "doc"],
            key="resume_upload",
        )
        st.markdown("#### Upload your job description")
        jd_file = st.file_uploader(
            "Upload the target job description (PDF, DOCX, DOC)",
            type=["pdf", "docx", "doc"],
            key="jd_upload",
        )

        if resume_file is not None:
            st.success(f"Uploaded resume: {resume_file.name}")
        if jd_file is not None:
            st.success(f"Uploaded job description: {jd_file.name}")

        if st.button("Generate ATS & Recruiter Report", key="generate_report"):
            if resume_file is None or jd_file is None:
                st.warning("Please upload both your resume and the job description before generating the report.")
            else:
                with st.spinner("Analyzing files and building your professional report..."):
                    try:
                        resume_text = extract_text_from_uploaded_file(resume_file)
                        jd_text = extract_text_from_uploaded_file(jd_file)
                        if not resume_text.strip():
                            st.error("Could not extract text from the uploaded resume. Please upload a different file.")
                        elif not jd_text.strip():
                            st.error("Could not extract text from the uploaded job description. Please upload a different file.")
                        else:
                            score = calculate_ats_score(resume_text)
                            score_color = "#ef4444" if score < 80 else "#4ade80" if score <= 92 else "#166534"
                            st.markdown(f"""
                            <div class='ats-score-card'>
                                <div class='ats-score-number' style='color: {score_color};'>{score}</div>
                                <div>
                                    <div class='ats-score-label'>ATS Score</div>
                                    <div class='ats-score-subtext'>{score}/100</div>
                                </div>
                            </div>
                            """, unsafe_allow_html=True)
                            report = analyze_resume_with_jd(resume_text, jd_text, score=score)
                            st.markdown("### Professional ATS & Recruiter Report")
                            st.markdown("""
                            <div style='padding: 24px; background: rgba(15, 23, 42, 0.96); border-radius: 22px; border: 1px solid rgba(148, 163, 184, 0.18);'>
                            <pre style='white-space: pre-wrap; color: #e2e8f0; font-size: 0.96rem; line-height: 1.7;'>""" + report + """</pre>
                            </div>
                            """, unsafe_allow_html=True)

                    except Exception as error:
                        st.error(f"Analysis failed: {error}")

    elif st.session_state.page == PAGE_LINKEDIN:
        st.header("LinkedIn Profile Score Prediction")
        st.write("Upload your LinkedIn profile extract PDF and enter your target job title to receive an AI-powered discoverability report.")
        linkedin_file = st.file_uploader(
            "Upload LinkedIn profile export (PDF)",
            type=["pdf"],
            key="linkedin_upload",
        )
        desired_title = st.text_input("Desired job title / target role", "Senior Product Leader")

        if linkedin_file is not None:
            st.success(f"Uploaded profile extract: {linkedin_file.name}")

        if st.button("Analyze LinkedIn Profile", key="analyze_linkedin"):
            if linkedin_file is None or not desired_title.strip():
                st.warning("Please upload your LinkedIn PDF and enter your desired job title before analyzing.")
            else:
                with st.spinner("Analyzing your LinkedIn profile for discoverability and role fit..."):
                    try:
                        profile_text = extract_text_from_uploaded_file(linkedin_file)
                        if not profile_text.strip():
                            st.error("Could not extract text from the uploaded LinkedIn PDF. Please try a different export.")
                        else:
                            report = generate_linkedin_profile_report(profile_text, desired_title)
                            st.markdown("### LinkedIn AI Profile Report")
                            st.markdown("""
                            <div style='padding: 24px; background: rgba(15, 23, 42, 0.96); border-radius: 22px; border: 1px solid rgba(148, 163, 184, 0.18);'>
                            <pre style='white-space: pre-wrap; color: #e2e8f0; font-size: 0.96rem; line-height: 1.75;'>""" + report + """</pre>
                            </div>
                            """, unsafe_allow_html=True)
                            st.markdown(
                                "<div style='margin-top: 18px; padding: 16px; border-radius: 18px; background: rgba(30, 41, 59, 0.85); border: 1px solid rgba(56, 189, 248, 0.18);'>" 
                                "<strong>Note:</strong> This is an AI-generated assessment. For a premium, human-reviewed LinkedIn profile audit and tailored rewriting support, upgrade to our premium services.</div>",
                                unsafe_allow_html=True,
                            )
                    except Exception as error:
                        st.error(f"Analysis failed: {error}")

    elif st.session_state.page == PAGE_PREMIUM:
        st.header("Create Premium Profiles")
        st.write("Choose the premium profile you want to create and get a professional-ready template.")
        st.markdown(
            "<div style='padding: 20px; margin-bottom: 20px; border-radius: 22px; background: linear-gradient(135deg, rgba(37, 99, 235, 0.1), rgba(15, 23, 42, 0.95)); border: 1px solid rgba(96, 165, 250, 0.18);'>"
            "<strong style='color: #7dd3fc; font-size: 1rem;'>Human-verified CV & LinkedIn enhancement</strong> — crafted for senior professionals who want crisp, interview-ready branding. "
            "<span style='color: #facc15; font-weight: 700;'>Charges starting at INR 499 only.</span>"
            "</div>",
            unsafe_allow_html=True,
        )
        premium_option = st.radio(
            "Select premium profile type:",
            ["Create Premium Resume", "Create Premium LinkedIn"],
            index=0,
        )

        if premium_option == "Create Premium Resume":
            st.subheader("Premium Resume Builder")
            resume_goal = st.text_input("Your professional role or target job title", "Senior Product Manager")
            if st.button("Generate Premium Resume", key="generate_premium_resume"):
                if not resume_goal.strip():
                    st.warning("Please enter a target role to create a premium resume.")
                else:
                    st.success("Premium resume outline created")
                    st.write(
                        f"**Target Role:** {resume_goal}"
                    )
                    st.write(
                        "Use this premium structure to shape your resume with crisp summaries, achievement statements, and a strong career narrative."
                    )
                    st.markdown(
                        "**Premium Resume Outline**\n\n"
                        "- Human-verified resume design for stronger recruiter impact.\n"
                        "- Professional Summary: Focus on impact, leadership, and value for the target role.\n"
                        "- Core Strengths: List 5-7 skills tied to results and your industry.\n"
                        "- Experience: Use bullet points with metrics, outcomes, and action verbs.\n"
                        "- Education & Certifications: Highlight relevant training and career-enhancing credentials.\n"
                    )

        else:
            st.subheader("Premium LinkedIn Profile Builder")
            linkedin_focus = st.text_input("Your LinkedIn headline focus", "Driving revenue growth through product innovation")
            if st.button("Generate Premium LinkedIn", key="generate_premium_linkedin"):
                if not linkedin_focus.strip():
                    st.warning("Please enter a headline focus to create a premium LinkedIn profile.")
                else:
                    st.success("Premium LinkedIn profile outline created")
                    st.write(
                        f"**Headline Focus:** {linkedin_focus}"
                    )
                    st.write(
                        "Use this outline to write a compelling LinkedIn headline and summary that positions you as a premium professional."
                    )
                    st.markdown(
                        "**Premium LinkedIn Outline**\n\n"
                        "- Human-verified LinkedIn enhancer designed for discoverability and recruiter appeal.\n"
                        "- Headline: [Role] | [Specialization] | [Impact Statement]\n"
                        "- Summary: Open with your core value, highlight achievements, and close with a call to action.\n"
                        "- Experience: Add concise results-oriented summaries with measurable outcomes.\n"
                        "- Skills: Showcase top strengths aligned with your career goals.\n"
                    )

    elif st.session_state.page == PAGE_JOBTRACKER:
        st.header("Job Tracker")
        st.markdown(
            "<div style='padding: 28px; border-radius: 24px; background: rgba(15, 23, 42, 0.96); border: 1px solid rgba(96, 165, 250, 0.18);'>"
            "<h2 style='color: #f8fafc; margin-bottom: 0.5rem;'>Feeling overwhelmed by your job search?</h2>"
            "<p style='color: #cbd5e1; font-size: 1rem; line-height: 1.75;'>Track opportunities, deadlines, and application statuses in one polished dashboard. Never lose sight of a single lead again.</p>"
            "<p style='color: #7dd3fc; font-size: 1rem; font-weight: 700;'>Services available starting from <span style='color: #facc15;'>INR 599</span>.</p>"
            "<ul style='color: #e2e8f0; margin-top: 18px; font-size: 0.98rem; line-height: 1.7;'>"
            "<li>Job tracking with clear status updates</li>"
            "<li>Application prioritization and follow-up reminders</li>"
            "<li>Career search organization designed for busy professionals</li>"
            "<li>Premium tips for staying focused and landing faster</li>"
            "</ul>"
            "</div>",
            unsafe_allow_html=True,
        )
        st.write("Choose a plan tailored for career momentum and keep your job hunt elegantly organized.")
        st.markdown("- Simple subscription from INR 599 for streamlined job tracking")
        st.markdown("- Upgrade options for premium guided support and follow-up coaching")
        st.markdown("- Designed for ambitious professionals who want clarity, speed, and a better application process")
