# 📄 ATS Resume Analyzer

A Streamlit app that scores a resume for Applicant Tracking Systems (ATS) and tells you exactly how to improve it. The AI review is powered by the [Groq API](https://console.groq.com).

## Features

- Upload a resume as **PDF, DOCX or TXT**
- **ATS score out of 100**, made of two parts:
  - **Rule-based checks (40%)**: contact info, standard section headings, length, bullet formatting, unreadable symbols, action verbs, quantified achievements, and keyword match against a job description
  - **AI review (60%)**: keyword relevance, content quality, impact, and clarity, plus strengths and weaknesses
- Prioritised **improvement suggestions** (High / Medium / Low)
- **Missing keywords** and **bullet-point rewrites**
- Optional **job description** field for role-specific keyword matching
- Downloadable report (`.md`)

## Models

Default model is `openai/gpt-oss-120b`. You can also pick `openai/gpt-oss-20b` (faster) or `llama-3.3-70b-versatile`, or type any model ID into the "Custom model ID" box in the sidebar. Groq changes its model list over time, so check <https://console.groq.com/docs/models> for the current IDs.

## Project structure

```
app.py             # the whole application
requirements.txt   # Python dependencies
README.md          # this file
```

## Run locally

1. Install Python 3.10+.
2. Get a free API key at <https://console.groq.com/keys>.
3. In a terminal:

```bash
pip install -r requirements.txt
streamlit run app.py
```

4. Paste your Groq API key into the sidebar, or set it beforehand:

```bash
# macOS / Linux
export GROQ_API_KEY="gsk_..."
# Windows PowerShell
$env:GROQ_API_KEY="gsk_..."
```

---

## Push to GitHub using the website (no command line)

1. Create an account at <https://github.com> and sign in.
2. Click the **+** at the top right, then **New repository**.
3. Enter a name such as `ats-resume-analyzer`, choose **Public** (needed for the free Streamlit Cloud tier unless you connect a private repo), and click **Create repository**.
4. On the new empty repository page, click the **uploading an existing file** link.
5. Drag in `app.py`, `requirements.txt` and `README.md`. The files must be at the top level of the repository, not inside a folder.
6. Scroll down, leave "Commit directly to the main branch" selected, and click **Commit changes**.

> **Never upload your API key.** The app reads it from Streamlit secrets or the sidebar, so it should never appear in any file you upload.

To update later: open the file on GitHub, click the pencil icon, edit, and click **Commit changes**. Streamlit Cloud redeploys automatically.

## Deploy on Streamlit Community Cloud

1. Go to <https://share.streamlit.io> and sign in with your GitHub account (authorise access if asked).
2. Click **Create app**, then choose **Deploy a public app from GitHub**.
3. Select your repository, set **Branch** to `main`, and **Main file path** to `app.py`.
4. Click **Advanced settings**. In the **Secrets** box paste:

```toml
GROQ_API_KEY = "gsk_your_key_here"
```

5. Click **Deploy**. The first build takes a few minutes while dependencies install. You will get a public `https://your-app.streamlit.app` URL.

To change the secret later: open the app on Streamlit Cloud, then **Settings → Secrets**.

**Who pays for usage?** If you store `GROQ_API_KEY` in Secrets, every visitor uses *your* Groq quota. To make visitors bring their own key, leave the secret out. The app then shows an API key box in the sidebar.

## How the score works

| Category | Max points |
| --- | --- |
| Contact information (email, phone, profile link) | 15 |
| Standard sections (experience, education, skills, summary, projects/certs) | 25 |
| Length and formatting (word count, bullets, clean characters) | 15 |
| Content quality (action verbs, quantified achievements) | 25 |
| Job-description keyword match (only if a job description is pasted) | 20 |

The rule-based score is the percentage of applicable points earned. The final score is `0.4 × rule-based + 0.6 × AI review`.

**Important:** this is an estimate. Real ATS products (Workday, Greenhouse, Lever, Taleo and others) each parse resumes differently and none publish their scoring. Use the result as a guide to making your resume clear, keyword-rich and machine-readable, not as a guarantee.

## Privacy

Resume text is sent to Groq's API for analysis. Don't upload documents you are not comfortable sharing with a third-party service. The app does not store your resume.

## Troubleshooting

| Problem | Fix |
| --- | --- |
| "Almost no text could be extracted" | Your PDF is probably a scanned image. Export a text-based PDF or DOCX from Word / Google Docs. |
| "Groq rejected the API key" | Re-copy the key from console.groq.com/keys and make sure there are no spaces or quotes around it. |
| Rate limit message | Wait a minute, or switch to a different model in the sidebar. |
| Model not found / bad request | The model may have been retired. Check the Groq models page and use the "Custom model ID" box. |
| Streamlit Cloud build fails | Make sure `requirements.txt` is at the top level of the repo and spelled exactly that way. |
