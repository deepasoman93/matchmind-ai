# MatchMind AI Candidate and Recruiter App

MatchMind AI provides two separate experiences using the same explainable
matching engine:

- **Candidate:** compare one résumé with a job, review matching and missing
  evidence, add verified information, and download an enhanced DOCX and PDF.
- **Recruiter:** compare one job with multiple résumés, rank applicants, and
  review selective reasons for suitability or unsuitability without modifying
  candidate résumés.

It also includes OCR fallback for scanned PDFs, Google OIDC login, local
account registration/login, and per-user analysis history in SQLite. Each
history option includes its creation time and unique analysis ID, so searches
with the same job title remain distinguishable.

## Candidate workflow

1. Choose **Candidate** on the shared welcome page and sign in.
2. Add the mandatory job title and job description.
3. Upload one PDF, DOCX, TXT or scanned-PDF résumé.
4. Review suitability, matching evidence, missing skills and genuine gaps.
5. Confirm or add optional verified information.
6. Download the enhanced editable DOCX and searchable PDF.

Candidates may continue to enhancement even when the current résumé is not
suitable. Missing qualifications remain clearly reported and are never added
as possessed qualifications.

## Recruiter workflow

1. Choose **Recruiter** on the shared welcome page and sign in.
2. Add the mandatory job title and job description.
3. Upload multiple PDF, DOCX, TXT or scanned-PDF résumés.
4. Review the ranked shortlist, suitability decision, strengths, gaps,
   experience and skill coverage for every candidate.
5. Filter candidates by suitability, experience, overall match or skill
   coverage.
6. Download the filtered table as CSV. When one candidate remains, download
   that résumé directly; when several remain, download their original résumés
   and the filtered table together as one ZIP.

Recruiter mode does not alter or enhance candidate résumés. The step header is
a progress display, and explicit Back and Continue buttons control movement.

## Phone and tablet use

The Candidate and Recruiter experiences are responsive for phone, tablet and
desktop screens. On narrow screens, workflow progress becomes a compact numbered
indicator, cards and controls stack vertically, long candidate tab names are
shortened, and tabs can scroll horizontally. Touch controls remain full-width
and at least 44 pixels high.

The wide recruiter ranking table stays available on desktop. On phones it is
hidden to avoid sideways page scrolling; the ranked candidate cards immediately
below contain the same decision details and remain fully available. Resume upload,
review, enhancement and download actions work from the mobile layout.

## Complete resume enhancement

In Candidate mode, MatchMind creates both an editable DOCX and a selectable,
searchable PDF. The built-in evidence-based enhancer compares the
job requirements with the complete resume and strengthens supported summary, skills, experience and
project wording. It also reports strong matches, remaining genuine gaps and
targeted questions where verified information could improve a later version.
Questions that can be answered completely with Yes or No use a visibly selected
segmented control.
Questions requiring information use a text box. Affirmative confirmations and
typed evidence are inserted into the enhanced résumé rather than merely stored.
Confirmed missing skills are forced into the Technical Skills section even when
the uploaded DOCX uses a table-based layout. Text-based PDF résumés are edited
in place so their styling, page dimensions and page count remain unchanged.
Free-text answers with no clear connection to the job or existing résumé are
reported as skipped instead of being inserted into professional experience.

There are no coloured keyword blocks, no filtered-out sections and no invented
claims. No OpenAI key or external AI account is required. New numbers and new
skills are never added. DOCX uploads are edited in place to retain paragraph
styles and page settings. Text-based PDF uploads use fast structural conversion
to editable DOCX instead of Microsoft Word's slow PDF-import operation. Scanned
PDFs are rebuilt from OCR text without embedding page images. The final PDF is
exported from the tailored DOCX, rendered page by page, and checked for
selectable text, blank pages and page-count consistency.

Eligibility requires at least 50% overall job match, at least 50% required-skill
coverage, and the required experience stated by the job description. Job
Description Match is explanatory; it is not an additional rejection gate. An
explicit 3-5 year range accepts 3, 4 or 5 detected years. A single value such as
"1 year experience" is treated as a minimum, so a candidate with more experience
is not rejected as overqualified.

## Windows setup

Open PowerShell inside the extracted project folder:

```powershell
py -3.10 -m venv resume_ai_env
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\resume_ai_env\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
streamlit run app.py
```

For later runs:

```powershell
cd C:\path\to\AI_Resume_Shortlisting_App
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\resume_ai_env\Scripts\Activate.ps1
streamlit run app.py
```

The transformer downloads on first use. EasyOCR downloads its English OCR
weights only when a scanned document requires OCR.

During matching and resume generation, the interface displays the current
loading step and progress instead of appearing frozen.

Enhancement preparation for up to three candidates runs concurrently. One hidden
Microsoft Word session is then reused only for the fast final DOCX-to-PDF
exports. Uploaded PDFs are never opened in Word, removing the operation that
previously caused 30-minute waits.

## Enable Google sign-in on a local PC

1. In Google Cloud Console, create or select a project and configure the OAuth
   consent screen.
2. Create an OAuth 2.0 Client ID with application type **Web application**.
3. Add this exact authorized redirect URI:

   ```text
   http://localhost:8501/oauth2callback
   ```

4. Copy `.streamlit/secrets.local.example.toml` to
   `.streamlit/secrets.toml` and replace the Client ID, Client Secret and
   cookie-secret placeholders. Never upload or share `secrets.toml`.
5. Stop Streamlit with `Ctrl+C`, then restart it with `streamlit run app.py`.

The Continue with Google button appears automatically. Local username/password
login remains available.

## Enable Google sign-in on Streamlit Cloud

Create a Web application OAuth client and use this redirect URI:

```text
https://YOUR-APP.streamlit.app/oauth2callback
```

Then add these values to the deployed app's Streamlit Secrets:

```toml
[auth]
redirect_uri = "https://YOUR-APP.streamlit.app/oauth2callback"
cookie_secret = "REPLACE_WITH_A_LONG_RANDOM_SECRET"
client_id = "REPLACE_WITH_GOOGLE_CLIENT_ID"
client_secret = "REPLACE_WITH_GOOGLE_CLIENT_SECRET"
server_metadata_url = "https://accounts.google.com/.well-known/openid-configuration"
```

Never commit or share the real Client ID, Client Secret or cookie secret.

## Score composition

- 55% transformer semantic similarity
- 20% normalized required-skill coverage
- 10% exact TF-IDF similarity
- 10% experience requirement
- 5% education requirement

Each résumé's TF-IDF component is calculated against the job description
independently. The same job description and résumé therefore produce the same
match values in Candidate and Recruiter modes, regardless of the other résumés
in a recruiter batch.

These starting weights should be calibrated with recruiter-reviewed examples
before production use.

## Main files

- `Resume_Shortlisting_AI_Development.ipynb`: editable learning flow
- `advanced_engine.py`: scoring and eligibility logic
- `document_parser.py`: PDF, DOCX and TXT extraction
- `database.py`: users, password hashes and analysis history
- `resume_tailor.py`: complete monochrome DOCX and PDF enhancement
- `.streamlit/config.toml`: Cobalt visual theme
- `.streamlit/secrets.local.example.toml`: local Google OIDC template
- `assets/matchmind_star.png`: MatchMind AI logo
- `app.py`: Streamlit application
