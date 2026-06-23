# 🔧 Lenovo Elite Engineer Certification Program

A proctored AI-powered assessment platform built with **Streamlit** + custom HTML5/WebGL components.

---

## Project Structure

```
lenovo_cert_app/
├── app.py                  # Main Streamlit app (state machine router)
├── quiz_parser.py          # PDF MCQ extractor (+ hardcoded fallback)
├── data_manager.py         # Excel logging with styled openpyxl
├── certificate_engine.py  # PDF certificate generator (ReportLab / fpdf2)
├── requirements.txt
├── MCQ.pdf                 # ← Place your MCQ PDF here
└── components/
    └── proctoring.html     # Custom HTML5 component:
                            #   • MediaPipe face detection
                            #   • Web Speech API (TTS + STT)
                            #   • Keyboard / visibility / click monitoring
                            #   • Animated SVG avatar "Alex"
```

---

## Quick Start

```bash
# 1. Clone / copy the folder
cd lenovo_cert_app

# 2. Install dependencies
pip install -r requirements.txt

# 3. Place your MCQ PDF alongside app.py (optional – app works without it)
cp /path/to/MCQ.pdf .

# 4. Run
streamlit run app.py
```

Open the URL shown in the terminal (default: http://localhost:8501).

---

## Application Flow

```
Registration → Onboarding → Quiz → Results
```

| Phase | Description |
|---|---|
| **Registration** | Candidate enters name & Employee ID |
| **Onboarding** | Alex (animated avatar) reads briefing via TTS. Camera/mic permissions requested |
| **Quiz** | 11 MCQs; verbal or click answers; live proctoring runs in the right column |
| **Results** | Technical + Integrity scores; downloadable PDF certificate if ≥ 90% |

---

## Proctoring Engine (components/proctoring.html)

All proctoring runs **client-side in the browser** (no server load):

| Event | Penalty |
|---|---|
| Keyboard activity | −10 pts |
| Repeated keyboard | −20 pts |
| Tab / window switch | −25 pts |
| Looking away > 5 s | −10 pts |
| Multiple faces in frame | −30 pts |
| Mobile phone detected | −30 pts |
| Camera lost > 10 s | −20 pts |
| Unsanctioned mouse click | −10 pts |

Score starts at **100**. Results are posted to Streamlit via `window.parent.postMessage`.

---

## 3D / Animated Avatar

The embedded SVG avatar "Alex" (male) is animated with:
- **Blinking eyes** (CSS animation)
- **Mouth open/closed** toggle during TTS speech
- Drop-shadow glow on the avatar container

To swap in a Ready Player Me 3D avatar, replace the `<svg>` block in `proctoring.html`
with an `<iframe src="https://readyplayer.me/..." />` and wire the mouth-open class to
a WebSocket or postMessage from the RPM SDK.

---

## Certificate

Generated as a landscape A4 PDF using **ReportLab** (or fpdf2 as fallback):
- Dark Lenovo-branded design
- Candidate name, employee ID, date
- Technical Score + Integrity Score badges
- Issued only when Technical Score ≥ 90%

---

## Excel Logging

Every completed attempt appends a row to `assessment_results.xlsx`:

| Column | Description |
|---|---|
| Timestamp | ISO date-time |
| Employee ID | From registration |
| Full Name | From registration |
| Technical Score (%) | Correct / total × 100 |
| Integrity Score | Final proctoring score |
| Pass / Fail | ≥ 90% = PASS |
| Certificate Issued | Yes / No |
| Infractions Log | Pipe-separated events |

---

## Deployment (Streamlit Cloud)

1. Push folder to a GitHub repo
2. Go to [share.streamlit.io](https://share.streamlit.io) → New app
3. Set main file path: `app.py`
4. Add secrets / environment variables if needed

For shared `assessment_results.xlsx`, consider mounting a persistent volume
(Streamlit Cloud doesn't persist local files across restarts – use Google Sheets
or an S3-backed path instead).
