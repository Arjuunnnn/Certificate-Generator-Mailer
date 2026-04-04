# 🎓 Certificate Generator & Mailer

Automates generating personalized certificates for **Winners**, **Participants**, and **Coordinators** from PowerPoint templates — exported as **PDF or PNG** — and optionally emails them via **Zoho Mail**.

---

## ✨ Features

| Feature | Detail |
|---|---|
| 3 certificate categories | Winners, Participants, Coordinators — each with its own template and output folder |
| Dynamic placeholders | `{{ColumnName}}` in slides replaced with Excel values at runtime |
| QR code embedding | Each certificate gets a unique scannable QR linking to a verification URL |
| Unique certificate IDs | Auto-generated IDs (e.g. `IEDC/2026/K9X2`) stamped on every cert and logged |
| PDF export | High-quality via PowerPoint COM automation |
| PNG export | Converted from PDF using `pdf2image` (poppler); temp PDF auto-deleted |
| Dry run mode | Generates only the first record per sheet — safe for testing |
| Resume support | Skips already-generated files if re-run after a crash |
| Email delivery | Zoho SMTP with retry logic (up to 3 attempts per recipient) |
| Parallel emails | Thread pool sends multiple emails simultaneously |
| CSV master log | Every generated certificate's data (including CertID) saved to `certificate_master_log.csv` |

---

## 📂 Project Structure

```
├── certificate_generator.py        # Main script — run this
├── data.xlsx                       # Participant data (3 sheets: Winners, Participants, Coordinators)
├── template_winner.pptx            # PowerPoint template for Winners
├── template_participant.pptx       # PowerPoint template for Participants
├── template_coordinator.pptx       # PowerPoint template for Coordinators
├── email.txt                       # Email subject & body template
├── certificate_master_log.csv      # Auto-generated after each run (CertIDs + all data)
└── certificates/                   # Output folder (auto-created)
    ├── winners/
    │   ├── Alice Smith.pdf         # PDF mode output
    │   └── Alice Smith.png         # PNG mode output
    ├── participants/
    │   └── ...
    └── coordinators/
        └── ...
```

All files (script, templates, data, email template) must be in **the same folder**.

---

## 📊 Excel File — `data.xlsx`

The file must contain **exactly 3 sheets** with these names (case-sensitive):

| Sheet Name | Template Used |
|---|---|
| `Winners` | `template_winner.pptx` |
| `Participants` | `template_participant.pptx` |
| `Coordinators` | `template_coordinator.pptx` |

### Required columns

| Column | Required | Purpose |
|---|---|---|
| `Name` | ✅ Yes | Used in certificate text and as the output filename |
| `Email` | ⚠️ Optional | If present and filled, the certificate is emailed to this address |

### Extra columns

Any additional columns (e.g. `Prize`, `Role`, `Date`, `Event`) can be added freely. They can be referenced in the PowerPoint templates and email body as `{{ColumnName}}` placeholders.

### Example sheet layout

| Name | Email | Prize |
|---|---|---|
| Alice Smith | alice@example.com | Gold |
| Bob Jones | bob@example.com | Silver |

> Rows with an empty `Name` cell are automatically skipped with a warning.

---

## 🖼️ PowerPoint Templates

Place these 3 files in the same folder as the script:

```
template_winner.pptx
template_participant.pptx
template_coordinator.pptx
```

### Text placeholders

Inside each template, add text boxes containing `{{ColumnName}}` tags. These are replaced at runtime with the matching Excel column value for each person.

```
{{Name}}    {{Prize}}    {{Email}}    {{Date}}    {{Role}}    {{CertID}}
```

> `{{CertID}}` is auto-generated per record and available as a placeholder too.

### QR code placeholder

To embed a QR code on the certificate:

1. Add any shape (e.g. a rectangle or image) to the slide where the QR should appear.
2. Right-click the shape → **Edit Alt Text** → set the description to exactly:

```
QR_PLACEHOLDER
```

At generation time, the script finds this shape by its alt text, removes it, and places the QR code image in the exact same position and size.

> The QR links to `VALIDATION_URL + CertID` (configurable in the script settings).

---

## 📝 Email Template — `email.txt`

```
Subject: Congratulations {{Name}}! 🎉

Hi {{Name}},

We are delighted to share your certificate for the event.
Your personalized certificate is attached to this email.

Your certificate ID is: {{CertID}}
You can verify it at: https://yourcollege.edu/verify?id={{CertID}}

Best wishes,
Event Team
```

**Rules:**
- The first line **must** start with `Subject:` — everything after it becomes the email subject.
- Leave exactly one blank line between the subject line and the body.
- All `{{ColumnName}}` placeholders (including `{{CertID}}`) are supported in both subject and body.

---

## ⚙️ Installation

### 1. Install Python dependencies

```bash
pip install -r requirements.txt
```

**`requirements.txt`:**
```
pandas
openpyxl
python-pptx
pdf2image
qrcode[pil]
comtypes
tqdm
```

### 2. Install Poppler (required for PNG export)

**Windows:**
1. Download from: https://github.com/oschwartz10612/poppler-windows/releases
2. Extract the archive.
3. Add the `Library/bin` folder inside it to your system `PATH`.

> Poppler is only needed when exporting as PNG. PDF export works without it.

### 3. Microsoft PowerPoint

Microsoft PowerPoint **must be installed** on the machine running this script. The script drives it via Windows COM automation to convert PPTX to PDF. This is what ensures fonts, gradients, and design elements are rendered accurately.

> ❌ This script does not work on Linux or macOS — PowerPoint COM is Windows-only.

---

## 🔑 Zoho Mail Setup

1. Log in to your Zoho Mail account.
2. Go to **Settings → Security → App Passwords**.
3. Generate an **App-Specific Password** for this script.
4. In `certificate_generator.py`, update:

```python
EMAIL_USER = "your_email@zohomail.in"
EMAIL_PASS = "your_app_password_here"
```

> If you are located **outside India**, change `ZOHO_SMTP_HOST` from `smtp.zoho.in` to `smtp.zoho.com`.

To disable email sending entirely, set:
```python
SEND_EMAILS = False
```

---

## ⚙️ Script Configuration Reference

All settings are at the top of `certificate_generator.py`:

```python
# ── Core ──────────────────────────────────────────────
DATA_FILE       = "data.xlsx"           # Excel file name
OUTPUT_FOLDER   = "certificates"        # Root output folder
NAME_COLUMN     = "Name"               # Column used for filenames

# ── Certificate IDs & QR ──────────────────────────────
CERT_PREFIX     = "IEDC/2026/"          # Prefix for all cert IDs → e.g. IEDC/2026/K9X2
VALIDATION_URL  = "https://yourcollege.edu/verify?id="  # QR base URL
QR_PLACEHOLDER  = "QR_PLACEHOLDER"     # Alt text of the shape to replace with QR

# ── Categories ────────────────────────────────────────
CATEGORIES = [
    ("Winners",      "template_winner.pptx",      "winners"),
    ("Participants", "template_participant.pptx",  "participants"),
    ("Coordinators", "template_coordinator.pptx",  "coordinators"),
]

# ── Email ─────────────────────────────────────────────
SEND_EMAILS         = True
EMAIL_USER          = "your_email@zohomail.in"
EMAIL_PASS          = "your_app_password_here"
EMAIL_TEMPLATE_FILE = "email.txt"
ZOHO_SMTP_HOST      = "smtp.zoho.in"   # Change to smtp.zoho.com outside India
ZOHO_SMTP_PORT      = 465

# ── Performance ───────────────────────────────────────
PNG_DPI             = 300              # PNG export resolution (300 = print quality)
EMAIL_THREADS       = 3               # Parallel email sending threads
MAX_EMAIL_RETRIES   = 3               # Retry attempts per failed email
RETRY_DELAY_SECONDS = 5               # Wait between retries (seconds)
```

---

## ▶️ Running the Script

```bash
python certificate_generator.py
```

The script will prompt two questions:

```
Dry run (first record per sheet only)? (y/n):
```
Enter `y` to test with one record per sheet. Enter `n` to process everyone.

```
Export format:
  1 → PDF
  2 → PNG
Choose (1/2) [default: 1]:
```
Press Enter to use PDF, or type `2` for PNG.

---

## 🔄 How It Works — Step by Step

For every person in every sheet, the script does the following:

```
1. Assign a unique CertID      →  e.g. IEDC/2026/K9X2
2. Generate QR code            →  Links to VALIDATION_URL + CertID
3. Copy template PPTX          →  Replace all {{placeholders}} with Excel values
4. Embed QR image              →  Finds QR_PLACEHOLDER shape, replaces with QR
5. Save filled PPTX            →  Temporary file
6. Convert to PDF              →  Via PowerPoint COM automation (high quality)
7. Delete temporary PPTX       →  Cleanup
8. (PNG mode) Convert to PNG   →  Via pdf2image at 300 DPI, then delete temp PDF
9. Send email (if configured)  →  Attaches PDF or PNG, uses email.txt template
10. Log to CSV                 →  Appends record with CertID to master log
```

---

## 📤 Export Modes

| Mode | Output file | Email attachment | Temp files deleted |
|---|---|---|---|
| PDF | `certificates/<category>/Name.pdf` | `.pdf` | `.pptx` |
| PNG | `certificates/<category>/Name.png` | `.png` | `.pptx` + `.pdf` |

**Multi-slide templates** produce one PNG per slide:
```
Alice Smith_page1.png
Alice Smith_page2.png
```

---

## ⏭️ Resume / Skip Logic

If the script is interrupted (crash, keyboard interrupt, etc.) and re-run, it will **skip any record whose output file already exists** in the output folder. This avoids regenerating certificates that were already completed.

A skipped record is counted separately in the summary report and is **not** re-emailed.

---

## 📊 Output — Summary Report

After every run, the script prints a full summary:

```
─────────────────────────────────────────────
📊 Summary Report
   ✅ Generated  : 120
   ⏭️  Skipped    : 5   (already existed)
   ❌ Failed     : 0
   📧 Emails     : 120 sent / 0 failed

📁 Output structure:
   certificates/winners/
   certificates/participants/
   certificates/coordinators/

📊 CSV    : certificate_master_log.csv
─────────────────────────────────────────────
```

---

## 📋 CSV Master Log — `certificate_master_log.csv`

After every run, the script creates (or overwrites) `certificate_master_log.csv` containing all successfully generated records, including the auto-assigned `CertID` column.

This file is the authoritative record of which certificate was issued to whom. Keep it safe — it is the source of truth for your verification system.

**Example:**

| Name | Email | Prize | CertID |
|---|---|---|---|
| Alice Smith | alice@example.com | Gold | IEDC/2026/K9X2 |
| Bob Jones | bob@example.com | Silver | IEDC/2026/M4R7 |

> The CSV is written once at the end of the run, not incrementally. If the run is interrupted, only records processed before the crash are lost from the CSV — but their output files still exist and will be skipped on re-run.

---

## ⚠️ Error Handling

| Situation | Behaviour |
|---|---|
| Template file not found | That category is skipped; others continue |
| Sheet not found in Excel | That category is skipped; others continue |
| `Name` column missing | That category is skipped; others continue |
| Sheet has no valid records | Warning logged; skipped |
| `Email` column missing | Warning logged; emails skipped for that category |
| Individual record fails | Error logged; script moves to next record |
| Email send fails | Retried up to `MAX_EMAIL_RETRIES` times with delay; failure logged |

---

## 💻 Platform Support

| Platform | Support |
|---|---|
| Windows (with MS PowerPoint) | ✅ Fully supported |
| Linux / macOS | ❌ Not supported — requires PowerPoint COM automation |

---

## 🗒️ Quick Checklist Before Running

- [ ] `data.xlsx` has sheets named exactly `Winners`, `Participants`, `Coordinators`
- [ ] All three `.pptx` template files are in the script folder
- [ ] Each template has `{{Name}}` (and any other placeholders you need)
- [ ] The QR placeholder shape has alt text set to `QR_PLACEHOLDER`
- [ ] `email.txt` exists and starts with `Subject:`
- [ ] `EMAIL_USER` and `EMAIL_PASS` are set (or `SEND_EMAILS = False`)
- [ ] `CERT_PREFIX` and `VALIDATION_URL` are updated to your values
- [ ] Poppler is installed and in PATH (PNG mode only)
- [ ] Microsoft PowerPoint is installed