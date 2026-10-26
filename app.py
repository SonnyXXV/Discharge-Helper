import datetime
import hashlib
import json
from pathlib import Path

import io

import pandas as pd
import streamlit as st

# Word files for the operation description, primary evaluation and decursus:
#   py -3.13 -m pip install python-docx
try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False
import streamlit.components.v1 as components

# Only the "old panel" option uses this add-on (it is already installed on your computer):
#   py -3.13 -m pip install streamlit-sortables
try:
    from streamlit_sortables import sort_items
    HAS_SORTABLES = True
except ImportError:
    HAS_SORTABLES = False

# ---------------------------------------------------------------
# 1. DATA  (you will edit this part as the program grows)
# ---------------------------------------------------------------

# Starter list only. Later we will replace it with the full
# official ICD-10 list loaded from a file.
DIAGNOSES = {
    "N10": "Acute tubulo-interstitial nephritis (acute pyelonephritis)",
    "N13.3": "Other and unspecified hydronephrosis",
    "N13.0": "Hydronephrosis with ureteropelvic junction obstruction (PUJ stricture)",
    "D29.4": "Scrotal mass – benign tumour / cyst of the scrotum (excision)",
    "N21.0": "Bladder stone (cystolithotripsy)",
    "R33": "Urinary retention – acute (NMC: catheter, sent home)",
    "N20": "Calculus of kidney and ureter (kidney stone disease)",
    "N23": "Unspecified renal colic",
    "N30.0": "Acute cystitis (NMC: sent home)",
    "N32.0": "Bladder-neck obstruction",
    "N35": "Urethral stricture",
    "N39.0": "Urinary tract infection, site not specified",
    "N40": "Benign prostatic hyperplasia",
    "N41.0": "Acute bacterial prostatitis, prostatic abscess (NMC or 4. Urology)",
    "N43": "Hydrocele and spermatocele",
    "N45": "Orchitis / epididymitis, testicular abscess (NMC or 4. Urology)",
    "N47": "Phimosis (4. Urology) / paraphimosis – reduction or dorsumcision (NMC)",
    "R31": "Hematuria (NMC: sent home / 4. Urology: admitted)",
    "C61": "Malignant neoplasm of prostate",
    "C64": "Malignant neoplasm of kidney",
    "C67": "Malignant neoplasm of bladder (TURU / radical cystectomy)",
    "Z43.6": "Nephrostomy – change (48) / fell out, new one by the radiologist (NMC or 4. Urology)",
    "T83": "Complications of genitourinary prosthetic devices, implants and grafts (urinary catheter dysfunction)",
}

# Where the patient is discharged from
DEPARTMENTS = {
    "NMC": "NMC (emergency department)",
    "4": "4. Urology department",
    "48": "48. Urology daily department",
}

# All modifiers that exist.
#   "single"  = choose one option
#   "toggle"  = a tick box (yes/no)
#   "date"    = a date picker
#   "number"  = a number box (with min, max, step); empty until you type
#   "show_if" = only show this modifier when another modifier has this value
#               (or one of several values, written as a list)
#   "default_index" = which option is preselected (without it nothing is preselected)
#   "default"  = for a tick box: True means it starts ticked
# Gleason scores with their ISUP grade group (used by C61, prostate cancer).
GLEASON_OPTIONS = ["3+3=6 (ISUP 1)", "3+4=7 (ISUP 2)", "4+3=7 (ISUP 3)", "4+4=8 (ISUP 4)",
                   "3+5=8 (ISUP 4)", "5+3=8 (ISUP 4)", "4+5=9 (ISUP 5)", "5+4=9 (ISUP 5)",
                   "5+5=10 (ISUP 5)"]

# Co-diagnoses (causes of the nephrostomy) for Z43.6, with ICD-10 codes: label -> Latvian text.
NEPHROSTOMY_CODIAGNOSES = {
    "C53 Cervical cancer": "C53 Dzemdes kakla ļaundabīgs audzējs",
    "C54 Uterine body cancer": "C54 Dzemdes ķermeņa ļaundabīgs audzējs",
    "C56 Ovarian cancer": "C56 Olnīcas ļaundabīgs audzējs",
    "C61 Prostate cancer": "C61 Prostatas ļaundabīgs audzējs",
    "C67 Bladder cancer": "C67 Urīnpūšļa ļaundabīgs audzējs",
    "C65 Renal pelvis cancer": "C65 Nieres bļodiņas ļaundabīgs audzējs",
    "C66 Ureter cancer": "C66 Urīnvada ļaundabīgs audzējs",
    "C18 Colon cancer": "C18 Resnās zarnas ļaundabīgs audzējs",
    "C20 Rectal cancer": "C20 Taisnās zarnas ļaundabīgs audzējs",
    "C77.2 Intra-abdominal lymph node metastases": "C77.2 Intraabdominālo limfmezglu sekundārs ļaundabīgs audzējs",
    "C78.6 Retroperitoneal / peritoneal metastases": "C78.6 Retroperitoneja un vēderplēves sekundārs ļaundabīgs audzējs",
    "N20.1 Ureteral stone": "N20.1 Urīnvada akmens",
    "N20.0 Kidney stone": "N20.0 Nieres akmens",
    "N13.1 Hydronephrosis with ureteral stricture": "N13.1 Hidronefroze ar urīnvada striktūru",
    "N13.6 Pyonephrosis": "N13.6 Pionefroze",
    "N13.3 Hydronephrosis, other": "N13.3 Cita un neprecizēta hidronefroze",
    "Z94.0 Kidney transplant status": "Z94.0 Stāvoklis pēc nieres transplantācijas",
}

MODIFIERS = {
    "Admission": {"type": "single", "options": ["Acute", "Planned"]},
    "Stay": {"type": "single", "options": ["Inpatient", "Outpatient"]},
    "Infection": {"type": "single", "options": ["With infection", "Without infection"]},
    "Side": {"type": "single", "options": ["Right", "Left", "Both"]},
    "Surgery": {"type": "toggle"},
    "Surgery type": {
        "type": "single",
        "options": ["Circumcision", "Dorsumcision"],
        "show_if": {"Surgery": True},
    },
    "Operation date": {"type": "date", "show_if": {"Surgery": True}},
    "Anesthesia": {
        "type": "single",
        "options": ["Spinal", "General", "Local"],
        "show_if": {"Surgery": True},
    },
    "Phimosis degree": {"type": "single", "options": ["Subtotal", "Total"]},
    "Balanoposthitis": {"type": "toggle"},
    "Catheter type": {"type": "single", "options": ["Urethral", "Cystostomy"]},
    "Catheter size (Fr)": {"type": "single", "options": ["12", "14", "16", "18", "20", "22", "24"]},
    "Catheter reason": {
        "type": "single",
        "options": ["Not specified", "Prostate cancer", "BPH"],
        "default_index": 0,
    },
    "Stone size (mm)": {
        "type": "number", "min": 0.0, "max": 100.0, "step": 0.5,
        "show_if": {"Side": ["Right", "Left"]},
    },
    "Right stone size (mm)": {
        "type": "number", "min": 0.0, "max": 100.0, "step": 0.5,
        "show_if": {"Side": "Both"},
    },
    "Left stone size (mm)": {
        "type": "number", "min": 0.0, "max": 100.0, "step": 0.5,
        "show_if": {"Side": "Both"},
    },
    "CRO elevated": {"type": "toggle"},
    "CRO (mg/L)": {
        "type": "number", "min": 0.0, "max": 1000.0, "step": 1.0,
        "show_if": {"CRO elevated": True},
    },
    "Leukocytosis": {"type": "toggle"},
    "Urine infection signs": {"type": "toggle"},

    # ----- N39.0 urinary tract infection -----
    "Admission date": {"type": "date", "show_if": {"_department": "4"}},
    "Discharge date": {"type": "date", "show_if": {"_department": "4"}},
    "Symptoms": {
        "type": "multi", "custom": True,
        "options": ["Lumbar pain", "Flank pain", "Dysuria", "Pollakiuria", "Urgency",
                    "Fever", "Chills", "Hematuria", "General weakness"],
        "custom_help": "Type it in the form that follows 'par', e.g. sliktu dūšu",
    },
    "Infection complexity": {"type": "single", "options": ["Complicated", "Uncomplicated", "Not classified"]},
    "Infection site": {
        "type": "single",
        "options": ["Not specified", "Lower urinary tract", "Upper urinary tract", "Pyelonephritis"],
    },
    "Pyelonephritis side": {
        "type": "single", "options": ["Right", "Left", "Both"],
        "show_if": {"Infection site": "Pyelonephritis"},
    },
    "Infection context": {
        "type": "single",
        "options": ["Catheter present", "Stents present", "Catheter and stents present",
                    "After urological manipulation", "Other / not determined"],
    },
    "Complications": {
        "type": "multi", "custom": True, "optional": True,
        "options": ["Urinary tract obstruction", "Acute kidney injury", "Sepsis"],
        "custom_help": "Type it in Latvian, as it should appear in the diagnosis",
    },
    "Complications status": {
        "type": "single", "options": ["Resolved", "Ongoing"],
        "show_if": {"Complications": "*"},
    },
    "Comorbidities": {"type": "text", "optional": True, "placeholder": "type in Latvian, e.g. cukura diabēts"},
    "Diagnosis certainty": {"type": "single", "options": ["Confirmed", "Possible"], "default_index": 0},

    "Has catheter": {"type": "toggle"},
    "Catheter action": {
        "type": "single", "options": ["Changed", "Inserted", "Left in place"],
        "show_if": {"Has catheter": True},
    },
    "Catheter Fr": {
        "type": "single", "options": ["12", "14", "16", "18", "20"],
        "show_if": {"Has catheter": True},
    },
    "Catheter indication": {
        "type": "single", "default_index": 0,
        "options": ["Not specified", "Urinary retention", "Urine output monitoring",
                    "Postoperative care", "Securing urine drainage"],
        "show_if": {"Has catheter": True},
    },
    "Catheter at discharge": {
        "type": "single", "options": ["In place", "Removed"],
        "show_if": {"Has catheter": True},
    },
    "Urine drainage": {
        "type": "single", "options": ["Free", "Reduced", "None", "Not documented"],
        "show_if": {"Catheter at discharge": "In place"},
    },
    "Voiding after removal": {
        "type": "single",
        "options": ["Voids independently", "Feeling of incomplete emptying",
                    "Increased residual urine", "Urinary retention", "Not tested"],
        "show_if": {"Catheter at discharge": "Removed"},
    },

    "Has stents": {"type": "toggle"},
    "Stent count": {"type": "single", "options": ["1", "2"], "show_if": {"Has stents": True}},
    "Stent side": {"type": "single", "options": ["Right", "Left"], "show_if": {"Stent count": "1"}},
    "Stent age": {
        "type": "text", "optional": True,
        "placeholder": "optional, as it follows 'pirms aptuveni', e.g. 3 mēnešiem (or give the date above)",
        "show_if": {"Has stents": True},
    },
    "Stents at discharge": {
        "type": "single", "options": ["In place", "Removed"],
        "show_if": {"Has stents": True},
    },
    "Stent plan": {
        "type": "single",
        "options": ["On a specific date", "At the outpatient urology review", "Not yet scheduled"],
        "show_if": {"Stents at discharge": "In place"},
    },
    "Stent plan date": {"type": "date", "show_if": {"Stent plan": "On a specific date"}},

    "Antibiotics": {"type": "table", "table": "antibiotics"},
    "Antibiotic allergy status": {
        "type": "single", "options": ["No known allergy", "Documented allergy", "Unknown"],
    },
    "Antibiotic renal adjustment": {
        "type": "single",
        "options": ["Not required", "Dose adjusted", "Interval adjusted", "Requires review"],
    },
    "Antibiotic review": {
        "type": "single",
        "options": ["Reviewed and confirmed", "Pending culture", "Requires prescriber review"],
    },

    "Urine specimen": {
        "type": "single",
        "options": ["Fresh urine specimen", "From catheter sampling port", "Not taken"],
    },
    "Urine culture": {
        "type": "single",
        "options": ["Result not yet known", "Positive", "Negative",
                    "Possibly contaminated sample", "Not performed"],
    },
    "Culture organism": {
        "type": "text", "optional": True, "placeholder": "organism and susceptibility, if known",
        "show_if": {"Urine culture": "Positive"},
    },
    "Culture checked by": {
        "type": "single", "options": ["Family doctor", "Treating urologist", "Hospital department"],
        "show_if": {"Urine culture": "Result not yet known"},
    },
    "Blood culture": {
        "type": "single", "default_index": 0,
        "options": ["Not performed", "Pending", "Negative", "Positive"],
    },
    "Labs": {"type": "table", "table": "labs", "optional": True},

    "Clinical status": {
        "type": "single", "options": ["Improved", "Stable", "Deteriorated", "Incomplete assessment"],
    },
    "Temperature trend": {
        "type": "single", "options": ["Afebrile", "Fever resolved", "Persistent fever", "Recurrent fever"],
    },
    "Pain trend": {
        "type": "single",
        "options": ["Resolved", "Reduced", "Unchanged", "Increased", "Not present initially"],
    },
    "Urinary symptoms trend": {
        "type": "single", "options": ["Resolved", "Improved", "Persistent", "Not present initially"],
    },
    "Inflammatory markers trend": {
        "type": "single", "options": ["Decreasing", "Stable", "Increasing", "Not available"],
    },
    "Renal function trend": {
        "type": "single", "options": ["Stable", "Improved", "Deteriorated", "Not assessed"],
    },
    "Hemodynamic status": {"type": "single", "options": ["Stable", "Unstable", "Not documented"]},
    "Oral intake": {"type": "single", "options": ["Adequate", "Reduced", "Unable to tolerate"]},
    "Treatment tolerance": {
        "type": "single", "options": ["Well tolerated", "Adverse reaction", "Treatment changed"],
    },
    "Discharge readiness": {
        "type": "single", "options": ["Clinician confirms discharge", "Requires reassessment"],
    },

    "Urology follow-up": {"type": "single", "options": ["On a specific date", "Not yet scheduled"]},
    "Urology follow-up date": {"type": "date", "show_if": {"Urology follow-up": "On a specific date"}},
    "Fluid intake": {"type": "text", "placeholder": "e.g. vismaz 2 l dienā"},
    "Warning signs": {
        "type": "multi", "optional": True,
        "options": ["Fever", "Chills", "Flank or lumbar pain", "Urine drainage problems",
                    "Inability to urinate", "Marked hematuria", "General deterioration"],
        "default_selected": ["Fever", "Chills", "Flank or lumbar pain", "Urine drainage problems",
                             "Marked hematuria", "General deterioration"],
    },

    # ----- N20 kidney stones (emergency visit and planned treatment) -----
    "Urostasis grade": {
        "type": "single", "options": ["No urostasis", "I", "I/II", "II", "II/III", "III", "IV"],
    },
    "Pain controlled": {"type": "single", "options": ["Yes", "No"]},
    "Stone location": {
        "type": "single",
        "options": ["Kidney", "Renal pelvis", "Upper ureter", "Middle ureter", "Lower ureter", "Ureteral orifice"],
    },
    "Renal colic": {"type": "toggle", "default": True},
    "Urinary tract infection": {"type": "single", "options": ["No", "Yes"], "default_index": 0},
    "Procedure": {"type": "single", "options": ["URS", "FURS", "PCNL"]},
    "Lithotripsy": {  # what was done with the stone; empty = nothing (e.g. URS with a stent only)
        "type": "multi", "optional": True,
        "options": ["Laser lithotripsy", "Pneumatic lithotripsy", "Basket extraction"],
    },
    "DJ stent placed": {"type": "toggle", "default": True},
    "Procedure date": {"type": "date"},
    "Intraoperative course": {"type": "single", "options": ["Without complications", "With complications"]},
    "Intraoperative complications": {
        "type": "text", "placeholder": "in Latvian, e.g. neliela asiņošana",
        "show_if": {"Intraoperative course": "With complications"},
    },
    "Postoperative course": {"type": "single", "options": ["Uneventful", "Satisfactory", "With complications"]},
    "Postoperative complications": {
        "type": "text", "placeholder": "in Latvian",
        "show_if": {"Postoperative course": "With complications"},
    },
    "Discharge status": {"type": "single", "options": ["Satisfactory", "Good", "Stable"]},
    "Investigations": {"type": "text", "placeholder": "in Latvian, e.g. PAA, bioķīmija, urīna analīze, CT"},
    "Discharged with catheter": {"type": "toggle"},
    "Catheter removal date": {"type": "date", "optional": True, "show_if": {"Discharged with catheter": True}},
    "Stent removal": {
        "type": "single", "options": ["On a specific date", "At the outpatient urology review"],
        "show_if": {"DJ stent placed": True},
    },
    "Stent removal date": {"type": "date", "show_if": {"Stent removal": "On a specific date"}},

    # ----- C67 bladder tumours (TURU) -----
    "Tumour verification": {"type": "multi", "options": ["Cystoscopy", "Ultrasound", "CT"]},
    "Tumour number": {"type": "single", "options": ["Single", "Two", "Three", "Multiple"]},
    "Tumour location": {
        "type": "multi",
        "options": ["Left lateral wall", "Right lateral wall", "Posterior wall", "Anterior wall",
                    "Dome", "Trigone", "Bladder neck"],
    },
    "TURU type": {"type": "single", "options": ["TURU", "Re-TURU"]},
    "Operation details": {
        "type": "text", "optional": True,
        "placeholder": "optional, in Latvian, a full sentence",
    },
    "Intravesical treatment": {
        "type": "single",
        "options": ["Epirubicin single instillation", "Other (type below)", "Not performed"],
        "default_index": 0,
    },
    "Intravesical treatment (other)": {
        "type": "text", "placeholder": "in Latvian, as it follows 'veikta ...'",
        "show_if": {"Intravesical treatment": "Other (type below)"},
    },
    "Catheter removal": {
        "type": "single",
        "options": ["Day 1 after surgery", "Day 2 after surgery", "Day 3 after surgery",
                    "On a specific date", "Discharged with catheter"],
    },
    "Catheter removed on": {"type": "date", "show_if": {"Catheter removal": "On a specific date"}},
    "Planned catheter removal (optional)": {
        "type": "date", "optional": True,
        "show_if": {"Catheter removal": "Discharged with catheter"},
    },
    "Voiding after catheter removal": {
        "type": "single",
        "options": ["Voids independently", "Voids independently, urine clear",
                    "Voids independently, mild hematuria", "Voids independently, mild incontinence",
                    "Urinary retention"],
        "show_if": {"Catheter removal": ["Day 1 after surgery", "Day 2 after surgery",
                                         "Day 3 after surgery", "On a specific date"]},
    },
    "Histology": {"type": "single", "options": ["Pending", "Known"]},
    "Tumour stage": {
        "type": "single",
        "options": ["pTa", "pTis (CIS)", "pT1", "pT2", "Other / not malignant"],
        "show_if": {"Histology": "Known"},
    },
    # ----- N41.0 acute bacterial prostatitis, prostatic abscess -----
    "Symptoms (AP)": {"type": "multi", "optional": True, "default_selected": ["Fever", "Dysuria / frequency"],
                      "options": ["Fever", "Chills", "Dysuria / frequency", "Perineal / suprapubic pain",
                                  "Difficulty voiding / weak stream", "Urinary retention"]},
    "Symptom duration (AP)": {"type": "single", "default_index": 1,
                              "options": ["1 day", "2–3 days", "4–7 days", "About 1 week", "More than a week"]},
    "Context (AP)": {"type": "single", "default_index": 0,
                     "options": ["Spontaneous (community)", "After prostate biopsy", "After catheter / instrumentation"]},
    "Risk factors (AP)": {"type": "multi", "optional": True,
                          "options": ["Recurrent prostatitis", "BPH / outlet obstruction", "Diabetes",
                                      "Immunosuppression", "Indwelling catheter", "Fluoroquinolone in the last 6 months",
                                      "STI risk (young, sexually active)"]},
    "DRE (AP)": {"type": "single", "default_index": 0,
                 "options": ["Enlarged, very tender, warm", "Tender, fluctuating area (abscess?)", "Not performed"]},
    "Labs (AP)": {"type": "single", "default_index": 0,
                  "options": ["Not stated", "Mildly raised CRP", "Raised CRP / leukocytosis", "Raised CRP and creatinine"]},
    "Urine tests (AP)": {"type": "multi", "optional": True,
                         "default_selected": ["Urinalysis: leukocytes / nitrite", "Midstream urine culture"],
                         "options": ["Urinalysis: leukocytes / nitrite", "Midstream urine culture", "Blood culture",
                                     "First-void urine NAAT (STI)"]},
    "Imaging (AP)": {"type": "single", "default_index": 0,
                     "options": ["Not done", "Ultrasound – no residual urine, no abscess", "Ultrasound – residual urine",
                                 "TRUS – prostatitis, no abscess", "TRUS / CT / MRI – prostatic abscess"]},
    "Abscess size (cm) (AP)": {"type": "number", "min": 0.0, "max": 15.0, "step": 0.5, "optional": True,
                               "show_if": {"Imaging (AP)": "TRUS / CT / MRI – prostatic abscess"}},
    "Sepsis signs (AP)": {"type": "toggle"},
    "Retention drainage (AP)": {"type": "single", "default_index": 0,
                                "options": ["No retention", "Urethral catheter (small calibre)", "Suprapubic catheter"]},
    "Drained volume (AP)": {"type": "number", "min": 0.0, "max": 5000.0, "step": 50.0, "optional": True,
                            "show_if": {"Retention drainage (AP)": ["Urethral catheter (small calibre)",
                                                                   "Suprapubic catheter"]}},
    "First IV dose at NMC (AP)": {"type": "text", "optional": True,
                                  "placeholder": "optional, e.g. Sol. Ceftriaxoni 2 g i/v"},
    "IV antibiotic (AP)": {"type": "text", "optional": True,
                           "placeholder": "optional, e.g. Sol. Ceftriaxoni 2 g i/v 1 reizi dienā"},
    "Abscess treatment (AP)": {"type": "single", "default_index": 0,
                               "options": ["Antibiotics only", "TRUS-guided aspiration / drainage",
                                           "Transurethral drainage (deroofing)"],
                               "show_if": {"Imaging (AP)": "TRUS / CT / MRI – prostatic abscess"}},
    "Abscess procedure date (AP)": {"type": "date",
                                    "show_if": {"Abscess treatment (AP)": ["TRUS-guided aspiration / drainage",
                                                                          "Transurethral drainage (deroofing)"]}},
    "Culture result (AP)": {"type": "text", "optional": True, "placeholder": "optional, e.g. E. coli ESBL+"},
    "Ward course (AP)": {"type": "single", "default_index": 0,
                         "options": ["Fever resolved, CRP falling", "Uneventful"]},
    "Catheter at discharge (AP)": {"type": "single", "default_index": 0,
                                   "options": ["Removed – voids, no residual urine", "Discharged with catheter"],
                                   "show_if": {"Retention drainage (AP)": ["Urethral catheter (small calibre)",
                                                                          "Suprapubic catheter"]}},
    "Catheter removal date (AP)": {"type": "date", "optional": True,
                                   "show_if": {"Catheter at discharge (AP)": "Discharged with catheter"}},
    "Control visit (AP)": {"type": "single", "default_index": 0,
                           "options": ["Urologist in 2–3 weeks", "Family doctor in 1 week", "Not needed"]},
    # ----- N45 epididymo-orchitis, testicular / epididymal abscess -----
    "Side (EO)": {"type": "single", "options": ["Right", "Left", "Both"]},
    "Inflammation site (EO)": {"type": "single", "default_index": 1,
                               "options": ["Epididymitis", "Epididymo-orchitis", "Orchitis"]},
    "Onset (EO)": {"type": "single", "default_index": 0, "options": ["Gradual (hours to days)", "Sudden"]},
    "Symptom duration (EO)": {"type": "single", "default_index": 1,
                              "options": ["1 day", "2–3 days", "4–7 days", "More than a week"]},
    "Symptoms (EO)": {"type": "multi", "optional": True, "default_selected": ["Scrotal redness / swelling"],
                      "options": ["Fever", "Dysuria / frequency", "Urethral discharge", "Scrotal redness / swelling"]},
    "Likely source (EO)": {"type": "single", "default_index": 0,
                           "options": ["Enteric / urinary (UTI, obstruction, catheter, instrumentation)",
                                       "Sexually transmitted", "Both possible", "Mumps (viral)", "Not clear"]},
    "Risk factors (EO)": {"type": "multi", "optional": True,
                          "options": ["Indwelling catheter / epicystostomy", "Recent urological procedure",
                                      "Bladder outlet obstruction (BPH)", "Diabetes", "Immunosuppression",
                                      "Insertive anal intercourse"]},
    "Scrotal US (EO)": {"type": "multi", "optional": True,
                        "default_selected": ["Epididymitis (hyperaemia)", "Normal testicular blood flow"],
                        "options": ["Epididymitis (hyperaemia)", "Orchitis", "Reactive hydrocele",
                                    "Abscess of the epididymis", "Abscess of the testis", "Pyocele",
                                    "Testicular infarction / ischaemia", "Normal testicular blood flow"]},
    "Labs (EO)": {"type": "single", "default_index": 0,
                  "options": ["Not stated", "No inflammatory changes", "Raised CRP / leukocytosis",
                              "Raised CRP and creatinine"]},
    "Urine tests (EO)": {"type": "multi", "optional": True, "default_selected": ["Midstream urine culture"],
                         "options": ["Midstream urine culture", "First-void urine NAAT (chlamydia, gonorrhoea, "
                                                                "M. genitalium)", "Leukocyturia"]},
    "Sepsis signs (EO)": {"type": "toggle"},
    "Ceftriaxone given at NMC": {"type": "toggle"},
    "Treatment (EO)": {"type": "single", "default_index": 0, "options": ["Conservative (IV antibiotics)", "Surgery"]},
    "IV antibiotic (EO)": {"type": "text", "optional": True,
                           "placeholder": "optional, e.g. Sol. Ceftriaxoni 2 g i/v 1 reizi dienā"},
    "After failed conservative therapy": {"type": "toggle", "show_if": {"Treatment (EO)": "Surgery"}},
    "Operation (EO)": {"type": "single",
                       "options": ["Orchiectomy (scrotal)", "Inguinal orchifuniculectomy", "Epididymectomy",
                                   "Abscess drainage / revision", "Scrotal exploration"],
                       "show_if": {"Treatment (EO)": "Surgery"}},
    "Anesthesia (EO)": {"type": "single", "default_index": 0, "options": ["Spinal", "General"],
                        "show_if": {"Treatment (EO)": "Surgery"}},
    "Postoperative course (EO)": {"type": "single", "default_index": 0, "options": ["Uneventful", "With complications"],
                                  "show_if": {"Treatment (EO)": "Surgery"}},
    "Postoperative complications (EO)": {"type": "text", "placeholder": "in Latvian",
                                         "show_if": {"Postoperative course (EO)": "With complications"}},
    "Drain (EO)": {"type": "single", "default_index": 0,
                   "options": ["Removed in the ward", "No drain", "Discharged with drain"],
                   "show_if": {"Treatment (EO)": "Surgery"}},
    "Histology (EO)": {"type": "single", "default_index": 0, "options": ["Pending", "Known", "Not sent"],
                       "show_if": {"Treatment (EO)": "Surgery"}},
    "Histology result (EO)": {"type": "text", "placeholder": "in Latvian, e.g. strutains orhīts ar abscesu",
                              "show_if": {"Histology (EO)": "Known"}},
    "Culture result (EO)": {"type": "text", "optional": True,
                            "placeholder": "optional, e.g. Staphylococcus aureus MS"},
    "Control visit (EO)": {"type": "single", "default_index": 0,
                           "options": ["Urologist in 2 weeks", "Family doctor in 1 week", "4. Urology (sutures)",
                                       "Not needed"]},
    # ----- C67: which operation, NMIBC risk factors (EAU), radical cystectomy -----
    "C67 operation": {"type": "single", "default_index": 0, "options": ["TURU / re-TURU", "Radical cystectomy"],
                      "show_if": {"_department": "4"}},
    "Smoker (C67)": {"type": "toggle"},
    "Recurrent tumour": {"type": "toggle"},
    "Largest tumour over 3 cm": {"type": "toggle"},
    "Age over 70 (C67)": {"type": "toggle"},
    "Detrusor muscle in specimen": {"type": "single", "default_index": 0, "options": ["Present", "Absent", "Not stated"],
                                    "show_if": {"Histology": "Known"}},
    "Variant histology / LVI": {"type": "toggle", "show_if": {"Histology": "Known"}},
    "Adjuvant plan (NMIBC)": {"type": "single", "default_index": 0,
                              "options": ["Decide after histology / council", "No further instillations",
                                          "Intravesical chemotherapy course", "BCG induction + maintenance",
                                          "Discuss radical cystectomy"]},
    "Indication (RC)": {"type": "single", "default_index": 0,
                        "options": ["Muscle-invasive bladder cancer", "Very high-risk / BCG-unresponsive NMIBC",
                                    "Palliative (bleeding, pain, fistula)"]},
    "Clinical T (RC)": {"type": "single", "default_index": 1, "options": ["cT1", "cT2", "cT3", "cT4a", "cT4b"]},
    "Clinical N (RC)": {"type": "single", "default_index": 0, "options": ["cN0", "cN1", "cN2", "cN3"]},
    "Clinical M (RC)": {"type": "single", "default_index": 0, "options": ["cM0", "cM1"]},
    "Previous treatment (RC)": {"type": "multi", "optional": True, "default_selected": ["TURU"],
                                "options": ["TURU", "Re-TURU", "BCG instillations", "Intravesical chemotherapy",
                                            "Neoadjuvant cisplatin chemotherapy", "Perioperative durvalumab (NIAGARA)",
                                            "Pelvic radiotherapy", "Nephrostomy right", "Nephrostomy left"]},
    "Persistent hematuria (RC)": {"type": "toggle"},
    "Anaemia (RC)": {"type": "toggle"},
    "Preoperative transfusion (RC, units)": {"type": "number", "min": 0.0, "max": 20.0, "step": 1.0, "optional": True},
    "Cystectomy type (man)": {"type": "single", "default_index": 0,
                              "options": ["Radical cystoprostatectomy", "Nerve-sparing cystoprostatectomy",
                                          "Palliative cystectomy"],
                              "show_if": {"_sex": "Man"}},
    "Cystectomy type (woman)": {"type": "single", "default_index": 0,
                                "options": ["With hysterectomy, adnexa and anterior vaginal wall",
                                            "With anterior vaginal wall (no uterus)",
                                            "Organ-sparing (uterus and vagina preserved)", "Palliative cystectomy"],
                                "show_if": {"_sex": "Woman"}},
    "Urethrectomy (RC)": {"type": "toggle"},
    "Approach (RC)": {"type": "single", "default_index": 0, "options": ["Open", "Robot-assisted", "Laparoscopic"]},
    "Lymph node dissection (RC)": {"type": "single", "default_index": 0,
                                   "options": ["Standard bilateral (obturator, internal and external iliac)",
                                               "Extended bilateral", "Not performed"]},
    "Urinary diversion": {"type": "single",
                          "options": ["Ileal conduit", "Orthotopic neobladder", "Ureterocutaneostomy",
                                      "No diversion – nephrostomies"]},
    "Anesthesia (RC)": {"type": "single", "default_index": 0, "options": ["Epidural + general", "General"]},
    "Postoperative transfusion (RC, units)": {"type": "number", "min": 0.0, "max": 20.0, "step": 1.0,
                                              "optional": True},
    "Bowel recovery (RC)": {"type": "single", "default_index": 0,
                            "options": ["Peristalsis and stool recovered", "Ileus, resolved conservatively"]},
    "Drain (RC)": {"type": "single", "default_index": 0,
                   "options": ["Removed in the ward", "Discharged with the drain"]},
    "Drain removed on day (RC)": {"type": "number", "min": 1.0, "max": 60.0, "step": 1.0, "optional": True,
                                  "show_if": {"Drain (RC)": "Removed in the ward"}},
    "Ureteral stents (RC)": {"type": "single", "default_index": 0,
                             "options": ["Discharged with stents", "Removed in the ward", "No stents"],
                             "show_if": {"Urinary diversion": ["Ileal conduit", "Orthotopic neobladder",
                                                               "Ureterocutaneostomy"]}},
    "Stent removal date (RC)": {"type": "date", "optional": True,
                                "show_if": {"Ureteral stents (RC)": "Discharged with stents"}},
    "Stent removal time (RC)": {"type": "text", "optional": True, "placeholder": "optional, e.g. 09:00",
                                "show_if": {"Ureteral stents (RC)": "Discharged with stents"}},
    "Neobladder catheter removal (RC)": {"type": "date", "optional": True,
                                         "show_if": {"Urinary diversion": "Orthotopic neobladder"}},
    "Histology (RC)": {"type": "single", "default_index": 0, "options": ["Pending", "Known"]},
    "Histology expected (RC)": {"type": "date", "optional": True, "show_if": {"Histology (RC)": "Pending"}},
    "pT (RC)": {"type": "single",
                "options": ["pT0", "pTis", "pTa", "pT1", "pT2a", "pT2b", "pT3a", "pT3b", "pT4a", "pT4b"],
                "show_if": {"Histology (RC)": "Known"}},
    "pN (RC)": {"type": "single", "default_index": 0, "options": ["pN0", "pN1", "pN2", "pN3", "pNx"],
                "show_if": {"Histology (RC)": "Known"}},
    "Nodes positive / removed (RC)": {"type": "text", "optional": True, "placeholder": "optional, e.g. 0/18",
                                      "show_if": {"Histology (RC)": "Known"}},
    "Margins (RC)": {"type": "single", "default_index": 0, "options": ["R0", "R1"],
                     "show_if": {"Histology (RC)": "Known"}},
    "Histology details (RC)": {"type": "text", "optional": True,
                               "placeholder": "optional, in Latvian, e.g. uroteliāla karcinoma HG, LVI+",
                               "show_if": {"Histology (RC)": "Known"}},
    "Diabetes diet (RC)": {"type": "toggle"},
    "Urology visit (RC)": {"type": "single", "default_index": 1, "options": ["4–6 weeks", "3 months", "6 months"]},
    "Tumour grade": {
        "type": "single", "options": ["Low grade (LG)", "High grade (HG)", "Not stated"],
        "show_if": {"Tumour stage": ["pTa", "pT1", "pT2"]},
    },
    "Concomitant CIS": {"type": "toggle", "show_if": {"Tumour stage": ["pTa", "pT1", "pT2"]}},
    "N stage": {
        "type": "single", "default_index": 0,
        "options": ["Not assessed", "cN0", "cN1", "cN2", "cN3"],
        "show_if": {"Tumour stage": ["pTa", "pTis (CIS)", "pT1", "pT2"]},
    },
    "M stage": {
        "type": "single", "default_index": 0,
        "options": ["Not assessed", "cM0", "cM1a", "cM1b"],
        "show_if": {"Tumour stage": ["pTa", "pTis (CIS)", "pT1", "pT2"]},
    },
    "Histology details": {
        "type": "text", "optional": True,
        "placeholder": "optional, in Latvian, e.g. uroteliāla karcinoma (required for 'Other')",
        "show_if": {"Histology": "Known"},
    },
    "Discharge paper": {"type": "single", "options": ["Temporary", "Final"]},
    "Activity restriction": {"type": "single", "options": ["1 week", "2 weeks", "3 weeks", "4 weeks", "1 month",
                                                           "6 weeks"]},
    "Urology visit": {"type": "single", "options": ["2 weeks", "4 weeks", "6 weeks", "3 months"]},
    "Oncology council": {
        "type": "single",
        "options": ["Will be discussed after histology", "Already discussed", "Not applicable"],
    },
    "Oncology council date": {"type": "date", "show_if": {"Oncology council": "Already discussed"}},

    # ----- C61 prostate cancer (radical prostatectomy, TURP, NMC visit) -----
    # "_department" is filled in by the program: it is where the patient is discharged from.
    "Operation": {
        "type": "single", "options": ["Radical prostatectomy", "TURP"],
        "show_if": {"_department": "4"},
    },
    "PSA (ng/ml)": {"type": "number", "min": 0.0, "max": 10000.0, "step": 0.1, "optional": True},
    "Clinical T stage": {
        "type": "single", "default_index": 0,
        "options": ["Not stated", "cT1c", "cT2a", "cT2b", "cT2c", "cT3a", "cT3b", "cT4"],
    },
    "Prostate N stage": {"type": "single", "default_index": 0, "options": ["Not assessed", "cN0", "cN1"]},
    "Prostate M stage": {
        "type": "single", "default_index": 0,
        "options": ["Not assessed", "cM0", "cM1a", "cM1b", "cM1c"],
    },
    "Biopsy Gleason score": {"type": "single", "default_index": 0, "options": ["Not known"] + GLEASON_OPTIONS},
    "RP approach": {"type": "single", "options": ["Open retropubic", "Laparoscopic", "Robot-assisted"]},
    "Lymph node dissection": {"type": "toggle"},
    "TURP anesthesia": {"type": "single", "options": ["Spinal", "General"]},
    "Bladder irrigation after surgery": {"type": "toggle", "default": True},
    "Pathological T stage": {
        "type": "single", "options": ["pT2", "pT3a", "pT3b", "pT4"],
        "show_if": {"Histology": "Known", "Operation": "Radical prostatectomy"},
    },
    "Pathological N stage": {
        "type": "single", "options": ["pN0", "pN1", "pNx (no lymph node dissection)"],
        "show_if": {"Histology": "Known", "Operation": "Radical prostatectomy"},
    },
    "Surgical margin": {
        "type": "single", "options": ["R0", "R1"],
        "show_if": {"Histology": "Known", "Operation": "Radical prostatectomy"},
    },
    "Final Gleason score": {
        "type": "single", "options": GLEASON_OPTIONS + ["Not stated"],
        "show_if": {"Histology": "Known", "Operation": "Radical prostatectomy"},
    },
    "Prostate histology details": {
        "type": "text", "optional": True,
        "placeholder": "in Latvian; optional after prostatectomy, required after TURP",
        "show_if": {"Histology": "Known"},
    },
    "NMC reason": {"type": "single", "options": ["Urinary retention", "Hematuria", "Pain"]},
    "NMC catheter Fr": {
        "type": "single", "options": ["12", "14", "16", "18", "20", "22"],
        "show_if": {"NMC reason": "Urinary retention"},
    },
    "Drained urine (ml)": {
        "type": "number", "min": 0.0, "max": 5000.0, "step": 50.0, "optional": True,
        "show_if": {"NMC reason": "Urinary retention"},
    },
    "Hematuria management": {
        "type": "single", "options": ["Observation", "Three-way catheter and bladder irrigation"],
        "show_if": {"NMC reason": "Hematuria"},
    },
    "Hemodynamically stable": {"type": "single", "options": ["Yes", "No"], "show_if": {"NMC reason": "Hematuria"}},
    "Urine clearing": {"type": "single", "options": ["Yes", "No"], "show_if": {"NMC reason": "Hematuria"}},
    "Pain location": {
        "type": "text", "placeholder": "in Latvian, as it follows 'sāpēm', e.g. mugurā",
        "show_if": {"NMC reason": "Pain"},
    },
    "Fever at NMC": {"type": "toggle"},
    "Urology consultation": {"type": "single", "options": ["Within 1 week", "In 1–2 weeks", "In 4 weeks"]},

    # ----- C64 kidney cancer (radical or partial nephrectomy, NMC visit) -----
    "Kidney operation": {
        "type": "single", "options": ["Radical nephrectomy", "Partial nephrectomy"],
        "show_if": {"_department": "4"},
    },
    "Kidney side": {"type": "single", "options": ["Right", "Left"]},
    "Tumour size (cm)": {"type": "number", "min": 0.0, "max": 40.0, "step": 0.1, "optional": True},
    "RCC clinical T stage": {
        "type": "single", "default_index": 0,
        "options": ["Not stated", "cT1a", "cT1b", "cT2a", "cT2b", "cT3a", "cT3b", "cT3c", "cT4"],
    },
    "RCC N stage": {"type": "single", "default_index": 0, "options": ["Not assessed", "cN0", "cN1"]},
    "RCC M stage": {"type": "single", "default_index": 0, "options": ["Not assessed", "cM0", "cM1"]},
    "Kidney approach": {"type": "single", "options": ["Open", "Laparoscopic", "Robot-assisted"]},
    "Adrenalectomy": {"type": "toggle"},
    "Renal function at discharge": {"type": "single", "options": ["Stable", "Mildly impaired, stable", "Deteriorating"]},
    "Creatinine at discharge (µmol/l)": {"type": "number", "min": 0.0, "max": 3000.0, "step": 1.0, "optional": True},
    "eGFR at discharge": {"type": "number", "min": 0.0, "max": 200.0, "step": 1.0, "optional": True},
    "RCC subtype": {
        "type": "single", "options": ["Clear cell", "Papillary", "Chromophobe", "Other RCC", "Benign"],
        "show_if": {"Histology": "Known"},
    },
    "RCC pT stage": {
        "type": "single", "options": ["pT1a", "pT1b", "pT2a", "pT2b", "pT3a", "pT3b", "pT3c", "pT4"],
        "show_if": {"RCC subtype": ["Clear cell", "Papillary", "Chromophobe", "Other RCC"]},
    },
    "RCC pN stage": {"type": "single", "options": ["pNx", "pN0", "pN1"], "show_if": {"RCC subtype": ["Clear cell", "Papillary", "Chromophobe", "Other RCC"]}},
    "RCC margin": {"type": "single", "options": ["R0", "R1"], "show_if": {"RCC subtype": ["Clear cell", "Papillary", "Chromophobe", "Other RCC"]}},
    "RCC ISUP grade": {
        "type": "single", "options": ["1", "2", "3", "4", "Not graded"],
        "show_if": {"RCC subtype": ["Clear cell", "Papillary", "Chromophobe", "Other RCC"]},
    },
    "Kidney histology details": {
        "type": "text", "optional": True,
        "placeholder": "in Latvian; required for 'Other RCC' and 'Benign', e.g. onkocitoma",
        "show_if": {"Histology": "Known"},
    },
    "Follow-up risk group": {
        "type": "single", "options": ["Low", "Intermediate", "High"],
        "show_if": {"RCC subtype": ["Clear cell", "Papillary", "Chromophobe", "Other RCC"]},
    },

    # ----- added in the big update (N10 NMC, N47, T83, N20 drainage, N39 modules) -----
    "NMC imaging": {"type": "single", "options": ["Ultrasound", "CT", "None"]},
    "Urostasis on imaging": {
        "type": "single", "options": ["No", "Yes"], "show_if": {"NMC imaging": ["Ultrasound", "CT"]},
    },
    "Sepsis signs": {"type": "single", "options": ["No", "Yes"]},
    "Able to take oral medication": {"type": "single", "options": ["Yes", "No"]},
    "Pregnant": {"type": "single", "options": ["No", "Yes"], "show_if": {"_sex": "Woman"}},
    "Initial IV antibiotic at NMC": {"type": "toggle"},
    "Conservative reason": {
        "type": "single", "options": ["Balanoposthitis", "Paraphimosis"], "show_if": {"Surgery": False},
    },
    "Frenulum breve": {"type": "toggle"},
    "Frenuloplasty": {"type": "toggle", "show_if": {"Frenulum breve": True, "Surgery": True}},
    # ----- N47 paraphimosis at NMC -----
    "Paraphimosis duration": {"type": "single", "default_index": 1,
                              "options": ["A few hours", "1 day", "2 days", "More than 2 days"]},
    "Paraphimosis cause": {"type": "multi", "optional": True,
                           "options": ["After intercourse / masturbation", "After urinary catheterisation",
                                       "After examination / hygiene", "Not known"]},
    "Glans appearance": {"type": "single", "default_index": 0,
                         "options": ["Pink, viable", "Dusky / cyanotic", "Necrotic areas"]},
    "Paraphimosis treatment": {"type": "single",
                               "options": ["Manual reduction", "Dorsumcision (manual reduction failed)"]},
    "Before reduction (PF)": {"type": "multi", "optional": True, "default_selected": ["Local anaesthesia"],
                              "options": ["Local anaesthesia", "Manual compression of the oedema",
                                          "Puncture of the oedema (Dundee)"]},
    "Dorsumcision sutures": {"type": "single", "default_index": 0,
                             "options": ["4-0 Vicryl interrupted", "4-0 Vicryl Rapide interrupted", "No sutures"],
                             "show_if": {"Paraphimosis treatment": "Dorsumcision (manual reduction failed)"}},
    "Control visit (PF)": {"type": "single", "default_index": 0,
                           "options": ["4. Urology department", "Urologist outpatient", "Family doctor"]},
    "Control visit date (PF)": {"type": "date", "optional": True},
    "Control visit time (PF)": {"type": "text", "optional": True, "placeholder": "optional, e.g. 12.00"},
    "Suture removal at the visit": {"type": "toggle", "default": True,
                                    "show_if": {"Paraphimosis treatment": "Dorsumcision (manual reduction failed)"}},
    "Plan circumcision": {"type": "toggle", "default": True},
    "Urinary retention (diagnosis)": {"type": "toggle"},
    "Drained volume (ml)": {"type": "number", "min": 0.0, "max": 5000.0, "step": 50.0},
    "Urine appearance at change": {
        "type": "single", "options": ["Yellow, clear", "Cloudy", "Blood-tinged", "With clots"],
    },
    "Urosepsis": {"type": "toggle"},
    "Drainage": {"type": "single", "options": ["DJ stent", "Percutaneous nephrostomy"]},
    "Drainage date": {"type": "date"},
    "Purulent urine on drainage": {"type": "toggle"},
    "Definitive stone treatment": {
        "type": "single", "options": ["URS/FURS", "PCNL", "To be decided by the urologist"],
    },
    "Definitive treatment date": {
        "type": "date", "optional": True, "show_if": {"Definitive stone treatment": ["URS/FURS", "PCNL"]},
    },
    "Symptom duration": {
        "type": "single", "default_index": 0,
        "options": ["Not documented", "Hours", "1–2 days", "3–7 days", "Over a week"],
    },
    "Symptom onset": {
        "type": "single", "default_index": 0,
        "options": ["Not documented", "Sudden", "Gradual", "After catheter change", "After stent insertion",
                    "After urological manipulation"],
    },
    "Fever status": {
        "type": "single", "default_index": 0, "options": ["Not documented", "Febrile", "Subfebrile", "Afebrile"],
    },
    "Relevant history": {
        "type": "multi", "optional": True, "custom": True,
        "options": ["Previous UTIs", "Recurrent UTIs", "Nephrolithiasis", "Urinary tract obstruction",
                    "Urological surgery", "Diabetes", "Immunosuppression"],
        "custom_help": "Type it in Latvian, as it should appear after 'Anamnēzē:'",
    },
    "Catheter route": {
        "type": "single", "default_index": 0,
        "options": ["Urethral (Foley)", "Suprapubic (epicistostoma)"],
        "show_if": {"Has catheter": True},
    },
    "Stent type": {"type": "single", "default_index": 0, "options": ["Not specified", "DJ"], "show_if": {"Has stents": True}},
    "Stent indication": {
        "type": "single", "default_index": 0,
        "options": ["Not specified", "Obstruction", "Ureteral stone", "Postoperative drainage"],
        "show_if": {"Has stents": True},
    },
    "Stent insertion date": {"type": "date", "optional": True, "show_if": {"Has stents": True}},
    "Urine appearance": {
        "type": "single", "default_index": 0,
        "options": ["Not documented", "Clear", "Cloudy", "Blood-tinged", "With clots"],
    },
    "Imaging": {"type": "single", "default_index": 0, "options": ["Not done", "Ultrasound", "CT"]},
    "Hydronephrosis": {
        "type": "single", "options": ["None", "Right", "Left", "Both"],
        "show_if": {"Imaging": ["Ultrasound", "CT"]},
    },
    "Obstruction drained": {
        "type": "single", "options": ["Yes", "No"], "show_if": {"Hydronephrosis": ["Right", "Left", "Both"]},
    },
    "Chronic medication changes": {
        "type": "text", "optional": True,
        "placeholder": "optional, in Latvian, e.g. Metformin nelietot līdz kontrolei",
    },

    # ----- N35 urethral stricture -----
    "Stricture location": {
        "type": "single",
        "options": ["Meatus", "Penile urethra", "Bulbar urethra", "Membranous (posterior) urethra", "Not specified"],
    },
    "Stricture length (cm)": {"type": "number", "min": 0.0, "max": 25.0, "step": 0.5, "optional": True},
    "Stricture cause": {
        "type": "single", "default_index": 0,
        "options": ["Unknown", "Iatrogenic (catheter / transurethral surgery)", "Lichen sclerosus", "Trauma", "Infection"],
    },
    "Previous stricture treatment": {"type": "toggle"},
    "Stricture compensation": {"type": "single", "default_index": 0,
                               "options": ["Not stated", "Compensated", "Subcompensated", "Decompensated"]},
    "Stricture verified by": {"type": "multi", "optional": True, "options": ["Cystoscopy", "Urethrography"],
                              "default_selected": ["Cystoscopy"]},
    "Urethrotomy instrument": {"type": "single", "options": ["Cold knife (Sachse)", "Laser"],
                               "show_if": {"Stricture operation": "Internal urethrotomy (DVIU)"}},
    "Also performed (N35)": {"type": "text", "optional": True,
                             "placeholder": "optional, in Latvian, e.g. urīnpūšļa konkrementa evakuācija"},
    "Rest after endoscopic treatment": {"type": "single", "default_index": 0,
                                        "options": ["1 week", "2 weeks", "4 weeks", "1 month"],
                                        "show_if": {"Stricture operation": ["Urethral dilation",
                                                                            "Internal urethrotomy (DVIU)"]}},
    "Catheter removal place (N35)": {"type": "single", "default_index": 0,
                                     "options": ["PSKUS outpatient clinic", "Family doctor's practice", "Not stated"],
                                     "show_if": {"Catheter at discharge (N35)": "In place"}},
    "Catheter removal time (N35)": {"type": "text", "optional": True, "placeholder": "optional, e.g. 12:45",
                                    "show_if": {"Catheter at discharge (N35)": "In place"}},
    "Stricture operation": {
        "type": "single",
        "options": ["Meatotomy", "Meatoplasty", "Urethral dilation", "Internal urethrotomy (DVIU)",
                    "Urethroplasty – excision and primary anastomosis", "Urethroplasty – buccal mucosa graft"],
    },
    "N35 anesthesia": {"type": "single", "options": ["Spinal", "General", "Local"]},
    "Tissue for histology": {"type": "single", "default_index": 0, "options": ["Not sent", "Pending", "Known"]},
    "Histology result (N35)": {
        "type": "text", "placeholder": "in Latvian, e.g. lichen sclerosus", "show_if": {"Tissue for histology": "Known"},
    },
    "Catheter Fr (N35)": {"type": "single", "options": ["12", "14", "16", "18", "20", "22"]},
    "Catheter at discharge (N35)": {"type": "single", "options": ["In place", "Removed"]},
    "Planned catheter removal (N35)": {
        "type": "date", "optional": True, "show_if": {"Catheter at discharge (N35)": "In place"},
    },
    "Voiding after removal (N35)": {
        "type": "single", "options": ["Good stream", "Weak stream", "Urinary retention"],
        "show_if": {"Catheter at discharge (N35)": "Removed"},
    },
    "Self-dilation taught": {
        "type": "toggle",
        "show_if": {"Stricture operation": ["Urethral dilation", "Internal urethrotomy (DVIU)", "Meatotomy"]},
    },
    # ----- Z43.6 nephrostomy change (48. daily department) -----
    "Nephrostomy side": {"type": "single", "options": ["Right", "Left", "Both"]},
    "Nephrostomy co-diagnoses": {
        "type": "multi", "optional": True, "custom": True,
        "options": list(NEPHROSTOMY_CODIAGNOSES),
        "custom_help": "Type it in Latvian, with the ICD code if you have it, e.g. C54 Dzemdes ķermeņa ļaundabīgs audzējs",
    },
    "Nephrostomy Ch": {"type": "single", "options": ["8", "10", "12", "14"], "default_index": 1},
    "Nephrostomy function": {"type": "single", "options": ["Satisfactory", "Reduced"], "default_index": 0},
    "Irrigation performed": {"type": "toggle", "default": True},
    "Fixed with skin sutures": {"type": "toggle", "default": True},
    "Infection signs at change": {"type": "toggle"},
    "Fever (48)": {"type": "toggle"},
    "Supervising doctor": {"type": "single", "options": ["Oncologist", "Urologist", "Family doctor"]},
    "Next nephrostomy change": {"type": "date", "optional": True},
    "Investigations (48)": {"type": "text", "optional": True, "placeholder": "optional, in Latvian"},
    # ----- Z43.6 nephrostomy fell out (NMC, or admitted to 4. Urology) -----
    "Time since it fell out": {"type": "single", "default_index": 0,
                               "options": ["Less than 24 hours", "1–3 days", "More than 3 days"]},
    "Systemic signs (NF)": {"type": "multi", "optional": True,
                            "options": ["Fever 38 °C or more", "Chills", "Hypotension / tachycardia",
                                        "Flank pain not controlled"]},
    "Solitary kidney": {"type": "toggle"},
    "Hydronephrosis on US (NF)": {"type": "single", "default_index": 0,
                                  "options": ["Present", "Not present", "Not assessed"]},
    "Blood tests (NF)": {"type": "single", "default_index": 0,
                         "options": ["Not done", "No inflammatory changes", "Raised CRP / leukocytosis",
                                     "Raised creatinine", "Raised CRP and creatinine"]},
    "Premedication given": {"type": "toggle", "default": True},
    "Puncture result (NF)": {"type": "single", "default_index": 0,
                             "options": ["Without complications", "Bleeding through the tube",
                                         "Complicated / not successful"]},
    "Urine from new nephrostomy": {"type": "single", "default_index": 0,
                                   "options": ["Clear", "Blood-tinged", "Cloudy / purulent"]},
    "Urine culture taken (NF)": {"type": "toggle"},
    "Ward treatment (NF)": {"type": "multi", "optional": True, "default_selected": ["IV antibiotics"],
                            "options": ["IV antibiotics", "IV fluids", "Analgesia"]},
    "Antibiotic in the ward (NF)": {"type": "text", "optional": True,
                                    "placeholder": "optional, e.g. Sol. Ceftriaxoni 2 g i/v 1 reizi dienā",
                                    "show_if": {"Ward treatment (NF)": "IV antibiotics"}},
    "Urine culture result (NF)": {"type": "text", "optional": True,
                                  "placeholder": "optional, e.g. E. coli, jutīga pret cefuroksīmu"},
    "Ward course (NF)": {"type": "single", "default_index": 0,
                         "options": ["Fever resolved, CRP falling", "Uneventful"]},
    "Definitive plan (NF)": {"type": "single", "default_index": 0,
                             "options": ["Urologist: plan definitive treatment", "Oncologist: oncology plan",
                                         "Long-term nephrostomy"]},
    # ----- N40 benign prostatic hyperplasia (TURP) -----
    "BPH indication": {
        "type": "multi",
        "options": ["Bothersome LUTS despite medication", "Recurrent urinary retention", "Recurrent UTIs",
                    "Recurrent hematuria", "Bladder stones", "Upper tract dilatation / renal impairment"],
    },
    "Prostate volume (ml)": {"type": "number", "min": 0.0, "max": 500.0, "step": 1.0, "optional": True},
    "Preoperative catheter": {"type": "toggle"},
    "TURP technique": {"type": "single", "options": ["Monopolar", "Bipolar"]},
    "Histology result (BPH)": {
        "type": "text", "placeholder": "e.g. labdabīga prostatas hiperplāzija", "show_if": {"Histology": "Known"},
    },
    "Previous BPH medication": {
        "type": "single", "default_index": 0, "options": ["None", "Stop after surgery", "Continue until review"],
    },
    "Failed urethral catheterisation": {"type": "toggle", "default": True},
    "Ultrasound guided": {"type": "toggle", "default": True},
    "Creatinine raised (N35)": {"type": "toggle"},
    "Investigations (N35)": {
        "type": "text", "optional": True,
        "placeholder": "optional; empty = 'veikti ambulatorā kārtā, pieļauj operāciju'",
    },
    # ----- C61 radical prostatectomy: drain, catheter removal place, wound -----
    "RP drain": {"type": "single", "default_index": 3,
                 "options": ["No drain", "Removed on day 1", "Removed on day 2", "Removed on day 3", "Removed on day 4",
                             "Removed on day 5", "Discharged with drain"]},
    "Catheter removal place": {"type": "single", "default_index": 0,
                               "options": ["PSKUS outpatient clinic", "Family doctor's practice", "Not stated"],
                               "show_if": {"Catheter removal": "Discharged with catheter"}},
    "Catheter removal time": {"type": "text", "optional": True, "placeholder": "optional, e.g. 10:00",
                              "show_if": {"Catheter removal": "Discharged with catheter"}},
    "Sauna and pool": {"type": "single", "default_index": 1, "options": ["Same as the rest period", "3 months"],
                       "show_if": {"Operation": "Radical prostatectomy"}},
    # ----- N43 hydrocele (hydrocele plasty in 4. Urology) -----
    "PUJ side": {"type": "single", "options": ["Right", "Left"]},
    # ----- D29.4 scrotal mass (excision in 4. Urology) -----
    # ----- N21.0 bladder stone (cystolithotripsy in 4. Urology) -----
    # ----- R33 acute urinary retention (NMC, sent home) -----
    "Retention cause": {"type": "multi", "optional": True,
                        "options": ["BPH", "Prostate cancer", "Urethral stricture", "Blocked indwelling catheter",
                                    "Clot retention (hematuria)", "Drug-induced", "After surgery / anaesthesia",
                                    "Constipation", "Urinary tract infection / prostatitis", "Neurogenic bladder",
                                    "Bladder hypotonia suspected", "Not clear yet"]},
    "Retention type": {"type": "single", "default_index": 0,
                       "options": ["Acute", "Acute-on-chronic (large volume)", "Chronic (overflow)"]},
    "Retention duration": {"type": "single", "default_index": 2,
                           "options": ["Hours", "1 day", "2 days", "3 days", "More than 3 days"]},
    "Previous retention episodes": {"type": "toggle"},
    "Ultrasound at NMC": {"type": "toggle", "default": True},
    "Enlarged prostate on US": {"type": "toggle", "show_if": {"Ultrasound at NMC": True, "_sex": "Man"}},
    "Upper tract dilatation": {"type": "single", "default_index": 0, "options": ["None", "Present"],
                               "show_if": {"Ultrasound at NMC": True}},
    "Labs at NMC": {"type": "single", "default_index": 0,
                    "options": ["Not stated", "No inflammatory changes", "Urinary tract infection", "Raised creatinine",
                                "UTI and raised creatinine"]},
    "Other NMC findings (optional)": {"type": "text", "optional": True,
                                      "placeholder": "optional, full sentence in Latvian, e.g. Veic CT m/k jostas daļai…"},
    "Back pain or neurological symptoms": {"type": "toggle"},
    "Retention drainage": {"type": "single",
                           "options": ["Urethral Foley catheter", "Urethral Tiemann (coudé) catheter",
                                       "Catheter changed (blocked catheter)", "Suprapubic catheter (epicystostomy)",
                                       "In-out catheterisation (no catheter left)"]},
    "Retention catheter Fr": {"type": "single", "default_index": 2, "options": ["12", "14", "16", "18", "20", "22"]},
    "Catheter plan (R33)": {"type": "single",
                            "options": ["Trial without catheter in 2–3 days", "Trial without catheter in 1 week",
                                        "Trial without catheter in 2 weeks", "Urologist decides",
                                        "Long-term catheter"],
                            "show_if": {"Retention drainage": ["Urethral Foley catheter",
                                                               "Urethral Tiemann (coudé) catheter",
                                                               "Catheter changed (blocked catheter)",
                                                               "Suprapubic catheter (epicystostomy)"]}},
    # ----- N30.0 acute cystitis (NMC, sent home) -----
    "Cystitis symptoms": {"type": "multi", "optional": True, "default_selected": ["Dysuria", "Frequency"],
                          "options": ["Dysuria", "Frequency", "Urgency", "Suprapubic pain", "Visible hematuria",
                                      "New urge incontinence", "Cloudy / foul-smelling urine"]},
    "Symptom duration (N30)": {"type": "single", "default_index": 1,
                               "options": ["Since today", "1–2 days", "3–7 days", "More than a week"]},
    "Pyelonephritis signs (N30)": {"type": "multi", "optional": True,
                                   "options": ["Fever 38 °C or more", "Chills", "Flank pain / costovertebral tenderness",
                                               "Nausea / vomiting"]},
    "Vaginal discharge": {"type": "toggle", "show_if": {"_sex": "Woman"}},
    "Pregnant (N30)": {"type": "toggle", "show_if": {"_sex": "Woman"}},
    "Pregnancy week": {"type": "number", "min": 1.0, "max": 42.0, "step": 1.0, "optional": True,
                       "show_if": {"Pregnant (N30)": True}},
    "Postmenopausal": {"type": "toggle", "show_if": {"_sex": "Woman"}},
    "Recurrent cystitis": {"type": "toggle"},
    "Same organism soon after treatment": {"type": "toggle", "show_if": {"Recurrent cystitis": True}},
    "Antibiotic in the last 3 months": {"type": "toggle"},
    "Risk factors (N30)": {"type": "multi", "optional": True,
                           "options": ["Diabetes", "Immunosuppression", "Urinary tract abnormality / stone / obstruction",
                                       "Recent urological procedure", "Indwelling catheter or self-catheterisation",
                                       "Renal impairment", "Neurogenic bladder", "Age 65 or older / frail"]},
    "Penicillin allergy": {"type": "toggle"},
    "Suprapubic tenderness": {"type": "toggle", "default": True},
    "Urinalysis (N30)": {"type": "multi", "optional": True, "default_selected": ["Leukocytes", "Nitrite"],
                         "options": ["Leukocytes", "Nitrite", "Erythrocytes", "Bacteria"]},
    "Urine culture taken": {"type": "toggle"},
    "Blood tests (N30)": {"type": "single", "default_index": 0,
                          "options": ["Not done", "No inflammatory changes", "Raised CRP / leukocytosis",
                                      "Raised creatinine"]},
    "Ultrasound (N30)": {"type": "multi", "optional": True,
                         "options": ["No pathology", "Post-void residual urine", "Upper tract dilatation",
                                     "Bladder wall thickening", "Bladder stone", "Enlarged prostate"]},
    "Residual urine (ml)": {"type": "number", "min": 0.0, "max": 3000.0, "step": 10.0, "optional": True,
                            "show_if": {"Ultrasound (N30)": "Post-void residual urine"}},
    "Treatment approach (N30)": {"type": "single", "default_index": 0,
                                 "options": ["Antibiotic", "Symptomatic first (delayed antibiotic)"]},
    "Referral (N30)": {"type": "single", "default_index": 0,
                       "options": ["Family doctor only", "Urologist", "Gynaecologist"]},
    # ----- R31 hematuria (NMC consultation, or admitted to 4. Urology) -----
    "Hematuria type": {"type": "single", "default_index": 0,
                       "options": ["Visible (macrohematuria)", "Non-visible (microhematuria)"]},
    "Hematuria duration": {"type": "single", "default_index": 1,
                           "options": ["Since today", "1–2 days", "3–7 days", "More than a week", "Recurrent episodes"]},
    "Hematuria symptoms": {"type": "multi", "optional": True,
                           "options": ["Clots in urine", "Suprapubic pain", "Flank pain", "Dysuria / frequency"]},
    "Clot retention (bladder tamponade)": {"type": "toggle"},
    "Hematuria cause": {"type": "multi", "optional": True, "default_selected": ["Not clear yet"],
                        "options": ["Not clear yet", "Bladder cancer (known)", "Prostate cancer", "Radiation cystitis",
                                    "BPH (prostatic bleeding)", "Urinary stone",
                                    "Urinary tract infection (hemorrhagic cystitis)", "After a urological procedure",
                                    "Catheter trauma"]},
    "Risk factors (R31)": {"type": "multi", "optional": True,
                           "options": ["Smoker or ex-smoker", "Occupational exposure (aromatic amines, dyes)",
                                       "Pelvic radiotherapy in the past", "Cyclophosphamide in the past"]},
    "Antithrombotic therapy": {"type": "multi", "optional": True,
                               "options": ["Aspirin", "Clopidogrel", "Ticagrelor", "Warfarin", "Apixaban",
                                           "Rivaroxaban", "Dabigatran", "Edoxaban", "LMWH (Fraxiparine / Clexane)"]},
    "Antithrombotic plan": {"type": "single",
                            "options": ["Continue", "Pause until the urine is clear", "Prescriber decides"],
                            "show_if": {"Antithrombotic therapy": "*"}},
    "INR (optional)": {"type": "number", "min": 0.0, "max": 20.0, "step": 0.1, "optional": True,
                       "show_if": {"Antithrombotic therapy": "Warfarin"}},
    "Hemodynamics (R31)": {"type": "single", "default_index": 0, "options": ["Stable", "Unstable"]},
    "Hemoglobin (g/L)": {"type": "number", "min": 0.0, "max": 250.0, "step": 1.0, "optional": True},
    "Ultrasound findings (R31)": {"type": "multi", "optional": True,
                                  "options": ["Clots in the bladder", "Bladder wall mass", "Enlarged prostate",
                                              "Kidney mass", "Kidney or ureteral stone", "Hydronephrosis",
                                              "Residual urine"],
                                  "show_if": {"Ultrasound at NMC": True}},
    "CT at NMC (R31)": {"type": "single", "default_index": 0,
                        "options": ["Not done", "CT urography – no source found", "CT – bladder mass",
                                    "CT – kidney / upper tract mass", "CT – stone"]},
    "NMC catheter (R31)": {"type": "single",
                           "options": ["No catheter – voids", "Catheter, washout, removed – voids",
                                       "Discharged with a catheter", "Catheter changed (blocked)"]},
    "Catheter type (R31)": {"type": "single", "default_index": 0, "options": ["3-way catheter", "Foley (2-way)"],
                            "show_if": {"NMC catheter (R31)": ["Catheter, washout, removed – voids",
                                                               "Discharged with a catheter",
                                                               "Catheter changed (blocked)"]}},
    "Catheter Fr (R31)": {"type": "single", "default_index": 3, "options": ["16", "18", "20", "22", "24"],
                          "show_if": {"NMC catheter (R31)": ["Catheter, washout, removed – voids",
                                                             "Discharged with a catheter",
                                                             "Catheter changed (blocked)"]}},
    "Bladder washout (R31)": {"type": "toggle", "default": True,
                              "show_if": {"NMC catheter (R31)": ["Catheter, washout, removed – voids",
                                                                 "Discharged with a catheter",
                                                                 "Catheter changed (blocked)"]}},
    "Urine colour at discharge": {"type": "single",
                                  "options": ["Clear", "Light pink", "Pink, no clots", "Red / with clots"]},
    "Ward treatment (R31)": {"type": "multi", "optional": True,
                             "default_selected": ["Bladder irrigation (3-way catheter)"],
                             "options": ["Bladder irrigation (3-way catheter)", "Manual bladder washout",
                                         "IV tranexamic acid", "IV antibiotics"]},
    "Transfusion (RBC units)": {"type": "number", "min": 0.0, "max": 30.0, "step": 1.0, "optional": True},
    "Hematuria procedure": {"type": "multi", "optional": True,
                            "options": ["Cystoscopy, clot evacuation, coagulation", "TURU (resection / biopsy)",
                                        "TURP (prostatic bleeding)", "Selective embolisation",
                                        "Alum bladder irrigation (1%)"]},
    "Cystoscopy date (R31)": {"type": "date", "show_if": {"Hematuria procedure": ["Cystoscopy, clot evacuation, coagulation", "TURU (resection / biopsy)", "TURP (prostatic bleeding)"]}},
    "Anesthesia (R31)": {"type": "single", "options": ["Spinal", "General"],
                         "show_if": {"Hematuria procedure": ["Cystoscopy, clot evacuation, coagulation", "TURU (resection / biopsy)", "TURP (prostatic bleeding)"]}},
    "Cystoscopy finding (R31)": {"type": "single",
                                 "options": ["No tumour, no bleeding source", "Diffuse bleeding (radiation cystitis)",
                                             "Bleeding from the prostate", "Bladder tumour",
                                             "Blood from a ureteric orifice"],
                                 "show_if": {"Hematuria procedure": ["Cystoscopy, clot evacuation, coagulation", "TURU (resection / biopsy)", "TURP (prostatic bleeding)"]}},
    "Embolisation date": {"type": "date", "show_if": {"Hematuria procedure": "Selective embolisation"}},
    "Histology (R31)": {"type": "single", "default_index": 0, "options": ["Not taken", "Pending", "Known"],
                        "show_if": {"Hematuria procedure": ["Cystoscopy, clot evacuation, coagulation", "TURU (resection / biopsy)", "TURP (prostatic bleeding)"]}},
    "Histology result (R31)": {"type": "text", "placeholder": "in Latvian, e.g. uroteliāla karcinoma pTa LG",
                               "show_if": {"Histology (R31)": "Known"}},
    "Discharge paper (R31)": {"type": "single", "default_index": 0, "options": ["Temporary", "Final"],
                              "show_if": {"Histology (R31)": ["Pending", "Known"]}},
    "Posthemorrhagic anemia": {"type": "toggle"},
    "Hemoglobin at discharge (g/L)": {"type": "number", "min": 0.0, "max": 250.0, "step": 1.0, "optional": True},
    "Green corridor (GP referral)": {"type": "single", "default_index": 0,
                                     "options": ["Not needed", "Bladder – Z03.167", "Kidney – Z03.164"]},
    "Urology consultation (R31)": {"type": "single", "options": ["Within 1 week", "In 1–2 weeks", "In 4 weeks"],
                                   "show_if": {"Green corridor (GP referral)": "Not needed"}},
    "Bladder stone number": {"type": "single", "default_index": 0, "options": ["Single", "Multiple"]},
    "Bladder stone size (cm)": {"type": "number", "min": 0.0, "max": 20.0, "step": 0.5, "optional": True},
    "Bladder stone found by": {"type": "multi", "optional": True, "options": ["Ultrasound", "CT", "KUB X-ray", "Cystoscopy"],
                               "default_selected": ["Ultrasound"]},
    "Bladder stone cause": {"type": "multi", "optional": True,
                            "options": ["BPH / bladder outlet obstruction", "Urethral stricture", "Neurogenic bladder",
                                        "Long-term catheter", "Foreign body / clips or sutures after surgery",
                                        "Augmented bladder / reservoir", "Not found"]},
    "Cystolithotripsy method": {"type": "single",
                                "options": ["Transurethral – laser (Ho:YAG)", "Transurethral – pneumatic",
                                            "Transurethral – mechanical lithotriptor", "Percutaneous cystolithotripsy",
                                            "Open cystolithotomy"]},
    "Done at the same time (N21)": {"type": "multi", "optional": True,
                                    "options": ["TURP", "Optical urethrotomy", "Foreign body removal"]},
    "Anesthesia (N21)": {"type": "single", "options": ["Spinal", "General"]},
    "Urine culture (N21)": {"type": "single", "default_index": 0,
                            "options": ["Not taken", "Negative", "Positive", "Result pending"]},
    "Culture organism (N21)": {"type": "text", "optional": True, "placeholder": "optional, e.g. E. coli",
                               "show_if": {"Urine culture (N21)": "Positive"}},
    "Stone sent for analysis": {"type": "toggle", "default": True},
    "Scrotal mass location": {"type": "single", "default_index": 0,
                              "options": ["Scrotal wall (subcutaneous)", "Scrotum and perineum", "Scrotal skin",
                                          "Spermatic cord / paratesticular"]},
    "Scrotal mass side": {"type": "single", "options": ["Right", "Left", "Midline (raphe)", "Both sides"]},
    "Scrotal mass number": {"type": "single", "default_index": 0, "options": ["Single", "Multiple"]},
    "Scrotal mass size (cm)": {"type": "number", "min": 0.0, "max": 30.0, "step": 0.5, "optional": True},
    "Scrotal mass type": {"type": "single", "default_index": 0,
                          "options": ["Not specified", "Epidermoid (sebaceous) cyst", "Lipoma", "Fibroma",
                                      "Scrotal calcinosis", "Other (type below)"]},
    "Scrotal mass type (other)": {"type": "text", "placeholder": "in Latvian, nominative, e.g. hemangioma",
                                  "show_if": {"Scrotal mass type": "Other (type below)"}},
    "Inflamed / infected lesion": {"type": "toggle"},
    "Scrotal anesthesia": {"type": "single", "options": ["Spinal", "General", "Local"]},
    "Scrotal drain": {"type": "toggle"},
    "Scrotal histology result": {"type": "text", "placeholder": "in Latvian, e.g. epidermoidāla cista",
                                 "show_if": {"Histology": "Known"}},
    "Histology expected (date)": {"type": "date", "optional": True, "show_if": {"Histology": "Pending"}},
    "Other diagnoses (optional)": {"type": "text", "optional": True,
                                   "placeholder": "in Latvian, one line, e.g. Simptomātiska LPH"},
    "PUJ cause": {"type": "single", "default_index": 0,
                  "options": ["Not specified", "Intrinsic (congenital) stenosis", "Crossing lower-pole vessel",
                              "Secondary (after stones or surgery)"]},
    "Split renal function (%)": {"type": "number", "min": 0.0, "max": 100.0, "step": 1.0, "optional": True},
    "Pyeloplasty approach": {"type": "single", "options": ["Laparoscopic", "Robot-assisted", "Open"]},
    "Pyeloplasty technique": {"type": "single", "default_index": 0,
                              "options": ["Dismembered (Anderson–Hynes)", "Non-dismembered (Y-V plasty)"]},
    "Vessel transposition": {"type": "toggle", "default": True, "show_if": {"PUJ cause": "Crossing lower-pole vessel"}},
    "Pyeloplasty stent": {"type": "single", "default_index": 0, "options": ["DJ stent", "No stent"]},
    "Stent removal plan": {"type": "single", "default_index": 0,
                           "options": ["In 4–6 weeks", "On a specific date"], "show_if": {"Pyeloplasty stent": "DJ stent"}},
    "Stent removal date (N13)": {"type": "date", "show_if": {"Stent removal plan": "On a specific date"}},
    "Pyeloplasty drain": {"type": "single", "default_index": 2,
                          "options": ["No drain", "Removed on day 1", "Removed on day 2", "Removed on day 3",
                                      "Removed on day 4", "Removed on day 5", "Discharged with drain"]},
    "PUJ histology result": {"type": "text", "placeholder": "in Latvian, e.g. hroniska iekaisuma pazīmes, fibroze",
                             "show_if": {"Histology": "Known"}},
    "Hydrocele side": {"type": "single", "options": ["Right", "Left"]},
    "Scrotal ultrasound before surgery": {"type": "toggle", "default": True},
    "Hydrocele technique": {"type": "single", "options": ["Eversion (Winkelmann)", "Excision (Bergmann)",
                                                          "Plication (Lord)"]},
    "N43 anesthesia": {"type": "single", "options": ["Spinal", "General", "Local"]},
    "Hydrocele drain": {"type": "single", "options": ["Removed on day 1", "No drain", "Discharged with drain"]},
    "Drain removal date": {"type": "date", "optional": True, "show_if": {"Hydrocele drain": "Discharged with drain"}},
    "Scrotal hematoma": {"type": "toggle"},
    "Skin sutures": {"type": "single", "options": ["Absorbable", "Non-absorbable"]},
}

# Which modifiers are asked for which diagnosis (in this order).
# A diagnosis that is not listed here gets DEFAULT_MODIFIERS.
DEFAULT_MODIFIERS = ["Admission", "Stay", "Infection", "Surgery"]
DIAGNOSIS_MODIFIERS = {
    "N10": [
        ("1. Admission", ["Admission", "Stay", "Side", "Infection complexity"]),
        ("2. NMC findings",
         ["NMC imaging", "Urostasis on imaging", "Sepsis signs", "Able to take oral medication", "Pregnant",
          "Initial IV antibiotic at NMC", "Urology consultation"],
         {"_department": "NMC"}),
    ],
    "N47": [
        ("1. Admission", ["Stay", "Surgery", "Conservative reason"], {"_department": ["4", "48"]}),
        ("2. Findings", ["Phimosis degree", "Balanoposthitis", "Frenulum breve"], {"_department": ["4", "48"]}),
        ("3. Operation", ["Surgery type", "Frenuloplasty", "Operation date", "Anesthesia"],
         {"Surgery": True, "_department": ["4", "48"]}),
        ("1. Paraphimosis", ["Paraphimosis duration", "Paraphimosis cause", "Glans appearance", "Balanoposthitis"],
         {"_department": "NMC"}),
        ("2. Treatment at NMC", ["Procedure date", "Paraphimosis treatment", "Before reduction (PF)",
                                 "Dorsumcision sutures"], {"_department": "NMC"}),
        ("3. Recommendations", ["Control visit (PF)", "Control visit date (PF)", "Control visit time (PF)",
                                "Suture removal at the visit", "Plan circumcision"], {"_department": "NMC"}),
    ],
    "T83": ["Catheter type", "Catheter size (Fr)", "Catheter reason", "Urinary retention (diagnosis)",
            "Drained volume (ml)", "Urine appearance at change", "Fever at NMC"],
    # A section may have a third part: it is only shown when those choices are made.
    "N20": [
        ("1. Admission", ["Admission", "Stay", "Side", "Stone location"]),
        ("2. Stone and infection signs",
         ["Stone size (mm)", "Right stone size (mm)", "Left stone size (mm)", "Urostasis grade",
          "CRO elevated", "CRO (mg/L)", "Leukocytosis", "Urine infection signs"],
         {"Admission": "Acute"}),
        ("3. NMC decision", ["Pain controlled"], {"_department": "NMC"}),
        ("3. Drainage and treatment",
         ["Urosepsis", "Drainage", "Drainage date", "Purulent urine on drainage", "Antibiotics",
          "Urine culture", "Culture organism", "Culture checked by", "Clinical status", "Temperature trend",
          "Inflammatory markers trend"],
         {"Admission": "Acute", "_department": "4"}),
        ("4. Definitive treatment", ["Definitive stone treatment", "Definitive treatment date"],
         {"Admission": "Acute", "_department": "4"}),
        ("2. Diagnosis", ["Renal colic", "Urinary tract infection"], {"Admission": "Planned"}),
        ("3. Operation",
         ["Procedure", "Lithotripsy", "DJ stent placed", "Procedure date", "Intraoperative course",
          "Intraoperative complications", "Postoperative course", "Postoperative complications",
          "Discharge status", "Investigations"],
         {"Admission": "Planned"}),
        ("4. Recommendations",
         ["Fluid intake", "Stent removal", "Stent removal date", "Discharged with catheter",
          "Catheter removal date", "Urology follow-up", "Urology follow-up date"],
         {"Admission": "Planned"}),
    ],
    "N39.0": [
        ("1. Admission and symptoms",
         ["Admission", "Admission date", "Discharge date", "Symptoms", "Symptom duration", "Symptom onset",
          "Fever status", "Relevant history"]),
        ("2. Diagnosis",
         ["Infection complexity", "Infection site", "Pyelonephritis side", "Infection context",
          "Complications", "Complications status", "Comorbidities", "Diagnosis certainty"]),
        ("3. Urinary catheter",
         ["Has catheter", "Catheter route", "Catheter action", "Catheter Fr", "Catheter indication",
          "Catheter at discharge", "Urine drainage", "Voiding after removal"]),
        ("4. Ureteral stents",
         ["Has stents", "Stent count", "Stent side", "Stent type", "Stent indication", "Stent insertion date",
          "Stent age", "Stents at discharge", "Stent plan", "Stent plan date"]),
        ("5. Urine and imaging", ["Urine appearance", "Imaging", "Hydronephrosis", "Obstruction drained"]),
        ("6. Antibiotics and safety checks",
         ["Antibiotics", "Antibiotic allergy status", "Antibiotic renal adjustment", "Antibiotic review"]),
        ("7. Microbiology and laboratory results",
         ["Urine specimen", "Urine culture", "Culture organism", "Culture checked by", "Blood culture", "Labs"]),
        ("8. Clinical course at discharge",
         ["Clinical status", "Temperature trend", "Pain trend", "Urinary symptoms trend",
          "Inflammatory markers trend", "Renal function trend", "Hemodynamic status",
          "Oral intake", "Treatment tolerance", "Discharge readiness"],
         {"_department": "4"}),
        ("8. NMC decision", ["Sepsis signs", "Hemodynamic status", "Able to take oral medication"],
         {"_department": "NMC"}),
        ("9. Follow-up and advice",
         ["Urology follow-up", "Urology follow-up date", "Fluid intake", "Warning signs",
          "Chronic medication changes"]),
    ],
    "N41.0": [
        ("1. Complaints and history",
         ["Symptoms (AP)", "Symptom duration (AP)", "Context (AP)", "Risk factors (AP)", "Other diagnoses (optional)"]),
        ("2. Findings at NMC",
         ["DRE (AP)", "Labs (AP)", "Urine tests (AP)", "Imaging (AP)", "Abscess size (cm) (AP)", "Sepsis signs (AP)",
          "Retention drainage (AP)", "Drained volume (AP)", "Other NMC findings (optional)"]),
        ("3. Treatment at NMC", ["First IV dose at NMC (AP)", "Control visit (AP)"], {"_department": "NMC"}),
        ("3. Treatment in 4. Urology",
         ["IV antibiotic (AP)", "Abscess treatment (AP)", "Abscess procedure date (AP)", "Culture result (AP)",
          "Ward course (AP)", "Catheter at discharge (AP)", "Catheter removal date (AP)", "Control visit (AP)"],
         {"_department": "4"}),
    ],
    "N45": [
        ("1. Complaints and history",
         ["Side (EO)", "Inflammation site (EO)", "Onset (EO)", "Symptom duration (EO)", "Symptoms (EO)",
          "Likely source (EO)", "Risk factors (EO)", "Other diagnoses (optional)"]),
        ("2. Findings at NMC",
         ["Scrotal US (EO)", "Labs (EO)", "Urine tests (EO)", "Sepsis signs (EO)", "Other NMC findings (optional)"]),
        ("3. Treatment at NMC", ["Ceftriaxone given at NMC", "Control visit (EO)"], {"_department": "NMC"}),
        ("3. Treatment in 4. Urology",
         ["Treatment (EO)", "IV antibiotic (EO)", "After failed conservative therapy", "Operation (EO)",
          "Procedure date", "Anesthesia (EO)", "Postoperative course (EO)", "Postoperative complications (EO)",
          "Drain (EO)",
          "Histology (EO)", "Histology result (EO)", "Culture result (EO)"], {"_department": "4"}),
        ("4. Recommendations", ["Activity restriction", "Control visit (EO)"], {"_department": "4"}),
    ],
    "C67": [
        ("1. Admission", ["C67 operation", "Admission", "Smoker (C67)"]),
        ("2. Tumours (EAU risk factors)",
         ["Tumour verification", "Tumour number", "Tumour location", "Recurrent tumour", "Largest tumour over 3 cm",
          "Age over 70 (C67)"], {"C67 operation": "TURU / re-TURU"}),
        ("3. Operation and course",
         ["TURU type", "Procedure date", "Operation details", "Intravesical treatment",
          "Intravesical treatment (other)", "Postoperative course", "Postoperative complications",
          "Catheter removal", "Catheter removed on", "Planned catheter removal (optional)",
          "Voiding after catheter removal", "Histology", "Tumour stage", "N stage", "M stage",
          "Tumour grade", "Concomitant CIS", "Detrusor muscle in specimen", "Variant histology / LVI",
          "Histology details"], {"C67 operation": "TURU / re-TURU"}),
        ("4. Recommendations",
         ["Discharge paper", "Adjuvant plan (NMIBC)", "Activity restriction", "Urology visit", "Oncology council",
          "Oncology council date"], {"C67 operation": "TURU / re-TURU"}),
        ("2. Diagnosis before cystectomy",
         ["Indication (RC)", "Clinical T (RC)", "Clinical N (RC)", "Clinical M (RC)", "Previous treatment (RC)",
          "Persistent hematuria (RC)", "Anaemia (RC)", "Preoperative transfusion (RC, units)",
          "Other diagnoses (optional)"], {"C67 operation": "Radical cystectomy"}),
        ("3. Cystectomy",
         ["Procedure date", "Cystectomy type (man)", "Cystectomy type (woman)", "Urethrectomy (RC)", "Approach (RC)",
          "Lymph node dissection (RC)", "Urinary diversion", "Anesthesia (RC)", "Intraoperative course",
          "Intraoperative complications"], {"C67 operation": "Radical cystectomy"}),
        ("4. Course after cystectomy",
         ["Postoperative course", "Postoperative complications", "Postoperative transfusion (RC, units)",
          "Bowel recovery (RC)", "Drain (RC)", "Drain removed on day (RC)", "Ureteral stents (RC)",
          "Stent removal date (RC)", "Stent removal time (RC)", "Neobladder catheter removal (RC)",
          "Histology (RC)", "Histology expected (RC)", "pT (RC)", "pN (RC)", "Nodes positive / removed (RC)",
          "Margins (RC)", "Histology details (RC)"], {"C67 operation": "Radical cystectomy"}),
        ("5. Recommendations",
         ["Discharge paper", "Diabetes diet (RC)", "Skin sutures", "Urology visit (RC)", "Next nephrostomy change",
          "Oncology council", "Oncology council date"], {"C67 operation": "Radical cystectomy"}),
    ],
    "C61": [
        ("1. Admission", ["Admission", "Stay", "Operation"]),
        ("2. Prostate cancer",
         ["PSA (ng/ml)", "Clinical T stage", "Prostate N stage", "Prostate M stage", "Biopsy Gleason score"]),
        ("3. NMC visit",
         ["NMC reason", "NMC catheter Fr", "Drained urine (ml)", "Hematuria management",
          "Hemodynamically stable", "Urine clearing", "Pain location", "Pain controlled", "Fever at NMC",
          "Discharged with catheter", "Catheter removal date", "Urology consultation"],
         {"_department": "NMC"}),
        ("3. Radical prostatectomy",
         ["RP approach", "Lymph node dissection", "Procedure date", "Intraoperative course",
          "Intraoperative complications", "RP drain"],
         {"Operation": "Radical prostatectomy"}),
        ("3. TURP",
         ["TURP anesthesia", "Procedure date", "Intraoperative course",
          "Intraoperative complications", "Bladder irrigation after surgery"],
         {"Operation": "TURP"}),
        ("4. Course and histology",
         ["Postoperative course", "Postoperative complications", "Catheter removal", "Catheter removed on",
          "Planned catheter removal (optional)", "Catheter removal place", "Catheter removal time",
          "Voiding after catheter removal", "Histology",
          "Pathological T stage", "Pathological N stage", "Surgical margin", "Final Gleason score",
          "Prostate histology details"],
         {"_department": "4"}),
        ("5. Recommendations",
         ["Discharge paper", "Activity restriction", "Sauna and pool", "Skin sutures", "Urology visit",
          "Oncology council", "Oncology council date"],
         {"_department": "4"}),
    ],
    "C64": [
        ("1. Admission", ["Admission", "Stay", "Kidney operation"]),
        ("2. Kidney tumour",
         ["Kidney side", "Tumour size (cm)", "RCC clinical T stage", "RCC N stage", "RCC M stage"]),
        ("3. NMC visit",
         ["NMC reason", "NMC catheter Fr", "Drained urine (ml)", "Hematuria management",
          "Hemodynamically stable", "Urine clearing", "Pain location", "Pain controlled", "Fever at NMC",
          "Discharged with catheter", "Catheter removal date", "Urology consultation"],
         {"_department": "NMC"}),
        ("3. Operation",
         ["Kidney approach", "Adrenalectomy", "Lymph node dissection", "Procedure date",
          "Intraoperative course", "Intraoperative complications"],
         {"_department": "4"}),
        ("4. Course and histology",
         ["Postoperative course", "Postoperative complications", "Renal function at discharge",
          "Creatinine at discharge (µmol/l)", "eGFR at discharge", "Histology", "RCC subtype",
          "RCC pT stage", "RCC pN stage", "RCC margin", "RCC ISUP grade", "Kidney histology details",
          "Follow-up risk group"],
         {"_department": "4"}),
        ("5. Recommendations",
         ["Discharge paper", "Activity restriction", "Urology visit", "Oncology council", "Oncology council date"],
         {"_department": "4"}),
    ],
    "N35": [
        ("1. Admission and stricture",
         ["Admission", "Stricture location", "Stricture length (cm)", "Stricture compensation", "Stricture cause",
          "Previous stricture treatment", "Stricture verified by", "Other diagnoses (optional)"]),
        ("2. Epicystostomy at NMC",
         ["Failed urethral catheterisation", "Ultrasound guided", "Procedure date", "Catheter Fr (N35)",
          "Drained volume (ml)", "Urine appearance", "Creatinine raised (N35)", "Fever at NMC", "Urology consultation"],
         {"_department": "NMC"}),
        ("2. Operation",
         ["Stricture operation", "Urethrotomy instrument", "Also performed (N35)", "N35 anesthesia", "Procedure date",
          "Postoperative course", "Postoperative complications", "Tissue for histology", "Histology result (N35)"],
         {"_department": "4"}),
        ("3. Catheter and voiding",
         ["Catheter Fr (N35)", "Catheter at discharge (N35)", "Planned catheter removal (N35)",
          "Catheter removal place (N35)", "Catheter removal time (N35)", "Voiding after removal (N35)",
          "Self-dilation taught"],
         {"_department": "4"}),
        ("4. Other", ["Investigations (N35)", "Rest after endoscopic treatment"], {"_department": "4"}),
    ],
    "Z43.6": [
        ("1. Nephrostomy and co-diagnoses", ["Nephrostomy side", "Nephrostomy co-diagnoses"]),
        ("2. Nephrostomy change",
         ["Procedure date", "Nephrostomy Ch", "Nephrostomy function", "Irrigation performed",
          "Fixed with skin sutures", "Infection signs at change", "Fever (48)"], {"_department": "48"}),
        ("3. Recommendations", ["Supervising doctor", "Next nephrostomy change", "Investigations (48)"],
         {"_department": "48"}),
        ("2. Nephrostomy fell out",
         ["Time since it fell out", "Systemic signs (NF)", "Solitary kidney", "Hydronephrosis on US (NF)",
          "Blood tests (NF)", "Other NMC findings (optional)"], {"_department": ["NMC", "4"]}),
        ("3. New nephrostomy (radiologist)",
         ["Procedure date", "Premedication given", "Nephrostomy Ch", "Puncture result (NF)",
          "Urine from new nephrostomy", "Urine culture taken (NF)", "Fixed with skin sutures"],
         {"_department": ["NMC", "4"]}),
        ("4. Stay in 4. Urology",
         ["Admission date", "Discharge date", "Ward treatment (NF)", "Antibiotic in the ward (NF)",
          "Urine culture result (NF)", "Ward course (NF)"], {"_department": "4"}),
        ("5. Recommendations", ["Supervising doctor", "Next nephrostomy change", "Definitive plan (NF)"],
         {"_department": ["NMC", "4"]}),
    ],
    "N40": [
        ("1. Admission and indication", ["Admission", "BPH indication", "Prostate volume (ml)", "Preoperative catheter"]),
        ("2. Operation",
         ["TURP technique", "TURP anesthesia", "Procedure date", "Intraoperative course",
          "Intraoperative complications", "Bladder irrigation after surgery"]),
        ("3. Course and histology",
         ["Postoperative course", "Postoperative complications", "Catheter removal", "Catheter removed on",
          "Planned catheter removal (optional)", "Voiding after catheter removal", "Histology",
          "Histology result (BPH)"]),
        ("4. Recommendations",
         ["Discharge paper", "Activity restriction", "Urology visit", "Previous BPH medication"]),
    ],
    "N43": [
        ("1. Admission", ["Admission", "Hydrocele side", "Scrotal ultrasound before surgery"]),
        ("2. Operation",
         ["Hydrocele technique", "N43 anesthesia", "Procedure date", "Intraoperative course",
          "Intraoperative complications"]),
        ("3. Course", ["Postoperative course", "Postoperative complications", "Hydrocele drain", "Drain removal date",
                       "Scrotal hematoma"]),
        ("4. Recommendations", ["Activity restriction", "Skin sutures", "Urology visit"]),
    ],
    "N13.0": [
        ("1. Admission", ["Admission", "PUJ side", "PUJ cause", "Split renal function (%)"]),
        ("2. Operation",
         ["Pyeloplasty approach", "Pyeloplasty technique", "Vessel transposition", "Procedure date",
          "Intraoperative course", "Intraoperative complications", "Pyeloplasty stent"]),
        ("3. Course and histology",
         ["Postoperative course", "Postoperative complications", "Pyeloplasty drain", "Catheter removal",
          "Catheter removed on", "Planned catheter removal (optional)", "Voiding after catheter removal",
          "Histology", "PUJ histology result"]),
        ("4. Recommendations",
         ["Discharge paper", "Activity restriction", "Skin sutures", "Stent removal plan", "Stent removal date (N13)",
          "Urology visit"]),
    ],
    "D29.4": [
        ("1. Admission and the mass",
         ["Admission", "Scrotal mass location", "Scrotal mass side", "Scrotal mass number", "Scrotal mass size (cm)",
          "Scrotal mass type", "Scrotal mass type (other)", "Scrotal ultrasound before surgery",
          "Inflamed / infected lesion", "Other diagnoses (optional)"]),
        ("2. Operation and course",
         ["Scrotal anesthesia", "Procedure date", "Intraoperative course", "Intraoperative complications",
          "Scrotal drain", "Postoperative course", "Postoperative complications", "Scrotal hematoma"]),
        ("3. Histology", ["Histology", "Scrotal histology result", "Histology expected (date)"]),
        ("4. Recommendations", ["Discharge paper", "Activity restriction", "Skin sutures", "Urology visit"]),
    ],
    "N30.0": [
        ("1. Complaints and history",
         ["Cystitis symptoms", "Symptom duration (N30)", "Pyelonephritis signs (N30)", "Vaginal discharge",
          "Pregnant (N30)", "Pregnancy week", "Postmenopausal", "Recurrent cystitis",
          "Same organism soon after treatment", "Antibiotic in the last 3 months", "Risk factors (N30)",
          "Penicillin allergy", "Other diagnoses (optional)"]),
        ("2. NMC findings",
         ["Suprapubic tenderness", "Urinalysis (N30)", "Urine culture taken", "Blood tests (N30)", "Ultrasound (N30)",
          "Residual urine (ml)", "Other NMC findings (optional)"]),
        ("3. Plan", ["Treatment approach (N30)", "Referral (N30)"]),
    ],
    "R31": [
        ("1. Complaints and history",
         ["Hematuria type", "Hematuria duration", "Hematuria symptoms", "Clot retention (bladder tamponade)",
          "Hematuria cause", "Risk factors (R31)", "Antithrombotic therapy", "Antithrombotic plan", "INR (optional)",
          "Fever at NMC", "Other diagnoses (optional)"]),
        ("2. Findings at NMC",
         ["Hemodynamics (R31)", "Hemoglobin (g/L)", "Labs at NMC", "Ultrasound at NMC", "Ultrasound findings (R31)",
          "CT at NMC (R31)", "Other NMC findings (optional)"]),
        ("3. Treatment at NMC",
         ["NMC catheter (R31)", "Catheter type (R31)", "Catheter Fr (R31)", "Bladder washout (R31)",
          "Urine colour at discharge"], {"_department": "NMC"}),
        ("4. Plan", ["Green corridor (GP referral)", "Urology consultation (R31)"], {"_department": "NMC"}),
        ("3. Treatment in 4. Urology",
         ["Ward treatment (R31)", "Transfusion (RBC units)", "Hematuria procedure", "Cystoscopy date (R31)",
          "Anesthesia (R31)", "Cystoscopy finding (R31)", "Embolisation date", "Histology (R31)",
          "Histology result (R31)"], {"_department": "4"}),
        ("4. Discharge",
         ["Urine colour at discharge", "Hemoglobin at discharge (g/L)", "Posthemorrhagic anemia",
          "Discharged with catheter", "Catheter removal date", "Discharge paper (R31)", "Activity restriction",
          "Green corridor (GP referral)", "Urology visit"], {"_department": "4"}),
    ],
    "R33": [
        ("1. Complaints and cause",
         ["Retention type", "Retention cause", "Retention duration", "Previous retention episodes",
          "Back pain or neurological symptoms", "Fever at NMC", "Other diagnoses (optional)"]),
        ("2. NMC findings",
         ["Ultrasound at NMC", "Enlarged prostate on US", "Upper tract dilatation", "Labs at NMC",
          "Other NMC findings (optional)"]),
        ("3. Drainage",
         ["Procedure date", "Retention drainage", "Retention catheter Fr", "Drained volume (ml)", "Urine appearance"]),
        ("4. Plan", ["Catheter plan (R33)", "Urology consultation"]),
    ],
    "N21.0": [
        ("1. Admission and the stone",
         ["Admission", "Bladder stone number", "Bladder stone size (cm)", "Bladder stone found by",
          "Bladder stone cause", "Urine culture (N21)", "Culture organism (N21)", "Other diagnoses (optional)"]),
        ("2. Operation",
         ["Cystolithotripsy method", "Done at the same time (N21)", "Anesthesia (N21)", "Procedure date",
          "Intraoperative course", "Intraoperative complications", "Stone sent for analysis",
          "Bladder irrigation after surgery"]),
        ("3. Course",
         ["Postoperative course", "Postoperative complications", "Catheter removal", "Catheter removed on",
          "Planned catheter removal (optional)", "Voiding after catheter removal"]),
        ("4. Recommendations", ["Activity restriction", "Urology visit"]),
    ],
}

# Automatic values. A modifier set here is NOT asked, it is filled in for you.
# By department:
DEPARTMENT_RULES = {
    "4": {"Stay": "Inpatient"},
    "NMC": {"Admission": "Acute"},
}
# By diagnosis:
DIAGNOSIS_RULES = {
    "N10": {"Infection": "With infection"},
}

# Latvian wording used inside the texts.
SIDE_TEXT = {
    "Right": "lb. sānos",
    "Left": "kr. sānos",
    "Both": "abos sānos",
}
DEGREE_TEXT = {
    "Subtotal": "Subtotāla fimoze",
    "Total": "Totāla fimoze",
}
BALANOPOSTHITIS_TEXT = {
    True: "Fimoze, epizodiski balanopostīti.",
    False: "Fimoze.",
}
SURGERY_TEXT = {
    "Circumcision": {"nom": "cirkumcīzija", "gen": "cirkumcīzijas"},
    "Dorsumcision": {"nom": "dorsumcīzija", "gen": "dorsumcīzijas"},
}

# Clinical diagnosis wording for pyelonephritis, by side.
PYELO_TEXT = {
    "Right": "Lb. puses pielonefrīts",
    "Left": "Kr. puses pielonefrīts",
    "Both": "Abpusējs pielonefrīts",
}

# Wording for stone side. {size} is filled in with e.g. "Ø 5 mm".
STONE_WORDS = {
    "Right": {
        "pain": "labajā pusē",
        "epi": "iesprūdušu konkrementu labās ostijas līmenī",
        "dx": "Labās puses konkrements {size} ostijas līmenī",
    },
    "Left": {
        "pain": "kreisajā pusē",
        "epi": "iesprūdušu konkrementu kreisās ostijas līmenī",
        "dx": "Kreisās puses konkrements {size} ostijas līmenī",
    },
    "Both": {
        "pain": "abās pusēs",
        "epi": "iesprūdušus konkrementus abu ostiju līmenī",
        "dx": "Abu pušu konkrementi ostiju līmenī (labajā pusē {size_right}, kreisajā pusē {size_left})",
    },
}

REASON_TEXT = {
    "Not specified": "",
    "Prostate cancer": "CA prostatae.\n",
    "BPH": "BPH.\n",
}

# ----- Wording and tables for the urinary tract infection (N39.0) text -----
# Words that follow "ar sūdzībām par ..." (the form after "par"). Anything you type
# yourself in the "other" box is used exactly as typed.
SYMPTOM_TEXT = {
    "Lumbar pain": "sāpēm jostas apvidū",
    "Flank pain": "sāpēm sānos",
    "Dysuria": "dizūriju",
    "Pollakiuria": "pollakiūriju",
    "Urgency": "neatliekamu urinācijas vajadzību",
    "Fever": "drudzi",
    "Chills": "drebuļiem",
    "Hematuria": "hematūriju",
    "General weakness": "vispārēju nespēku",
}
COMPLEXITY_TEXT = {
    "Complicated": "Komplicēta urīnceļu infekcija",
    "Uncomplicated": "Nekomplicēta urīnceļu infekcija",
    "Not classified": "Urīnceļu infekcija",
}
CONTEXT_TEXT = {
    "Catheter present": " urīnpūšļa katetra klātbūtnes apstākļos",
    "Stents present": " urīnvada stentu klātbūtnes apstākļos",
    "Catheter and stents present": " urīnvada stentu un urīnpūšļa katetra klātbūtnes apstākļos",
    "After urological manipulation": " pēc uroloģiskas manipulācijas",
    "Other / not determined": "",
}
SITE_TEXT = {
    "Lower urinary tract": "Apakšējo urīnceļu infekcija",
    "Upper urinary tract": "Augšējo urīnceļu infekcija",
}
COMPLICATION_TEXT = {
    "Urinary tract obstruction": "urīnceļu obstrukcija",
    "Acute kidney injury": "akūts nieru bojājums",
    "Sepsis": "sepse",
}
CATHETER_INDICATION_TEXT = {
    "Urinary retention": "urīna retence",
    "Urine output monitoring": "urīna diurēzes kontrole",
    "Postoperative care": "pēcoperācijas aprūpe",
    "Securing urine drainage": "urīna atteces nodrošināšana",
}
SPECIMEN_TEXT = {
    "Fresh urine specimen": "Paņemts svaigs urīna paraugs laboratoriskiem izmeklējumiem un urīna "
                            "uzsējumam ar antibakteriālās jutības noteikšanu.",
    "From catheter sampling port": "Paņemts urīna paraugs no katetra paraugu ņemšanas porta "
                                   "laboratoriskiem izmeklējumiem un urīna uzsējumam ar "
                                   "antibakteriālās jutības noteikšanu.",
    "Not taken": "Urīna uzsējumam paraugs netika ņemts.",
}
CULTURE_TEXT = {
    "Result not yet known": "rezultāts izrakstīšanas brīdī vēl nav zināms",
    "Positive": "pozitīvs",
    "Negative": "negatīvs",
    "Possibly contaminated sample": "paraugs iespējami kontaminēts",
    "Not performed": "nav veikts",
}
BLOOD_CULTURE_TEXT = {
    "Pending": "rezultāts vēl nav zināms",
    "Negative": "negatīvs",
    "Positive": "pozitīvs",
}
INFLAMMATORY_TEXT = {
    "Decreasing": "mazinās",
    "Stable": "stabili",
    "Increasing": "pieaug",
    "Not available": "dati nav pieejami",
}
DRAINAGE_TEXT = {
    "Free": "brīva",
    "Reduced": "samazināta",
    "None": "nav",
    "Not documented": "nav dokumentēta",
}
VOIDING_TEXT = {
    "Voids independently": "urinē patstāvīgi",
    "Feeling of incomplete emptying": "nepilnīgas iztukšošanās sajūta",
    "Increased residual urine": "paaugstināts atlieku urīna daudzums",
    "Urinary retention": "urīna retence",
    "Not tested": "nav pārbaudīts",
}
STENT_PLAN_TEXT = {  # "... stentu izņemšana vai maiņa ..."
    "At the outpatient urology review": "plānota ambulatorās urologa vizītes laikā",
    "Not yet scheduled": "tiks plānota pēc urologa norādījumiem",
}
CULTURE_CHECKER_TEXT = {
    "Family doctor": "ģimenes ārsts",
    "Treating urologist": "ārstējošais urologs",
    "Hospital department": "stacionāra nodaļa",
}
WARNING_TEXT = {  # each one is followed by "... gadījumā nekavējoties vērsties NMC"
    "Fever": "paaugstinātas ķermeņa temperatūras",
    "Chills": "drebuļu",
    "Flank or lumbar pain": "sāpju sānos vai jostas apvidū",
    "Urine drainage problems": "urīna atteces traucējumu",
    "Inability to urinate": "nespējas urinēt",
    "Marked hematuria": "izteiktas hematūrijas",
    "General deterioration": "vispārējā stāvokļa pasliktināšanās",
}

# Lab tests for the lab table: name -> (Latvian label, unit)
LAB_TESTS = {
    "CRO": ("CRO", "mg/l"),
    "Leukocytes": ("leikocīti", "×10⁹/l"),
    "Creatinine": ("kreatinīns", "µmol/l"),
    "eGFR": ("GFĀ", "ml/min/1,73 m²"),
}
INFLAMMATORY_TESTS = ("CRO", "Leukocytes")

# Antibiotic table choices.
ROUTE_WORD = {"IV": "intravenozi", "PO": "perorāli", "IM": "intramuskulāri"}
FREQUENCIES = [
    "1 reizi dienā", "2 reizes dienā", "3 reizes dienā", "4 reizes dienā",
    "ik pēc 6 stundām", "ik pēc 8 stundām", "ik pēc 12 stundām", "ik pēc 24 stundām",
]
# form = Tab. / Inj. in the recommendations, nom = name in the recommendations,
# acc = name after "ar ..." / "uz ...", acc_dose = the same with the dose ({dose}).
MEDICATION_CATALOG = {
    "Piperacillin/tazobactam": {
        "form": "Inj.", "nom": "Piperacilīns/tazobaktāms", "acc": "piperacilīnu/tazobaktāmu"},
    "Amoxicillin": {
        "form": "Tab.", "nom": "Amoksicilīns", "acc": "amoksicilīnu"},
    "Biseptol (sulfamethoxazole/trimethoprim)": {
        "form": "Tab.", "nom": "Biseptol",
        "acc": "sulfametoksazola/trimetoprima kombināciju (Biseptol)",
        "acc_dose": "sulfametoksazola/trimetoprima kombināciju (Biseptol {dose})"},
    "Amoxicillin/clavulanate (Amoksiklav)": {
        "form": "Tab.", "nom": "Amoksiklav", "acc": "Amoksiklav"},
    "Ciprofloxacin": {
        "form": "Tab.", "nom": "Ciprofloksacīns", "acc": "ciprofloksacīnu"},
    "Ceftriaxone": {
        "form": "Inj.", "nom": "Ceftriaksons", "acc": "ceftriaksonu"},
}

# ----- Wording for kidney stones (N20) -----
STONE_LOCATION_TEXT = {  # "... konkrementi"
    "Kidney": {"Right": "Labās nieres", "Left": "Kreisās nieres", "Both": "Abu nieru"},
    "Renal pelvis": {"Right": "Labās nieres bļodiņas", "Left": "Kreisās nieres bļodiņas",
                     "Both": "Abu nieru bļodiņu"},
    "Upper ureter": {"Right": "Labā urīnvada augšējās trešdaļas", "Left": "Kreisā urīnvada augšējās trešdaļas",
                     "Both": "Abu urīnvadu augšējās trešdaļas"},
    "Middle ureter": {"Right": "Labā urīnvada vidējās trešdaļas", "Left": "Kreisā urīnvada vidējās trešdaļas",
                      "Both": "Abu urīnvadu vidējās trešdaļas"},
    "Lower ureter": {"Right": "Labā urīnvada apakšējās trešdaļas", "Left": "Kreisā urīnvada apakšējās trešdaļas",
                     "Both": "Abu urīnvadu apakšējās trešdaļas"},
    "Ureteral orifice": {"Right": "Labā urīnvada ostijas", "Left": "Kreisā urīnvada ostijas",
                         "Both": "Abu urīnvadu ostiju"},
}
COLIC_TEXT = {
    "Right": "Labās puses nieres kolika.",
    "Left": "Kreisās puses nieres kolika.",
    "Both": "Abpusēja nieres kolika.",
}
PROCEDURE_SIDE = {"Right": "labās puses", "Left": "kreisās puses", "Both": "abpusēja"}
PROCEDURE_TEXT = {
    "URS": "ureterorenoskopija (URS)",
    "FURS": "fleksibla ureterorenoskopija (FURS)",
    "PCNL": "perkutāna nefrolitotomija (PCNL)",
}
LITHO_TEXT = {  # (after "ar ...", in the list of operations)
    "Laser lithotripsy": ("kontakta lāzerlitotripsiju", "kontakta lāzerlitotripsija"),
    "Pneumatic lithotripsy": ("pneimatisko litotripsiju", "pneimatiskā litotripsija"),
    "Basket extraction": ("konkrementu ekstrakciju", "konkrementu ekstrakcija"),
}
STENT_TEXT = ("DJ stenta ievietošanu", "DJ stenta ievietošana")
STENTS_TEXT = ("DJ stentu ievietošanu", "DJ stentu ievietošana")  # both sides
POSTOP_TEXT = {"Uneventful": "bez komplikācijām", "Satisfactory": "apmierinošs"}
DISCHARGE_STATUS_TEXT = {"Satisfactory": "apmierinošs", "Good": "labs", "Stable": "stabils"}

# ----- Wording for bladder tumours (C67, TURU) -----
VERIFICATION_TEXT = {"Cystoscopy": "cistoskopiski", "Ultrasound": "ultrasonogrāfiski", "CT": "datortomogrāfiski"}
TUMOUR_NUMBER_TEXT = {"Two": "diviem", "Three": "trim", "Multiple": "multipliem"}
TUMOUR_LOCATION_TEXT = {  # "urīnpūšļa ... jaunveidojumi"
    "Left lateral wall": "kreisās sānu sienas",
    "Right lateral wall": "labās sānu sienas",
    "Posterior wall": "mugurējās sienas",
    "Anterior wall": "priekšējās sienas",
    "Dome": "jumta",
    "Trigone": "trīsstūra",
    "Bladder neck": "kakla",
}
TURU_TEXT = {"TURU": "TURU", "Re-TURU": "atkārtota TURU (re-TURU)"}
TURU_SHORT = {"TURU": "TURU", "Re-TURU": "re-TURU"}  # for "St. pēc ..." in the diagnosis
GRADE_TEXT = {"Low grade (LG)": "LG", "High grade (HG)": "HG", "Not stated": ""}
INSTILLATION_TEXT = {"Epirubicin single instillation": "vienreizēja intravezikāla sol. Epirubicin instilācija"}
TEMPORARY_PAPER_TEXT = "Pagaidu izraksts. Histoloģiskās izmeklēšanas rezultāts būs pieejams 2 nedēļu laikā."
# Used everywhere a patient goes home with a urinary catheter.
CATHETER_CARE_LINES = [
    "Urīnpūšļa katetra maiņa ik pēc 4–6 nedēļām ģimenes ārsta praksē.",
    "Urīna savākšanas maisa maiņa ik pēc 3–7 dienām.",
]
C67_POSTOP_TEXT = {"Uneventful": "bez komplikācijām", "Satisfactory": "apmierinoši"}
CATHETER_DAY_TEXT = {"Day 1 after surgery": "1. pēcoperācijas dienā",
                     "Day 2 after surgery": "2. pēcoperācijas dienā",
                     "Day 3 after surgery": "3. pēcoperācijas dienā"}
C67_VOIDING_TEXT = {"Voids independently": "urinē patstāvīgi",
                    "Voids independently, urine clear": "urinē patstāvīgi, urīns dzidrs",
                    "Voids independently, mild hematuria": "urinē patstāvīgi, saglabājas neliela hematūrija",
                    "Voids independently, mild incontinence": "urinē patstāvīgi, saglabājas viegla urīna nesaturēšana",
                    "Urinary retention": "urīna retence"}
WEEKS_TEXT = {"1 week": "1 nedēļu", "2 weeks": "2 nedēļas", "3 weeks": "3 nedēļas", "4 weeks": "4 nedēļas",
              "1 month": "1 mēnesi",
              "6 weeks": "6 nedēļas", "3 months": "3 mēnešus"}

# ----- Wording for prostate cancer (C61) -----
RP_TEXT = {  # (as the operation done, in "St. pēc ...")
    "Open retropubic": ("atvērta retropubiska radikāla prostatektomija", "atvērtas retropubiskas radikālas prostatektomijas"),
    "Laparoscopic": ("laparoskopiska radikāla prostatektomija", "laparoskopiskas radikālas prostatektomijas"),
    "Robot-assisted": ("robotasistēta radikāla prostatektomija", "robotasistētas radikālas prostatektomijas"),
}
LND_TEXT = (" ar iegurņa limfadenektomiju", " un iegurņa limfadenektomijas")
ANESTHESIA_TEXT = {"Spinal": "spinālā anestēzijā", "General": "vispārējā anestēzijā"}
NMC_REASON_TEXT = {  # (after "ar sūdzībām par ...", in the diagnosis)
    "Urinary retention": ("nespēju urinēt", "Akūta urīna retence."),
    "Hematuria": ("asiņu piejaukumu urīnā", "Makrohematūrija."),
    "Pain": (None, "Sāpju sindroms."),
}
CONSULT_TEXT = {"Within 1 week": "1 nedēļas laikā", "In 1–2 weeks": "pēc 1–2 nedēļām", "In 4 weeks": "pēc 4 nedēļām"}
C61_RP_WARNING = ("Ja parādās drudzis, stipras sāpes vēderā, urīna atteces traucējumi pa katetru vai nespēja "
                  "urinēt, izteikta hematūrija, kājas pietūkums vai sāpes kājā, elpas trūkums vai būtiski "
                  "pasliktinās vispārējais stāvoklis, nekavējoties vērsties pēc neatliekamās medicīniskās palīdzības.")

# ----- Wording for kidney cancer (C64) -----
KIDNEY_SIDE_TEXT = {"Right": "labās", "Left": "kreisās"}
KIDNEY_APPROACH_TEXT = {  # (as done, in "St. pēc ...")
    "Open": ("atvērta", "atvērtas"),
    "Laparoscopic": ("laparoskopiska", "laparoskopiskas"),
    "Robot-assisted": ("robotasistēta", "robotasistētas"),
}
RCC_SUBTYPE_TEXT = {"Clear cell": "gaiššūnu karcinoma", "Papillary": "papilāra karcinoma",
                    "Chromophobe": "hromofoba karcinoma"}
RENAL_FUNCTION_TEXT = {
    "Stable": "Nieru funkcija izrakstīšanas brīdī stabila",
    "Mildly impaired, stable": "Izrakstīšanas brīdī nieru funkcija nedaudz pazemināta, stabila",
}
C64_WARNING = ("Ja parādās drudzis, stipras sāpes vēderā vai sānos, izteikta hematūrija, brūces apsārtums vai "
               "izdalījumi no brūces, kājas pietūkums vai sāpes kājā, elpas trūkums vai būtiski pasliktinās "
               "vispārējais stāvoklis, nekavējoties vērsties pēc neatliekamās medicīniskās palīdzības.")

# ----- Wording added in the big update -----
N10_IMAGING_TEXT = {"Ultrasound": "USS nierēm bez urostāzes. ", "CT": "CT vēdera dobumam bez urostāzes. ", "None": ""}
ANESTHESIA_N47_TEXT = {"Spinal": "spinālā anestēzijā", "General": "vispārējā anestēzijā", "Local": "vietējā anestēzijā"}
APPEARANCE_AT_CHANGE_TEXT = {"Yellow, clear": "dzeltena urīna", "Cloudy": "duļķaina urīna",
                             "Blood-tinged": "urīna ar asiņu piejaukumu", "With clots": "urīna ar asins recekļiem"}
# Stone place for N20: diagnosis wording, wording by side, area of pain, impacted (ureter) or not.
STONE_PLACE = {
    "Kidney": {"dx": "nieres kausiņu sistēmā", "both": "nieru kausiņu sistēmā",
               "Right": "labās nieres kausiņu sistēmā", "Left": "kreisās nieres kausiņu sistēmā",
               "Both": "abu nieru kausiņu sistēmā", "area": "jostas apvidū", "impacted": False},
    "Renal pelvis": {"dx": "nieres bļodiņā", "both": "nieru bļodiņās",
                     "Right": "labās nieres bļodiņā", "Left": "kreisās nieres bļodiņā",
                     "Both": "abu nieru bļodiņās", "area": "jostas apvidū", "impacted": False},
    "Upper ureter": {"dx": "urīnvada augšējā trešdaļā", "both": "urīnvadu augšējā trešdaļā",
                     "Right": "labā urīnvada augšējā trešdaļā", "Left": "kreisā urīnvada augšējā trešdaļā",
                     "Both": "abu urīnvadu augšējā trešdaļā", "area": "jostas apvidū", "impacted": True},
    "Middle ureter": {"dx": "urīnvada vidējā trešdaļā", "both": "urīnvadu vidējā trešdaļā",
                      "Right": "labā urīnvada vidējā trešdaļā", "Left": "kreisā urīnvada vidējā trešdaļā",
                      "Both": "abu urīnvadu vidējā trešdaļā", "area": "vēdera lejasdaļā", "impacted": True},
    "Lower ureter": {"dx": "urīnvada apakšējā trešdaļā", "both": "urīnvadu apakšējā trešdaļā",
                     "Right": "labā urīnvada apakšējā trešdaļā", "Left": "kreisā urīnvada apakšējā trešdaļā",
                     "Both": "abu urīnvadu apakšējā trešdaļā", "area": "vēdera lejasdaļā", "impacted": True},
    "Ureteral orifice": {"dx": "ostijas līmenī", "both": "ostiju līmenī",
                         "Right": "labās ostijas līmenī", "Left": "kreisās ostijas līmenī",
                         "Both": "abu ostiju līmenī", "area": "vēdera lejasdaļā", "impacted": True},
}
DRAIN_TEXT = {  # (as done, after "St. pēc ...", "izrakstīts ar ...")
    "DJ stent": {
        "Right": ("labās puses DJ stenta ievietošana", "labās puses DJ stenta ievietošanas", "ar DJ stentu"),
        "Left": ("kreisās puses DJ stenta ievietošana", "kreisās puses DJ stenta ievietošanas", "ar DJ stentu"),
        "Both": ("abpusēja DJ stentu ievietošana", "abpusējas DJ stentu ievietošanas", "ar DJ stentiem"),
    },
    "Percutaneous nephrostomy": {
        "Right": ("labās nieres perkutāna nefrostomija", "labās nieres perkutānas nefrostomijas", "ar nefrostomu"),
        "Left": ("kreisās nieres perkutāna nefrostomija", "kreisās nieres perkutānas nefrostomijas", "ar nefrostomu"),
        "Both": ("abu nieru perkutāna nefrostomija", "abu nieru perkutānas nefrostomijas", "ar nefrostomām"),
    },
}
DEFINITIVE_TEXT = {"URS/FURS": "ureterorenoskopija (URS/FURS)", "PCNL": "perkutāna nefrolitotomija (PCNL)"}
TEMPERATURE_COURSE_TEXT = {"Afebrile": "ķermeņa temperatūra normas robežās",
                           "Fever resolved": "ķermeņa temperatūra normalizējusies"}
DURATION_TEXT = {"Hours": "dažas stundas", "1–2 days": "1–2 dienas", "3–7 days": "3–7 dienas",
                 "Over a week": "ilgāk par nedēļu"}
ONSET_TEXT = {"Sudden": "pēkšņi", "Gradual": "pakāpeniski", "After catheter change": "pēc katetra maiņas",
              "After stent insertion": "pēc stenta ievietošanas",
              "After urological manipulation": "pēc uroloģiskas manipulācijas"}
FEVER_STATUS_TEXT = {"Febrile": "Iestāšanās brīdī febrila ķermeņa temperatūra.",
                     "Subfebrile": "Iestāšanās brīdī subfebrila ķermeņa temperatūra.",
                     "Afebrile": "Iestāšanās brīdī ķermeņa temperatūra normas robežās."}
HISTORY_TEXT = {"Previous UTIs": "iepriekš pārslimotas urīnceļu infekcijas",
                "Recurrent UTIs": "atkārtotas urīnceļu infekcijas", "Nephrolithiasis": "nefrolitiāze",
                "Urinary tract obstruction": "urīnceļu obstrukcija", "Urological surgery": "uroloģiskas operācijas",
                "Diabetes": "cukura diabēts", "Immunosuppression": "imūnsupresija"}
STENT_INDICATION_TEXT = {"Obstruction": "obstrukcijas dēļ", "Ureteral stone": "urīnvada konkrementa dēļ",
                         "Postoperative drainage": "pēcoperācijas drenāžai"}
URINE_APPEARANCE_TEXT = {"Clear": "Urīns dzidrs.", "Cloudy": "Urīns duļķains.",
                         "Blood-tinged": "Urīnā asiņu piejaukums.", "With clots": "Urīnā asins recekļi."}
IMAGING_TEXT = {"Ultrasound": "USS nierēm un urīnpūslim", "CT": "CT vēdera dobumam"}
HYDRO_TEXT = {"Right": "labās puses", "Left": "kreisās puses", "Both": "abpusēja"}

# ----- Wording for urethral stricture (N35) -----
STRICTURE_DX = {"Meatus": "Meatus striktūra", "Penile urethra": "Uretras brīvās daļas striktūra",
                "Bulbar urethra": "Uretras bulbārās daļas striktūra",
                "Membranous (posterior) urethra": "Uretras membranozās daļas striktūra",
                "Not specified": "Uretras striktūra"}
STRICTURE_EPI = {"Meatus": "uretras ārējās atveres (meatus) striktūru",
                 "Penile urethra": "uretras brīvās daļas striktūru",
                 "Bulbar urethra": "uretras bulbārās daļas striktūru",
                 "Membranous (posterior) urethra": "uretras membranozās daļas striktūru",
                 "Not specified": "uretras striktūru"}
STRICTURE_LOC = {"Meatus": "ārējā atverē (meatus)", "Penile urethra": "brīvajā daļā",
                 "Bulbar urethra": "bulbārajā daļā", "Membranous (posterior) urethra": "membranozajā daļā",
                 "Not specified": ""}  # "Uretras ... konstatē striktūru"
COMPENSATION = {"Compensated": ("kompensēta", "kompensētu"), "Subcompensated": ("subkompensēta", "subkompensētu"),
                "Decompensated": ("dekompensēta", "dekompensētu")}
VERIFIED_BY = {"Cystoscopy": "cistoskopiski", "Urethrography": "uretrogrāfiski"}
CATHETER_PLACE_N35 = {"PSKUS outpatient clinic": " P. Stradiņa KUS poliklīnikā",
                      "Family doctor's practice": " ģimenes ārsta praksē", "Not stated": ""}
STRICTURE_OP = {  # (as done, after "st. pēc")
    "Meatotomy": ("meatotomija", "meatotomijas"),
    "Meatoplasty": ("meatoplastika", "meatoplastikas"),
    "Urethral dilation": ("uretras dilatācija", "uretras dilatācijas"),
    "Internal urethrotomy (DVIU)": ("optiskā uretrotomija", "optiskās uretrotomijas"),
    "Urethroplasty – excision and primary anastomosis": (
        "uretroplastika ar striktūras ekscīziju un primāru anastomozi",
        "uretroplastikas ar striktūras ekscīziju un primāru anastomozi"),
    "Urethroplasty – buccal mucosa graft": ("uretroplastika ar vaiga gļotādas transplantātu",
                                            "uretroplastikas ar vaiga gļotādas transplantātu"),
}
STRICTURE_CAUSE_DX = {"Iatrogenic (catheter / transurethral surgery)": "Jatrogēna ģenēze.",
                      "Lichen sclerosus": "Lichen sclerosus.", "Trauma": "Posttraumatiska ģenēze.",
                      "Infection": "Postinfekcioza ģenēze."}
N35_ANESTHESIA = {"Spinal": "spinālā anestēzijā", "General": "vispārējā anestēzijā", "Local": "vietējā anestēzijā"}
N35_POSTOP = {"Uneventful": "bez sarežģījumiem", "Satisfactory": "apmierinošs"}
N35_VOIDING = {"Good stream": "urinē patstāvīgi ar labu strūklu", "Weak stream": "urinē patstāvīgi, strūkla pavājināta",
               "Urinary retention": "urīna retence"}
URETHROPLASTY = ("Urethroplasty – excision and primary anastomosis", "Urethroplasty – buccal mucosa graft")
ENDOSCOPIC = ("Urethral dilation", "Internal urethrotomy (DVIU)")

# ----- Wording for nephrostomy change (Z43.6) -----
SUPERVISOR_TEXT = {"Oncologist": "onkologa", "Urologist": "urologa", "Family doctor": "ģimenes ārsta"}
NEPH_SIDE = {  # (adjective, "labās nefrostomas maiņa", operation name)
    "Right": ("labās puses", "labās nefrostomas maiņa", "Labās puses nefrostomas maiņa"),
    "Left": ("kreisās puses", "kreisās nefrostomas maiņa", "Kreisās puses nefrostomas maiņa"),
    "Both": ("abpusējai", "abu nefrostomu maiņa", "Abpusēja nefrostomu maiņa"),
}

# ----- Wording for BPH / TURP (N40) -----
BPH_INDICATION_TEXT = {
    "Bothersome LUTS despite medication": "izteikti apakšējo urīnceļu simptomi, neefektīva medikamentoza terapija",
    "Recurrent urinary retention": "atkārtota urīna retence",
    "Recurrent UTIs": "atkārtotas urīnceļu infekcijas",
    "Recurrent hematuria": "atkārtota makrohematūrija",
    "Bladder stones": "urīnpūšļa akmeņi",
    "Upper tract dilatation / renal impairment": "augšējo urīnceļu dilatācija, nieru funkcijas pasliktināšanās",
}
BPH_COMPLICATED = ("Recurrent urinary retention", "Recurrent UTIs", "Recurrent hematuria", "Bladder stones",
                   "Upper tract dilatation / renal impairment")
TURP_TECHNIQUE_TEXT = {"Monopolar": "monopolāra", "Bipolar": "bipolāra"}

# Latvian words that change with the patient's sex.
# In the templates below they appear as {Pacients}, {iestajies}, etc.
SEX_FORMS = {
    "Woman": {
        "Pacients": "Paciente",
        "pacients": "paciente",
        "pacienta": "pacientes",
        "iestajies": "iestājusies",
        "izmekleta": "izmeklēta",
        "stacionets": "stacionēta",
        "gatavots": "gatavota",
        "izrakstits": "izrakstīta",
    },
    "Man": {
        "Pacients": "Pacients",
        "pacients": "pacients",
        "pacienta": "pacienta",
        "iestajies": "iestājies",
        "izmekleta": "izmeklēts",
        "stacionets": "stacionēts",
        "gatavots": "gatavots",
        "izrakstits": "izrakstīts",
    },
}

# Recommendation blocks you can drag into the recommendations, per diagnosis.
#   group = heading in the side panel (keep blocks of one group together)
#   title = what you see on the block
#   text  = the line that is added to the recommendations
#   note  = clinical remarks, shown when you hover a block and in "Clinical notes"
#           (they are never put in the document)
#   kind  = "avoid" gives the block a red edge (for things not to use)
G_ORAL = "Antibiotics - stable patient, oral"
G_IV = "Antibiotics - hospitalized patient, IV"
G_PAIN = "Pain and fever"
G_AVOID = "Not recommended"
G_PREVENT = "Prevention"

BLOCKS = {
    "N10": [
        {
            "id": "cipro",
            "group": G_ORAL,
            "title": "Ciprofloxacin 500 mg q12h, 7 days",
            "text": "Tab. Ciprofloxacin 500 mg pa 1 tabletei 2 reizes dienā 7 dienas.",
            "note": "500 mg PO every 12 hours, typically 7 days if susceptible and clinically "
                    "appropriate. Useful when local resistance is acceptable and the isolate is "
                    "susceptible. Consider tendon, CNS, QT, vascular and other fluoroquinolone risks.",
        },
        {
            "id": "tmp-smx",
            "group": G_ORAL,
            "title": "TMP-SMX 160/800 mg q12h, 14 days",
            "text": "Tab. Trimetoprims/Sulfametoksazols 160/800 mg pa 1 tabletei 2 reizes dienā 14 dienas.",
            "note": "160/800 mg PO every 12 hours, typically 14 days in traditional regimens. "
                    "Preferably use when susceptibility is known. Monitor potassium and creatinine, "
                    "especially in renal impairment or with ACE inhibitors, ARBs, or spironolactone.",
        },
        {
            "id": "amoxiclav",
            "group": G_ORAL,
            "title": "Amoxicillin/clavulanate 875/125 mg q12h, 5 days",
            "text": "Tab. Amoksiklav 875 mg/125 mg pa 1 tabletei 2 reizes dienā 5 dienas.",
            "note": "Kept from your original pyelonephritis recommendations.",
        },
        {
            "id": "ceftriaxone",
            "group": G_IV,
            "title": "Ceftriaxone 1-2 g q24h",
            "text": "Inj. Ceftriaxone 1-2 g i/v 1 reizi dienā.",
            "note": "Common initial choice, 1-2 g IV every 24 hours. Often appropriate for "
                    "community-acquired pyelonephritis without strong risk factors for resistant "
                    "organisms. Usually no renal dose adjustment is needed in isolated renal impairment.",
        },
        {
            "id": "pip-tazo",
            "group": G_IV,
            "title": "Piperacillin/tazobactam 4.5 g q8h",
            "text": "Inj. Piperacillin/Tazobactam 4.5 g i/v ik pēc 8 stundām.",
            "note": "4.5 g IV every 8 hours, an example of a conventional regimen for complicated "
                    "infection; severe infection protocols may use every 6 hours or prolonged "
                    "infusion. Adjust for renal function and follow the hospital protocol. Consider "
                    "when broader Gram-negative coverage, including Pseudomonas, is needed.",
        },
        {
            "id": "paracetamol",
            "group": G_PAIN,
            "title": "Paracetamol 500-1,000 mg q6-8h as needed",
            "text": "Tab. Paracetamol 500 mg pa 1-2 tabletēm ik pēc 6-8 stundām pie sāpēm vai "
                    "paaugstinātas temperatūras pēc nepieciešamības, ne vairāk kā 3000 mg dienā.",
            "note": "First-line symptomatic option. A practical adult regimen for fever and pain. "
                    "Keep the total at or below 3,000 mg/day for routine self-management; use a "
                    "lower maximum or longer intervals in liver disease, low body weight, "
                    "malnutrition or heavy alcohol use. Check combination products to avoid "
                    "duplicate paracetamol.",
        },
        {
            "id": "ultracod",
            "group": G_PAIN,
            "title": "Ultracod 500 mg/30 mg tablets",
            "text": "Tab. Ultracod 500 mg/30 mg pa 1 tabletei līdz 4 r./dienā pēc nepieciešamības pie sāpēm.",
            "note": "Paracetamol 500 mg + codeine phosphate 30 mg per tablet. Used for moderate to "
                    "severe pain when simpler analgesics are inadequate. The current Latvian product "
                    "information specifies a minimum 6-hour interval and treatment for no longer "
                    "than 3 days without medical reassessment.",
        },
        {
            "id": "nsaid",
            "group": G_AVOID,
            "kind": "avoid",
            "title": "NSAIDs (ibuprofen etc.) - avoid",
            "text": "Nelietot nefrotoksiskus medikamentus, nesteroīdos pretiekaisuma līdzekļus "
                    "(Ibuprofēns, Ibumetins, Diklofenaks u.c.).",
            "note": "Ibuprofen and other NSAIDs: generally avoid in acute pyelonephritis when renal "
                    "risk is present. NSAIDs can reduce pain and fever, but may precipitate or worsen "
                    "acute kidney injury, particularly with dehydration, sepsis, chronic kidney "
                    "disease or obstruction. For a patient with impaired renal function, paracetamol "
                    "is generally preferred.",
        },
        {
            "id": "canephron",
            "group": G_PREVENT,
            "title": "Canephron",
            "text": "Tab. Canephron pa 2 tabletēm 3 reizes dienā (ilgstoši).",
            "note": "Reasonable adjunct. Available in Latvia. Contains centaury herb, lovage root "
                    "and rosemary leaf. It is traditionally used as supportive treatment for mild "
                    "lower urinary tract inflammatory conditions and for kidney-stone prevention. "
                    "Clinical evidence also exists for acute uncomplicated cystitis in selected "
                    "women, but that is not evidence of efficacy in kidney infection.",
        },
        {
            "id": "herbal",
            "group": G_PREVENT,
            "title": "Urological herbal teas / diuretic preparations",
            "text": "Urīnceļu veselībai paredzētas augu tējas un diurētiskie augu preparāti "
                    "(piemēram, ar bērza lapām) pēc nepieciešamības.",
            "note": "Products containing birch leaf, goldenrod or other traditional urinary herbs "
                    "are marketed for urinary health. Evidence for treating bacterial infection "
                    "is insufficient.",
        },
    ],
}

# A warning appears in the side panel when all the listed blocks are chosen together.
BLOCK_WARNINGS = {
    "N10": [
        {
            "ids": ["paracetamol", "ultracod"],
            "text": "Paracetamol and Ultracod both contain paracetamol. Check the total daily "
                    "dose (keep it at or below 3,000 mg/day).",
        },
    ],
}

# N39.0 uses the pain/fever and supportive blocks of N10; its antibiotics come from the
# antibiotics table instead of blocks.
BLOCKS["N39.0"] = [b for b in BLOCKS["N10"] if b["group"] in (G_PAIN, G_PREVENT)]
BLOCK_WARNINGS["N39.0"] = BLOCK_WARNINGS["N10"]

# N20 kidney stones: tamsulosin and Dolmen from your NMC text, plus pain and supportive blocks.
G_ALPHA = "Alpha-blocker"
BLOCKS["N20"] = [
    {
        "id": "tamsulosin",
        "group": G_ALPHA,
        "title": "Tamsulosin 0.4 mg once daily, 7 days",
        "text": "Caps. Tamsulosini 0.4mg x 1 p/o, 7 dienas (izrakstīts e-recepšu veidā)",
        "note": "From your NMC kidney stone text. Alpha-blocker, used for medical expulsive therapy "
                "of distal ureteral stones and for stent-related symptoms.",
    },
    {
        "id": "dolmen",
        "group": G_PAIN,
        "title": "Dolmen (dexketoprofen) 25 mg, up to 3 a day for pain",
        "text": "Tab. Dolmeni 25mg x 1 p/o (līdz 3 tabl. DN), pie sāpēm",
        "note": "From your NMC kidney stone text. Dexketoprofen is an NSAID: check renal function "
                "and contraindications.",
    },
] + [b for b in BLOCKS["N10"] if b["group"] in (G_PAIN, G_PREVENT)]
BLOCK_WARNINGS["N20"] = BLOCK_WARNINGS["N10"]

# C61 prostate cancer blocks.
G_VTE = "Thromboprophylaxis"
G_ED = "Erectile function"
G_BOWEL = "Bowel"
FRAXIPARINE_BLOCK = {
    "id": "fraxiparine",
    "group": G_VTE,
    "title": "Fraxiparine 0.3 ml s.c. once daily, 10 days",
    "text": "Inj. Fraxiparine 0,3 ml (2850 SV anti-Xa) s/c 1 reizi dienā 10 dienas no operācijas dienas.",
    "note": "Department protocol: 10 days. Nadroparin label dose for moderate-risk surgery: 0.3 ml "
            "(2,850 anti-Xa IU) once daily for at least 7 days or until mobile. Reduce the dose by 25-33% if "
            "creatinine clearance is below 30 ml/min. Weigh bleeding risk.",
}
FRAXIPARINE_4W_BLOCK = {
    "id": "fraxiparine-4w",
    "group": G_VTE,
    "title": "Fraxiparine 0.3 ml s.c. once daily, 4 weeks (extended)",
    "text": "Inj. Fraxiparine 0,3 ml (2850 SV anti-Xa) s/c 1 reizi dienā 4 nedēļas no operācijas dienas.",
    "note": "Extended prophylaxis after major cancer surgery. EAU thromboprophylaxis guidance is risk-based: "
            "when drug prophylaxis is indicated it is continued for about 4 weeks; usually not needed after "
            "robot-assisted RP without lymph node dissection. Reduce the dose by 25-33% if creatinine clearance "
            "is below 30 ml/min. Weigh bleeding risk.",
}
G_SPASM = "Bladder spasm"
SOLIFENACIN_BLOCK = {
    "id": "solifenacin",
    "group": G_SPASM,
    "title": "Solifenacin 5 mg once daily for bladder spasm",
    "text": "Tab. Solifenacini 5 mg 1 reizi dienā pie urīnpūšļa spazmām pēc nepieciešamības.",
    "note": "Antimuscarinic for bladder spasm or catheter-related discomfort after transurethral surgery. "
            "Avoid in urinary retention and narrow-angle glaucoma; caution in the elderly.",
}
BLOCKS["C61"] = [
    dict(BLOCKS["N20"][0], note="Alpha-blocker. Before a trial without catheter after acute urinary retention, "
                                "an alpha-blocker improves the chance of voiding (NICE CG97)."),
] + [b for b in BLOCKS["N10"] if b["group"] == G_PAIN] + [
    FRAXIPARINE_BLOCK,
    FRAXIPARINE_4W_BLOCK,
    {
        "id": "tadalafil",
        "group": G_ED,
        "title": "Tadalafil 5 mg once daily",
        "text": "Tab. Tadalafil 5 mg 1 reizi dienā.",
        "note": "PDE5 inhibitors are first-line for erectile dysfunction after RP. Evidence that scheduled "
                "'penile rehabilitation' improves unassisted recovery is weak; on-demand use is an alternative.",
    },
    {
        "id": "macrogol",
        "group": G_BOWEL,
        "title": "Macrogol (Forlax) 10 g as needed",
        "text": "Pulv. Macrogol (Forlax) 10 g 1–2 reizes dienā pēc nepieciešamības aizcietējumu profilaksei.",
        "note": "To avoid straining after TURP or radical prostatectomy. First choice: a Cochrane review found "
                "macrogol (PEG) gave more bowel movements, softer stools and less need for extra laxatives "
                "than lactulose.",
    },
    {
        "id": "duphalac",
        "group": G_BOWEL,
        "title": "Duphalac (lactulose) 15-30 ml as needed",
        "text": "Sol. Duphalac 15–30 ml 1 reizi dienā pēc nepieciešamības aizcietējumu profilaksei.",
        "note": "Alternative to macrogol. Lactulose works less well than macrogol in trials and often causes "
                "bloating and flatulence.",
    },
    SOLIFENACIN_BLOCK,
]
BLOCK_WARNINGS["C61"] = BLOCK_WARNINGS["N10"]

# C64 kidney cancer blocks (no NSAIDs: they are excluded in the text).
_c61_blocks = {b["id"]: b for b in BLOCKS["C61"]}
BLOCKS["C64"] = [_c61_blocks["paracetamol"], _c61_blocks["ultracod"],
                 FRAXIPARINE_BLOCK,
                 dict(FRAXIPARINE_4W_BLOCK,
                      note="Extended prophylaxis after major cancer surgery. EAU thromboprophylaxis guidance is "
                           "risk-based: for nephrectomy the benefit of extended (about 4 weeks) prophylaxis "
                           "depends on patient risk and surgical approach. Reduce the dose by 25-33% if creatinine "
                           "clearance is below 30 ml/min, which matters more after nephrectomy."),
                 _c61_blocks["macrogol"], _c61_blocks["duphalac"]]
BLOCK_WARNINGS["C64"] = BLOCK_WARNINGS["N10"]

# Medication blocks for the other operations (Fraxiparine 10 days after every operation).
G_ABX = "Antibiotics"
G_TOPICAL = "Topical treatment"
BLOCKS["C67"] = [_c61_blocks["paracetamol"], _c61_blocks["ultracod"], SOLIFENACIN_BLOCK, FRAXIPARINE_BLOCK,
                 _c61_blocks["macrogol"], _c61_blocks["duphalac"]]
BLOCK_WARNINGS["C67"] = BLOCK_WARNINGS["N10"]
BLOCKS["N20"] = BLOCKS["N20"] + [FRAXIPARINE_BLOCK]
BLOCKS["N47"] = [
    _c61_blocks["paracetamol"], _c61_blocks["ultracod"],
    {
        "id": "amoxiclav-14",
        "group": G_ABX,
        "title": "Amoksiklav 875/125 mg twice daily, 14 days",
        "text": "Tab. Amoksiklav 875/125 mg pa 1 tabletei 2 reizes dienā 14 dienas.",
        "note": "From your original phimosis text. Prophylactic antibiotics are generally not needed after clean "
                "genital surgery; use it when there is infection, for example balanoposthitis.",
    },
    {
        "id": "betamethasone",
        "group": G_TOPICAL,
        "title": "Betamethasone 0.05% cream twice daily, 4-6 weeks",
        "text": "Ung. Betametazons 0,05% plānā kārtā uz priekšādiņas 2 reizes dienā 4–6 nedēļas.",
        "note": "Topical corticosteroid is the usual conservative treatment for phimosis without scarring; "
                "less effective with scarring (lichen sclerosus). Not during an acute infection.",
    },
    FRAXIPARINE_BLOCK,
    {"id": "amoxiclav-7-pf", "group": G_ABX, "title": "Amoksiklav 875/125 mg twice daily, 7 days",
     "text": "Tab. Amoksiklav 875/125 mg pa 1 tabletei 2 reizes dienā 7 dienas.",
     "note": "From your paraphimosis example. Routine antibiotics are not needed after reduction or a clean dorsal "
             "slit; use them with balanoposthitis, skin breakdown or necrosis."},
    {"id": "diclofenac-pf", "group": G_PAIN, "title": "Diclofenac 100 mg prolonged release once daily as needed",
     "text": "Tab. Diclofenaci 100 mg (ilgstošās darbības) 1 reizi dienā pēc ēšanas sāpju gadījumā.",
     "note": "From your example, where it was 100 mg 1–2 times daily – changed because the EU maximum is 150 mg/day. "
             "NSAID: avoid in renal impairment, GI bleeding risk, heart failure or anticoagulation; shortest course."},
]
BLOCK_WARNINGS["N47"] = BLOCK_WARNINGS["N10"]
BLOCK_WARNINGS["N10"].append({"ids": ["amoxiclav-14", "amoxiclav-7-pf"],
                              "text": "Two Amoksiklav courses are chosen: keep one."})
G_SKIN = "Lichen sclerosus"
G_MOUTH = "Graft donor site"
BLOCKS["N35"] = [
    _c61_blocks["paracetamol"], _c61_blocks["ultracod"], SOLIFENACIN_BLOCK,
    {
        "id": "clobetasol",
        "group": G_SKIN,
        "title": "Clobetasol 0.05% ointment, 3-month tapering course",
        "text": "Ung. Clobetasoli propionas 0,05% plānā kārtā uz skartās vietas 1 reizi dienā 1 mēnesi, tad katru "
                "otro dienu 1 mēnesi, tad 2 reizes nedēļā 1 mēnesi.",
        "note": "Standard treatment for genital lichen sclerosus (potent topical corticosteroid). Lichen sclerosus "
                "is a major cause of meatal and penile strictures and needs long-term follow-up.",
    },
    {
        "id": "chlorhexidine",
        "group": G_MOUTH,
        "title": "Chlorhexidine 0.12% mouthwash, 7 days",
        "text": "Sol. Chlorhexidini 0,12% mutes skalošanai 3 reizes dienā pēc ēšanas 7 dienas.",
        "note": "Oral care of the buccal mucosa donor site after graft urethroplasty.",
    },
    dict(FRAXIPARINE_BLOCK, note=FRAXIPARINE_BLOCK["note"] + " Mainly relevant after urethroplasty; usually not "
                                                            "needed after short endoscopic procedures."),
    _c61_blocks["macrogol"],
]
BLOCK_WARNINGS["N35"] = BLOCK_WARNINGS["N10"]
BLOCKS["Z43.6"] = [_c61_blocks["paracetamol"], _c61_blocks["ultracod"]]
BLOCK_WARNINGS["Z43.6"] = BLOCK_WARNINGS["N10"]



# Fraxiparine 0,4 ml and Urorec (added as options for discharge).
FRAXIPARINE_04_BLOCK = {
    "id": "fraxiparine-04",
    "group": G_VTE,
    "title": "Fraxiparine 0.4 ml s.c. once daily, 10 days",
    "text": "Inj. Fraxiparine 0,4 ml (3800 SV anti-Xa) s/c 1 reizi dienā 10 dienas no operācijas dienas.",
    "note": "Nadroparin 3,800 anti-Xa IU. Higher dose than the standard 0.3 ml: in the label it corresponds to "
            "weight-adjusted high-risk prophylaxis (about 38 IU/kg, i.e. patients of about 70 kg or more). Reduce the "
            "dose by 25-33% if creatinine clearance is below 30 ml/min. Weigh bleeding risk (after TURP: hematuria).",
}
UROREC_BLOCK = {
    "id": "urorec",
    "group": G_ALPHA,
    "title": "Urorec (silodosin) 8 mg once daily",
    "text": "Caps. Urorec 8 mg 1 reizi dienā ēdienreizes laikā līdz kontrolei pie urologa.",
    "note": "Silodosin, a uroselective alpha-1A blocker. Can ease voiding after catheter removal. Common: retrograde "
            "ejaculation; dizziness and orthostatic hypotension (take with a meal); dose 4 mg if moderate renal "
            "impairment. Usually not needed long-term after a successful TURP.",
}
for _code, _blocks in BLOCKS.items():  # offer 0,4 ml wherever 0,3 ml is offered
    _ids = [b["id"] for b in _blocks]
    if "fraxiparine" in _ids and "fraxiparine-04" not in _ids:
        _blocks.insert(_ids.index("fraxiparine") + 1, FRAXIPARINE_04_BLOCK)
BLOCKS["C61"].insert(1, UROREC_BLOCK)
BLOCKS["N40"] = [UROREC_BLOCK, _c61_blocks["paracetamol"], _c61_blocks["ultracod"], SOLIFENACIN_BLOCK,
                 FRAXIPARINE_BLOCK, FRAXIPARINE_04_BLOCK, _c61_blocks["macrogol"], _c61_blocks["duphalac"]]
BLOCK_WARNINGS["N40"] = BLOCK_WARNINGS["N10"]
IBUPROFEN_BLOCK = {
    "id": "ibuprofen", "group": G_PAIN, "title": "Ibuprofen 400 mg up to 3 times daily, max 5 days",
    "text": "Tab. Ibuprofen 400 mg 1 tablete līdz 3 reizēm dienā pēc ēšanas pēc nepieciešamības, ne ilgāk kā 5 dienas.",
    "note": "Helps scrotal pain and swelling after scrotal surgery. Avoid with renal impairment, peptic ulcer, "
            "heart failure or anticoagulants, and in the elderly prefer paracetamol.",
}
BLOCKS["N43"] = [_c61_blocks["paracetamol"], IBUPROFEN_BLOCK, _c61_blocks["ultracod"],
                 FRAXIPARINE_BLOCK, FRAXIPARINE_04_BLOCK]
BLOCK_WARNINGS["N43"] = BLOCK_WARNINGS["N10"]
AMOXICLAV_10_BLOCK = {
    "id": "amoxiclav-10", "group": G_ABX, "title": "Amoksiklav 875/125 mg twice daily, 10 days",
    "text": "Tab. Amoksiklav 875/125 mg pa 1 tabletei 2 reizes dienā p/o 10 dienas.",
    "note": "From your prostatectomy example. EAU: antibiotic prophylaxis is perioperative only; a course after "
            "discharge is not routinely needed – use it for an infection or if it is the department protocol.",
}
TADALAFIL_QUARTER_BLOCK = {
    "id": "tadalafil-quarter", "group": G_ED, "title": "Tadalafil 20 mg, 1/4 tablet at night",
    "text": "Tab. Tadalafil 20 mg pa 1/4 daļai tabletes uz nakti.",
    "note": "Same dose as Tadalafil 5 mg daily (penile rehabilitation after nerve-sparing RP). EAU: evidence that "
            "rehabilitation improves spontaneous erections is limited. Not with nitrates.",
}
BLOCKS["C61"] += [AMOXICLAV_10_BLOCK, TADALAFIL_QUARTER_BLOCK]
TAMSULOSIN_STENT_BLOCK = {
    "id": "tamsulosin-stent", "group": G_ALPHA, "title": "Tamsulosin 0.4 mg once daily until stent removal",
    "text": "Caps. Tamsulosini 0,4 mg 1 reizi dienā līdz DJ stenta izņemšanai (stenta radīto simptomu mazināšanai).",
    "note": "Alpha-blockers reduce DJ stent discomfort (urgency, flank pain on voiding); EAU urolithiasis guideline "
            "supports them for stent-related symptoms. Warn about orthostatic dizziness.",
}
BLOCKS["N13.0"] = [_c61_blocks["paracetamol"], _c61_blocks["ultracod"], TAMSULOSIN_STENT_BLOCK, SOLIFENACIN_BLOCK,
                   FRAXIPARINE_BLOCK, FRAXIPARINE_04_BLOCK, _c61_blocks["macrogol"]]
BLOCK_WARNINGS["N13.0"] = BLOCK_WARNINGS["N10"]
BLOCKS["D29.4"] = [_c61_blocks["paracetamol"], IBUPROFEN_BLOCK, _c61_blocks["ultracod"],
                   dict(AMOXICLAV_10_BLOCK, note="From your scrotal mass example. Clean scrotal surgery rarely needs "
                        "antibiotics (EAU / German S3 guideline: low evidence); keep it for an inflamed or infected "
                        "cyst or wound.")]
BLOCK_WARNINGS["D29.4"] = BLOCK_WARNINGS["N10"]
BLOCKS["N21.0"] = [_c61_blocks["paracetamol"], _c61_blocks["ultracod"], SOLIFENACIN_BLOCK,
                   dict(AMOXICLAV_10_BLOCK, note="Only for a positive urine culture or clinical infection, ideally "
                        "chosen by the culture result. Bladder stones are often infected (struvite)."),
                   _c61_blocks["macrogol"]]
BLOCK_WARNINGS["N21.0"] = BLOCK_WARNINGS["N10"]
RETENTION_ALPHA_NOTE = ("Cochrane review: an alpha-blocker started 1–3 days before the trial without catheter raises "
                        "the success rate (about 60% vs 38%) and lowers re-retention. Warn about dizziness / "
                        "orthostatic hypotension and tell the eye surgeon before cataract surgery (IFIS).")
BLOCKS["R33"] = [
    {"id": "tamsulosin-twoc", "group": G_ALPHA, "title": "Tamsulosin 0.4 mg once daily (before and after TWOC)",
     "text": "Caps. Tamsulosin 0,4 mg 1 reizi dienā pēc brokastīm – sākt tūlīt, turpināt arī pēc katetra evakuācijas.",
     "note": RETENTION_ALPHA_NOTE},
    {"id": "alfuzosin", "group": G_ALPHA, "title": "Alfuzosin 10 mg once daily (prolonged release)",
     "text": "Tab. Alfuzosin 10 mg (ilgstošās darbības) 1 reizi dienā pēc vakariņām.",
     "note": RETENTION_ALPHA_NOTE + " Alfuzosin is the best studied drug for TWOC."},
    UROREC_BLOCK,
    {"id": "dutasteride-tamsulosin", "group": G_ALPHA, "title": "Dutasteride 0.5 mg / tamsulosin 0.4 mg once daily",
     "text": "Caps. Dutasteridum 0,5 mg / Tamsulosinum 0,4 mg 1 reizi dienā p/o (ilgstoši).",
     "note": "EAU (strong): combination therapy for men with an enlarged prostate (> ~40 ml) reduces the long-term "
             "risk of AUR and surgery; the 5-ARI effect needs 6+ months. It halves PSA – double the PSA value when "
             "interpreting it."},
    {"id": "finasteride", "group": G_ALPHA, "title": "Finasteride 5 mg once daily (add to an alpha-blocker)",
     "text": "Tab. Finasteride 5 mg 1 reizi dienā p/o (ilgstoši).",
     "note": "5-ARI for a large prostate: reduces AUR risk and need for surgery long-term (EAU strong); halves PSA."},
    _c61_blocks["paracetamol"], _c61_blocks["macrogol"],
    {"id": "amoxiclav-7", "group": G_ABX, "title": "Amoksiklav 875/125 mg twice daily, 7 days (UTI)",
     "text": "Tab. Amoksiklav 875/125 mg pa 1 tabletei 2 reizes dienā 7 dienas; terapiju koriģēt pēc urīna "
             "uzsējuma rezultāta.",
     "note": "Only with a UTI. EAU: catheter-associated UTI – 7 days when symptoms settle quickly, 10–14 days if slow; "
             "take the culture from the new catheter before starting."},
    {"id": "cefuroxime-7", "group": G_ABX, "title": "Cefuroxime 500 mg twice daily, 7 days (UTI)",
     "text": "Tab. Cefuroxime 500 mg pa 1 tabletei 2 reizes dienā 7 dienas; terapiju koriģēt pēc urīna uzsējuma "
             "rezultāta.",
     "note": "Alternative oral option for a UTI with retention; adjust to the culture."},
    dict({b["id"]: b for b in BLOCKS["N10"]}["cipro"],
         note="Fluoroquinolones: EMA restricts their use (tendon, nerve, aortic and QT risks) – only if other options "
              "are unsuitable and local resistance is acceptable; adjust to the culture."),
]
BLOCK_WARNINGS["R33"] = BLOCK_WARNINGS["N10"]
BLOCK_WARNINGS["N10"].extend([
    {"ids": ["tamsulosin-twoc", "alfuzosin"], "text": "Two alpha-blockers are chosen: keep one."},
    {"ids": ["tamsulosin-twoc", "urorec"], "text": "Two alpha-blockers are chosen: keep one."},
    {"ids": ["alfuzosin", "urorec"], "text": "Two alpha-blockers are chosen: keep one."},
    {"ids": ["dutasteride-tamsulosin", "tamsulosin-twoc"], "text": "The combination capsule already contains tamsulosin."},
    {"ids": ["dutasteride-tamsulosin", "finasteride"], "text": "Two 5-ARIs are chosen: keep one."},
])

# R31 hematuria blocks.
G_BLEED = "Bleeding"
G_RADCYST = "Radiation cystitis"
_r33_blocks = {b["id"]: b for b in BLOCKS["R33"]}
BLOCKS["R31"] = [
    {"id": "tranexamic", "group": G_BLEED, "title": "Tranexamic acid 1 g three times daily, 3–5 days",
     "text": "Tab. Acidum tranexamicum 500 mg pa 2 tabletēm (1 g) 3 reizes dienā 3–5 dienas.",
     "note": "SmPC dose 1–1.5 g 2–4 times daily for a few days. Evidence in hematuria is weak and mixed (no clear "
             "benefit in trials after TURP/TURBT). Not with active thromboembolism or severe renal failure; reduce "
             "the dose if creatinine is raised (120–249 µmol/l: 15 mg/kg twice daily; 250–500: 15 mg/kg once "
             "daily). SmPC warning: with upper urinary tract bleeding clots may obstruct the ureter. Check that a "
             "tablet form is available in Latvian pharmacies."},
    {"id": "finasteride-hem", "group": G_BLEED, "title": "Finasteride 5 mg once daily (recurrent prostatic bleeding)",
     "text": "Tab. Finasteride 5 mg 1 reizi dienā p/o (ilgstoši).",
     "note": "Recurrent bleeding from BPH: small studies show fewer recurrences (effect over weeks to months). Not "
             "an EAU recommendation; exclude a tumour first. Halves PSA – double the PSA value when interpreting it."},
    _c61_blocks["paracetamol"], _c61_blocks["macrogol"], SOLIFENACIN_BLOCK,
    {"id": "fosfomycin-3g", "group": G_ABX, "title": "Fosfomycin 3 g single dose (women, cystitis)",
     "text": "Pulv. Fosfomycin 3 g vienreizēji, izšķīdinot ūdenī, vakarā pēc urīnpūšļa iztukšošanas.",
     "note": "EAU first-line for uncomplicated cystitis in women (hemorrhagic cystitis). Not for men, fever or upper "
             "tract signs – then a 7-day course adjusted to the culture."},
    _r33_blocks["amoxiclav-7"], _r33_blocks["cefuroxime-7"], _r33_blocks["cipro"],
    dict(FRAXIPARINE_BLOCK,
         note="Only once the urine is clear (e.g. after TURU or cystoscopy with coagulation): LMWH adds to the "
              "bleeding risk. Department protocol 10 days; reduce the dose if creatinine clearance is below 30 ml/min."),
    {"id": "hbo", "group": G_RADCYST, "title": "Consider hyperbaric oxygen (radiation cystitis)",
     "text": "Atkārtotas hematūrijas gadījumā apsvērt hiperbarisko oksigenācijas terapiju staru cistīta ārstēšanai "
             "(lems ārstējošais urologs).",
     "note": "Hyperbaric oxygen for hemorrhagic radiation cystitis: usually 20–40 sessions; bleeding stops in most "
             "patients in cohort series, and an RCT (RICH-ART) showed better symptoms in late radiation cystitis. "
             "Check where it is available in Latvia."},
]
BLOCK_WARNINGS["R31"] = BLOCK_WARNINGS["N10"]
_abx = ["fosfomycin-3g", "amoxiclav-7", "cefuroxime-7", "cipro"]
BLOCK_WARNINGS["N10"].extend({"ids": [a, b], "text": "Two antibiotics are chosen: keep one."}
                             for i, a in enumerate(_abx) for b in _abx[i + 1:])

# N30.0 acute cystitis blocks. Pivmecillinam and nitroxoline (EAU first line) are not on the Latvian market as far
# as Latvian outpatient sources say; the nitrofuran used in Latvia is furazidine (Furamag, Olainfarm).
G_CYST_FIRST = "Antibiotics – first line (EAU)"
G_CYST_ALT = "Antibiotics – alternatives / pregnancy"
G_CYST_MEN = "Antibiotics – men (7 days)"
G_CYST_PREVENT = "Prevention of recurrent cystitis"
_n10_blocks = {b["id"]: b for b in BLOCKS["N10"]}
SELF_START_TEXT = ("Nākotnē, parādoties tipiskiem cistīta simptomiem, drīkst uzsākt pašārstēšanos ar Pulv. Fosfomycin "
                   "3 g vienreizēji (pēc vienošanās ar ģimenes ārstu), pirms tam nododot urīnu uzsējumam.")
ADVICE_BLOCK_TEXTS = {SELF_START_TEXT}


def split_advice(blocks):
    """(medicines, advice lines): advice blocks are written as their own lines, not under 'Terapijā lietot'."""
    advice = [b for b in blocks if b in ADVICE_BLOCK_TEXTS]
    return [b for b in blocks if b not in advice], advice


BLOCKS["N30.0"] = [
    dict({b["id"]: b for b in BLOCKS["R31"]}["fosfomycin-3g"], group=G_CYST_FIRST,
         note="EAU (strong) first line in women: fosfomycin trometamol 3 g once. Safe in pregnancy. Not for men or "
              "with systemic signs. Long-term cure rates are a little lower than with nitrofurans."),
    {"id": "furamag", "group": G_CYST_FIRST, "title": "Furamag 100 mg three times daily, 7 days",
     "text": "Caps. Furamag 50 mg pa 2 kapsulām 3 reizes dienā pēc ēšanas 7 dienas.",
     "note": "Furazidine – the nitrofuran on the Latvian market (Olainfarm). Label: 50–100 mg three times daily after "
             "meals for 7–10 days. EAU names nitrofurantoin (5 days), not furazidine. Contraindicated in pregnancy, "
             "breastfeeding and severe renal failure; caution with G6PD deficiency. In men only if prostatitis is "
             "reliably excluded."},
    {"id": "nitrofurantoin", "group": G_CYST_FIRST, "title": "Nitrofurantoin 100 mg twice daily, 5 days",
     "text": "Caps. Nitrofurantoin 100 mg (ilgstošās darbības) 2 reizes dienā 5 dienas.",
     "note": "EAU (strong) first line. Check that it is available in Latvia (Latvian sources mention Furamag instead). "
             "Avoid if eGFR is below 45 (short courses possible at 30–44), at term pregnancy and in G6PD deficiency."},
    {"id": "cefuroxime-5", "group": G_CYST_ALT, "title": "Cefuroxime 250 mg twice daily, 5 days",
     "text": "Tab. Cefuroxime 250 mg pa 1 tabletei 2 reizes dienā 5 dienas.",
     "note": "Second choice. EAU lists cephalosporins (e.g. cefadroxil 500 mg twice daily for 3 days) as an "
             "alternative; Latvian outpatient material does not favour oral cephalosporins for empirical use. "
             "Useful in pregnancy or by culture."},
    {"id": "amoxiclav-5", "group": G_CYST_ALT, "title": "Amoksiklav 875/125 mg twice daily, 5 days (pregnancy / culture)",
     "text": "Tab. Amoksiklav 875/125 mg pa 1 tabletei 2 reizes dienā 5 dienas.",
     "note": "EAU (strong): no aminopenicillins for empirical treatment of cystitis. Use only by culture, or in "
             "pregnancy where penicillins are acceptable."},
    {"id": "cotrimoxazole-3", "group": G_CYST_ALT, "title": "Trimethoprim/sulfamethoxazole 160/800 mg twice daily, 3 days",
     "text": "Tab. Trimetoprims/Sulfametoksazols 160/800 mg pa 1 tabletei 2 reizes dienā 3 dienas.",
     "note": "EAU alternative only where local E. coli resistance is below 20% – Latvian outpatient sources report much "
             "higher resistance, so use it by culture. Not in the first trimester (trimethoprim) or the last "
             "trimester (sulfamethoxazole)."},
    {"id": "furamag-men", "group": G_CYST_MEN, "title": "Furamag 100 mg three times daily, 7 days (men)",
     "text": "Caps. Furamag 50 mg pa 2 kapsulām 3 reizes dienā pēc ēšanas 7 dienas.",
     "note": "EAU: nitrofurans can be used in younger men only if prostatic involvement is reliably excluded; there is "
             "no evidence on the duration – 7 days is usual. Take a urine culture first."},
    {"id": "cotrimoxazole-7", "group": G_CYST_MEN, "title": "Trimethoprim/sulfamethoxazole 160/800 mg twice daily, 7 days",
     "text": "Tab. Trimetoprims/Sulfametoksazols 160/800 mg pa 1 tabletei 2 reizes dienā 7 dienas.",
     "note": "EAU regimen for men (7 days) – only if the culture shows susceptibility (high resistance in Latvia)."},
    dict(_r33_blocks["cefuroxime-7"], group=G_CYST_MEN),
    dict(_r33_blocks["amoxiclav-7"], group=G_CYST_MEN,
         note="By culture only (EAU: no aminopenicillins empirically). Men: 7 days."),
    dict(_n10_blocks["cipro"], group=G_AVOID, kind="avoid",
         title="Ciprofloxacin – do not use for cystitis",
         note="EAU (strong): do not use fluoroquinolones for cystitis (EMA restrictions: tendon, nerve, aortic and "
              "QT risks). Only if every other option is unsuitable, by culture."),
    {"id": "ibuprofen", "group": G_PAIN, "title": "Ibuprofen 400 mg three times daily, 3 days",
     "text": "Tab. Ibuprofen 400 mg pa 1 tabletei 3 reizes dienā pēc ēšanas 3 dienas.",
     "note": "EAU (strong): non-antibiotic treatment is an alternative in non-geriatric patients after shared "
             "decision-making. In trials ibuprofen alone halved antibiotic use but gave slightly more pyelonephritis "
             "and longer symptoms. Not for the elderly, renal impairment, pregnancy (3rd trimester) or GI bleeding risk."},
    _n10_blocks["paracetamol"],
    dict(_n10_blocks["canephron"], text="Tab. Canephron pa 2 tabletēm 3 reizes dienā 7 dienas.",
         note="Herbal product popular in Latvia. In an RCT against fosfomycin it reduced antibiotic use but was not "
              "non-inferior; not an EAU recommendation."),
    {"id": "estriol", "group": G_CYST_PREVENT, "title": "Vaginal estriol (Ovestin) – postmenopausal women",
     "text": "Ovestin vaginālais krēms (estriols 1 mg/g) 0,5 g intravagināli katru vakaru 2 nedēļas, pēc tam 2 reizes "
             "nedēļā (ilgstoši).",
     "note": "EAU (strong): vaginal oestrogen to prevent recurrent cystitis in postmenopausal women (oral oestrogen "
             "does not work). Dosing as in the trials (0.5 mg estriol nightly for 2 weeks, then twice weekly)."},
    {"id": "cranberry", "group": G_CYST_PREVENT, "title": "Cranberry products daily",
     "text": "Dzērveņu produkti (sula vai kapsulas ar standartizētu proantocianidīnu saturu) ikdienā recidīvu "
             "profilaksei.",
     "note": "EAU (strong): advise cranberry products, but tell the patient the evidence is low quality and "
             "contradictory (Cochrane 2023: fewer symptomatic UTIs in women with recurrent UTI)."},
    {"id": "methenamine", "group": G_CYST_PREVENT, "title": "Methenamine hippurate 1 g twice daily",
     "text": "Tab. Methenamine hippurate 1 g pa 1 tabletei 2 reizes dienā 6–12 mēnešus recidīvu profilaksei.",
     "note": "EAU (strong): reduces recurrences in women without urinary tract abnormalities; the ALTAR trial found it "
             "non-inferior to daily antibiotic prophylaxis. Availability in Latvia not confirmed – check."},
    {"id": "d-mannose", "group": G_CYST_PREVENT, "title": "D-mannose 2 g daily",
     "text": "D-manoze 2 g 1 reizi dienā recidīvu profilaksei (uztura bagātinātājs).",
     "note": "EAU (weak): evidence is weak and contradictory – the large MERIT trial (2024) found no benefit."},
    {"id": "urovaxom", "group": G_CYST_PREVENT, "title": "Uro-Vaxom (OM-89) 6 mg daily, 90 days",
     "text": "Caps. Uro-Vaxom 6 mg 1 reizi dienā no rīta tukšā dūšā 90 dienas.",
     "note": "EAU 2025 downgraded immunoprophylaxis (weak, 'within a clinical trial'): the pooled analysis showed no "
             "significant short-term benefit of OM-89. Latvian outpatient material still recommends it."},
    {"id": "self-start", "group": G_CYST_PREVENT, "title": "Self-start fosfomycin for future episodes",
     "text": SELF_START_TEXT,
     "note": "EAU (strong): consider self-administered short-term therapy in reliable patients with recurrent "
             "cystitis. Written as its own line, not under 'Terapijā lietot'."},
]
BLOCK_WARNINGS["N30.0"] = BLOCK_WARNINGS["N10"]

# N45 epididymo-orchitis. EAU 2026 (strong): antibiotics active against C. trachomatis and Enterobacterales in young
# sexually active men, Enterobacterales only in older men without sexual risk; ceftriaxone 500 mg i/m or i/v once if
# gonorrhoea is likely, plus an anti-chlamydial drug. IUSTI 2024: doxycycline 14 days; levofloxacin / ofloxacin 10–14
# days for enteric organisms; moxifloxacin 14 days for M. genitalium.
G_EO_STI = "Antibiotics – STI likely"
G_EO_ENT = "Antibiotics – urinary / enteric source"
BLOCKS["N45"] = [
    {"id": "doxycycline", "group": G_EO_STI, "title": "Doxycycline 100 mg twice daily, 14 days",
     "text": "Caps. Doxycyclini 100 mg pa 1 kapsulai 2 reizes dienā 14 dienas.",
     "note": "EAU/IUSTI: anti-chlamydial therapy for sexually active men (with ceftriaxone if gonorrhoea is likely). "
             "Take upright with a full glass of water; photosensitivity – avoid sun."},
    {"id": "moxifloxacin", "group": G_EO_STI, "title": "Moxifloxacin 400 mg once daily, 14 days (M. genitalium)",
     "text": "Tab. Moxifloxacini 400 mg 1 reizi dienā 14 dienas.",
     "note": "IUSTI 2024: when Mycoplasma genitalium is found. Fluoroquinolone (EMA warnings: tendons, aorta, QT)."},
    {"id": "levofloxacin", "group": G_EO_ENT, "title": "Levofloxacin 500 mg once daily, 10 days",
     "text": "Tab. Levofloxacini 500 mg 1 reizi dienā 10 dienas.",
     "note": "EAU / IUSTI: first choice when an enteric (urinary) source is likely and gonorrhoea is unlikely; also "
             "covers chlamydia. EMA fluoroquinolone warnings (tendon rupture, neuropathy, aortic aneurysm, QT) – "
             "counsel the patient; avoid with previous fluoroquinolone reactions; adjust to the culture."},
    {"id": "cotrimoxazole-eo", "group": G_EO_ENT, "title": "Trimethoprim/sulfamethoxazole 160/800 mg twice daily, 14 days",
     "text": "Tab. Trimetoprims/Sulfametoksazols 160/800 mg pa 1 tabletei 2 reizes dienā 14 dienas.",
     "note": "EAU option for non-sexually-active men – only by culture (high resistance in Latvia)."},
    {"id": "amoxiclav-eo", "group": G_EO_ENT, "title": "Amoksiklav 875/125 mg twice daily, 10 days (by culture)",
     "text": "Tab. Amoksiklav 875/125 mg pa 1 tabletei 2 reizes dienā 10 dienas.",
     "note": "Not an EAU empirical choice; use by culture or after IV beta-lactam step-down."},
    {"id": "cefuroxime-eo", "group": G_EO_ENT, "title": "Cefuroxime 500 mg twice daily, 10 days (by culture)",
     "text": "Tab. Cefuroxime 500 mg pa 1 tabletei 2 reizes dienā 10 dienas.",
     "note": "Step-down after IV ceftriaxone / cefuroxime, by culture."},
    {"id": "clindamycin-eo", "group": G_EO_ENT, "title": "Clindamycin 300 mg four times daily, 10 days (S. aureus)",
     "text": "Tab. Clindamycin 300 mg 4 reizes dienā 10 dienas.",
     "note": "From your orchiectomy discharge (Staphylococcus aureus MS) – by culture only."},
    {"id": "ibuprofen-eo", "group": G_PAIN, "title": "Ibuprofen 400 mg three times daily after food",
     "text": "Tab. Ibuprofen 400 mg pa 1 tabletei 3 reizes dienā pēc ēšanas 5–7 dienas.",
     "note": "NSAIDs reduce pain and swelling (supportive care: rest, scrotal elevation). Avoid with renal impairment, "
             "GI bleeding risk or anticoagulants."},
    _n10_blocks["paracetamol"], _c61_blocks["ultracod"],
]
BLOCK_WARNINGS["N45"] = BLOCK_WARNINGS["N10"]

# N41.0 acute bacterial prostatitis. EAU 2026: treat as a systemic UTI (IV 3rd-generation cephalosporin, or
# aminopenicillin / 2nd-generation cephalosporin + aminoglycoside); ciprofloxacin only with resistance < 10% or
# beta-lactam problems, never empirically after fluoroquinolones in the last 6 months or in urology patients.
# NICE / HSE: oral ciprofloxacin or levofloxacin for 14 days, review and extend to 4 weeks if not fully settled.
G_AP_ORAL = "Antibiotics – oral (2–4 weeks in total)"
_r33 = {b["id"]: b for b in BLOCKS["R33"]}
_n45 = {b["id"]: b for b in BLOCKS["N45"]}
AP_FQ_NOTE = ("Fluoroquinolones reach the prostate best. EAU: not empirical after fluoroquinolones in the last "
              "6 months, after prostate biopsy (ESBL / quinolone resistance) or in urology in-patients; EMA warnings "
              "(tendons, nerves, aorta, QT). Review at 14 days; continue to 4 weeks if not fully settled.")
BLOCKS["N41.0"] = [
    {"id": "cipro-ap-14", "group": G_AP_ORAL, "title": "Ciprofloxacin 500 mg twice daily, 14 days",
     "text": "Tab. Ciprofloxacin 500 mg pa 1 tabletei 2 reizes dienā 14 dienas.", "note": AP_FQ_NOTE},
    {"id": "cipro-ap-28", "group": G_AP_ORAL, "title": "Ciprofloxacin 500 mg twice daily, 4 weeks",
     "text": "Tab. Ciprofloxacin 500 mg pa 1 tabletei 2 reizes dienā 4 nedēļas.",
     "note": AP_FQ_NOTE + " Your NMC example used 3 weeks."},
    {"id": "levofloxacin-ap", "group": G_AP_ORAL, "title": "Levofloxacin 500 mg once daily, 14 days",
     "text": "Tab. Levofloxacini 500 mg 1 reizi dienā 14 dienas.", "note": AP_FQ_NOTE},
    {"id": "cotrimoxazole-ap", "group": G_AP_ORAL, "title": "Trimethoprim/sulfamethoxazole 160/800 mg twice daily, 14 days",
     "text": "Tab. Trimetoprims/Sulfametoksazols 160/800 mg pa 1 tabletei 2 reizes dienā 14 dienas.",
     "note": "Good prostate penetration; only by culture (high resistance in Latvia)."},
    {"id": "amoxiclav-ap", "group": G_AP_ORAL, "title": "Amoksiklav 875/125 mg twice daily, 14 days (by culture)",
     "text": "Tab. Amoksiklav 875/125 mg pa 1 tabletei 2 reizes dienā 14 dienas.",
     "note": "Beta-lactams penetrate the prostate poorly – by culture or as step-down after IV beta-lactam."},
    {"id": "cefuroxime-ap", "group": G_AP_ORAL, "title": "Cefuroxime 500 mg twice daily, 14 days (by culture)",
     "text": "Tab. Cefuroxime 500 mg pa 1 tabletei 2 reizes dienā 14 dienas.",
     "note": "Step-down after IV ceftriaxone / cefuroxime, by culture; poorer prostate penetration."},
    dict(_n45["doxycycline"], note="Only with an STI (chlamydia) risk in young sexually active men – with the "
                                   "anti-enteric antibiotic, or by NAAT."),
    dict(_r33["tamsulosin-twoc"], title="Tamsulosin 0.4 mg once daily (voiding symptoms / retention)",
         text="Caps. Tamsulosin 0,4 mg 1 reizi dienā pēc brokastīm.",
         note="Alpha-blockers ease voiding symptoms in prostatitis and help a trial without catheter after retention."),
    _n45["ibuprofen-eo"], _n10_blocks["paracetamol"], _c61_blocks["macrogol"],
]
BLOCK_WARNINGS["N41.0"] = BLOCK_WARNINGS["N10"]
_ap_abx = ["cipro-ap-14", "cipro-ap-28", "levofloxacin-ap", "cotrimoxazole-ap", "amoxiclav-ap", "cefuroxime-ap"]
BLOCK_WARNINGS["N10"].extend({"ids": [a, b], "text": "Two antibiotics for the same organisms are chosen: keep one."}
                             for i, a in enumerate(_ap_abx) for b in _ap_abx[i + 1:])
_eo_abx = ["levofloxacin", "cotrimoxazole-eo", "amoxiclav-eo", "cefuroxime-eo", "clindamycin-eo", "moxifloxacin"]
BLOCK_WARNINGS["N10"].extend({"ids": [a, b], "text": "Two antibiotics for the same organisms are chosen: keep one."}
                             for i, a in enumerate(_eo_abx) for b in _eo_abx[i + 1:])
BLOCK_WARNINGS["N10"].append({"ids": ["levofloxacin", "moxifloxacin"], "text": "Two fluoroquinolones: keep one."})

# C67 additions (radical cystectomy). From the department's cystectomy discharge plus EAU MIBC.
G_STOMACH = "Stomach protection"
RC_BOWEL_TEXT = ("Sekot līdzi vēdera izejai; ja 2 dienas nav vēdera izejas, lietot Sol. Duphalac 66,5% 15 ml 2 reizes "
                 "dienā p/o vai Supp. Bisacodyl 5 mg p/r.")
ADVICE_BLOCK_TEXTS.add(RC_BOWEL_TEXT)
BLOCKS["C67"] = BLOCKS["C67"] + [
    dict(FRAXIPARINE_4W_BLOCK,
         note="EAU MIBC (strong): pharmacological VTE prophylaxis (LMWH) after radical cystectomy, from the first "
              "postoperative day, for at least 4 weeks. Reduce the dose if creatinine clearance is below 30 ml/min."),
    dict(_r33_blocks["amoxiclav-7"], id="amoxiclav-7-rc", group=G_ABX,
         text="Tab. Amoksiklav 875/125 mg pa 1 tabletei 2 reizes dienā 7 dienas.",
         note="From your cystectomy example. EAU: no prolonged antibiotic prophylaxis after RC – conduit urine is always "
              "colonised; use it for a documented infection (or per department protocol with stents)."),
    {"id": "omeprazole", "group": G_STOMACH, "title": "Omeprazole 20 mg once daily, 1 month",
     "text": "Caps. Omeprazol 20 mg 1 reizi dienā no rīta 30 min pirms ēšanas 1 mēnesi.",
     "note": "From your cystectomy example (stress ulcer / NSAID / LMWH period)."},
    {"id": "bowel-rc", "group": G_BOWEL, "title": "If no stool for 2 days: Duphalac or bisacodyl",
     "text": RC_BOWEL_TEXT,
     "note": "From your cystectomy example – written as its own line, not in the medicine list."},
    {"id": "bicarbonate", "group": G_PREVENT, "title": "Sodium bicarbonate (metabolic acidosis after bowel diversion)",
     "text": "Tab. Natrii hydrogencarbonas 500 mg 2–3 reizes dienā p/o (koriģēt pēc skābju-sārmu līdzsvara).",
     "note": "Only with proven hyperchloraemic metabolic acidosis (low bicarbonate), more common with neobladders. "
             "EAU follow-up: acidosis harms kidneys and bone; check a venous blood gas."},
    {b["id"]: b for b in BLOCKS["N30.0"]}["cranberry"],
]
BLOCKS["Z43.6"] = BLOCKS["Z43.6"] + [
    dict(_r33_blocks["cefuroxime-7"], group=G_ABX,
         note="Only with infection signs (fever, chills, cloudy urine and symptoms). A nephrostomy is usually colonised: "
              "EAU – do not treat asymptomatic bacteriuria. Adjust to the culture; obstructive pyelonephritis is "
              "treated 7–14 days."),
    dict(_r33_blocks["amoxiclav-7"], group=G_ABX,
         note="Only with infection signs; by culture (EAU: no aminopenicillins empirically for UTI)."),
]
_cyst_abx = ["fosfomycin-3g", "furamag", "nitrofurantoin", "cefuroxime-5", "amoxiclav-5", "cotrimoxazole-3",
             "furamag-men", "cotrimoxazole-7", "cefuroxime-7", "amoxiclav-7", "cipro"]
_have = {frozenset(w["ids"]) for w in BLOCK_WARNINGS["N10"]}
BLOCK_WARNINGS["N10"].extend({"ids": [a, b], "text": "Two antibiotics are chosen: keep one."}
                             for i, a in enumerate(_cyst_abx) for b in _cyst_abx[i + 1:]
                             if frozenset((a, b)) not in _have)
TAMSULOSIN_LONG_BLOCK = {
    "id": "tamsulosin-long", "group": G_ALPHA, "title": "Tamsulosin 0.4 mg once daily (long-term)",
    "text": "Caps. Tamsulosin 0,4 mg 1 reizi dienā p/o (ilgstoši).",
    "note": "From your urethrotomy example, for coexisting BPH symptoms. It does not treat the stricture itself.",
}
BLOCKS["N21.0"].insert(2, TAMSULOSIN_LONG_BLOCK)
BLOCKS["N35"] = BLOCKS["N35"] + [TAMSULOSIN_LONG_BLOCK,
                                 dict({b["id"]: b for b in BLOCKS["N10"]}["canephron"],
                                      text="Tab. Canephron pa 2 tabletēm 3 reizes dienā 1 mēnesi.",
                                      note="From your urethrotomy example. Herbal supportive product; there is no "
                                           "evidence that it prevents infection or stricture recurrence.")]
BLOCK_WARNINGS["N10"].append({"ids": ["tadalafil", "tadalafil-quarter"], "text": "Two tadalafil regimens are chosen: keep one."})
BLOCK_WARNINGS["N10"] += [
    {"ids": ["fraxiparine", "fraxiparine-04"], "text": "Two Fraxiparine doses are chosen (0,3 ml and 0,4 ml): keep one."},
    {"ids": ["fraxiparine-4w", "fraxiparine-04"], "text": "Two Fraxiparine regimens are chosen: keep one."},
    {"ids": ["urorec", "tamsulosin"], "text": "Urorec and tamsulosin are both alpha-blockers: keep one."},
]

# Text templates. A template is used when the diagnosis and department
# match AND every modifier listed under "when" has that value.
# Each template has: clinical_diagnosis (optional), epicrisis, recommendations.
TEMPLATES = [
    {
        # Acute bacterial prostatitis at NMC, oral antibiotics, sent home. Written by n41_nmc_parts.
        "id": "N41.0-NMC",
        "diagnosis": "N41.0",
        "department": "NMC",
        "when": {},
        "builder": "n41_nmc",
    },
    {
        # Acute bacterial prostatitis / prostatic abscess admitted to 4. Urology. Written by n41_ward_parts.
        "id": "N41.0-4",
        "diagnosis": "N41.0",
        "department": "4",
        "when": {},
        "builder": "n41_ward",
    },
    {
        # Epididymo-orchitis at NMC, oral antibiotics, sent home. Written by n45_nmc_parts.
        "id": "N45-NMC",
        "diagnosis": "N45",
        "department": "NMC",
        "when": {},
        "builder": "n45_nmc",
    },
    {
        # Epididymo-orchitis / abscess admitted to 4. Urology (IV antibiotics, surgery). Written by n45_ward_parts.
        "id": "N45-4",
        "diagnosis": "N45",
        "department": "4",
        "when": {},
        "builder": "n45_ward",
    },
    {
        # Radical cystectomy (any diversion) in 4. Urology. Written by c67_rc_parts.
        "id": "C67-4-RC",
        "diagnosis": "C67",
        "department": "4",
        "when": {"C67 operation": "Radical cystectomy"},
        "builder": "c67_rc",
    },
    {
        # Paraphimosis at NMC: manual reduction or dorsumcision under local anaesthesia. Written by n47_nmc_parts.
        "id": "N47-NMC",
        "diagnosis": "N47",
        "department": "NMC",
        "when": {},
        "builder": "n47_nmc",
    },
    {
        # Nephrostomy fell out, new one by the radiologist, sent home from NMC. Written by nf_nmc_parts.
        "id": "Z43.6-NMC",
        "diagnosis": "Z43.6",
        "department": "NMC",
        "when": {},
        "builder": "nf_nmc",
    },
    {
        # Nephrostomy fell out, new one by the radiologist, admitted to 4. Urology. Written by nf_ward_parts.
        "id": "Z43.6-4",
        "diagnosis": "Z43.6",
        "department": "4",
        "when": {},
        "builder": "nf_ward",
    },
    {
        # Acute cystitis at NMC, sent home. Written by n30_nmc_parts.
        "id": "N30.0-NMC",
        "diagnosis": "N30.0",
        "department": "NMC",
        "when": {},
        "builder": "n30_nmc",
    },
    {
        # Hematuria at NMC: consultation, sent home. Written by r31_nmc_parts.
        "id": "R31-NMC",
        "diagnosis": "R31",
        "department": "NMC",
        "when": {},
        "builder": "r31_nmc",
    },
    {
        # Hematuria admitted to 4. Urology (irrigation, transfusion, cystoscopy, embolisation). r31_ward_parts.
        "id": "R31-4",
        "diagnosis": "R31",
        "department": "4",
        "when": {},
        "builder": "r31_ward",
    },
    {
        # Acute urinary retention at NMC, catheter, sent home. Written by r33_parts.
        "id": "R33-NMC",
        "diagnosis": "R33",
        "department": "NMC",
        "when": {},
        "builder": "r33",
    },
    {
        # Bladder stone, cystolithotripsy in 4. Urology. Written by n21_parts.
        "id": "N21.0-4",
        "diagnosis": "N21.0",
        "department": "4",
        "when": {},
        "builder": "n21",
    },
    {
        # Scrotal mass excision in 4. Urology. Written by scrotal_parts.
        "id": "D29.4-4",
        "diagnosis": "D29.4",
        "department": "4",
        "when": {},
        "builder": "scrotal",
    },
    {
        # PUJ stricture, pyeloplasty in 4. Urology. Written by n13_parts.
        "id": "N13.0-4",
        "diagnosis": "N13.0",
        "department": "4",
        "when": {},
        "builder": "n13",
    },
    {
        # Hydrocele plasty in 4. Urology. Written by n43_parts.
        "id": "N43-4",
        "diagnosis": "N43",
        "department": "4",
        "when": {},
        "builder": "n43",
    },
    {
        # Pyelonephritis in 4. Urology (acute or planned). Written by n10_ward_parts.
        "id": "N10-4",
        "diagnosis": "N10",
        "department": "4",
        "when": {},
        "builder": "n10_ward",
    },
    {
        # Pyelonephritis seen at NMC and sent home. Written by n10_nmc_parts.
        "id": "N10-NMC",
        "diagnosis": "N10",
        "department": "NMC",
        "when": {"Stay": "Outpatient"},
        "builder": "n10_nmc",
    },
    {
        # Phimosis: surgery (spinal, general or local) or conservative. Written by n47_parts.
        "id": "N47-4",
        "diagnosis": "N47",
        "department": "4",
        "when": {},
        "builder": "n47",
    },
    {
        # Catheter dysfunction, urethral or cystostomy. Written by t83_parts.
        "id": "T83-NMC",
        "diagnosis": "T83",
        "department": "NMC",
        "when": {},
        "builder": "t83",
    },
    {
        # Emergency visit, discharged home. Written by n20_nmc_parts.
        "id": "N20-NMC-acute-outpatient",
        "diagnosis": "N20",
        "department": "NMC",
        "when": {"Admission": "Acute", "Stay": "Outpatient"},
        "builder": "n20_nmc",
    },
    {
        # Planned endoscopic treatment in 4. Urology. Written by n20_planned_parts.
        "id": "N20-4-planned",
        "diagnosis": "N20",
        "department": "4",
        "when": {"Admission": "Planned"},
        "builder": "n20_planned",
    },
    {
        # Bladder tumours, TURU in 4. Urology (acute or planned). Written by c67_parts.
        "id": "C67-4-TURU",
        "diagnosis": "C67",
        "department": "4",
        "when": {"C67 operation": "TURU / re-TURU"},
        "builder": "c67",
    },
    {
        "id": "C61-4-RP",
        "diagnosis": "C61",
        "department": "4",
        "when": {"Operation": "Radical prostatectomy"},
        "builder": "c61_rp",
    },
    {
        "id": "C61-4-TURP",
        "diagnosis": "C61",
        "department": "4",
        "when": {"Operation": "TURP"},
        "builder": "c61_turp",
    },
    {
        "id": "C61-NMC",
        "diagnosis": "C61",
        "department": "NMC",
        "when": {"Stay": "Outpatient"},
        "builder": "c61_nmc",
    },
    {
        "id": "C64-4-nephrectomy",
        "diagnosis": "C64",
        "department": "4",
        "when": {},
        "builder": "c64_surgery",
    },
    {
        "id": "C64-NMC",
        "diagnosis": "C64",
        "department": "NMC",
        "when": {"Stay": "Outpatient"},
        "builder": "c64_nmc",
    },
    {
        # Kidney stone with infection: acute admission and drainage. Written by n20_acute_parts.
        "id": "N20-4-acute",
        "diagnosis": "N20",
        "department": "4",
        "when": {"Admission": "Acute"},
        "builder": "n20_acute",
    },
    {
        # Urinary tract infection seen at NMC and sent home. Written by n39_nmc_parts.
        "id": "N39-NMC",
        "diagnosis": "N39.0",
        "department": "NMC",
        "when": {},
        "builder": "n39_nmc",
    },
    {
        # Urethral stricture, operation in 4. Urology. Written by n35_parts.
        "id": "N35-4",
        "diagnosis": "N35",
        "department": "4",
        "when": {},
        "builder": "n35",
    },
    {
        # Urethral stricture with acute retention: epicystostomy at NMC, sent home. Written by n35_acute_parts.
        "id": "N35-NMC",
        "diagnosis": "N35",
        "department": "NMC",
        "when": {},
        "builder": "n35_acute",
    },
    {
        # Planned nephrostomy change in 48. daily department. Written by neph_parts.
        "id": "Z43.6-48",
        "diagnosis": "Z43.6",
        "department": "48",
        "when": {},
        "builder": "nephrostomy",
    },
    {
        # BPH, TURP in 4. Urology. Written by n40_parts.
        "id": "N40-4-TURP",
        "diagnosis": "N40",
        "department": "4",
        "when": {},
        "builder": "n40",
    },
    {
        # The text of this one is written by a function (n39_parts) because it has many options.
        "id": "N39-4-acute",
        "diagnosis": "N39.0",
        "department": "4",
        "when": {"Admission": "Acute"},
        "builder": "n39",
    },
]

# ---------------------------------------------------------------
# 2. LOGIC
# ---------------------------------------------------------------


def fixed_values(code, department):
    """Modifier values that are set automatically for this diagnosis and department."""
    fixed = {}
    fixed.update(DEPARTMENT_RULES.get(department, {}))
    fixed.update(DIAGNOSIS_RULES.get(code, {}))
    return fixed


def find_template(code, department, values):
    """Return the first template that matches, or None."""
    for template in TEMPLATES:
        if template["diagnosis"] != code or template["department"] != department:
            continue
        if all(values.get(group) == wanted for group, wanted in template["when"].items()):
            return template
    return None


def block_label(block):
    """Text of a block in the old panel (which has no group headings)."""
    return f"{block['group']}: {block['title']}"


def add_blocks(recommendations, block_texts):
    """Put the chosen blocks into the recommendations: at the {blocks} line if the
    template has one, otherwise at the end."""
    lines = "".join(f"• {text}\n" for text in block_texts)
    if "{blocks}\n" in recommendations:
        return recommendations.replace("{blocks}\n", lines)
    if lines:
        return recommendations.rstrip("\n") + "\n" + lines.rstrip("\n")
    return recommendations


def build_forms(sex, values):
    """Collect every word that gets filled into the {placeholders} of a template."""
    forms = dict(SEX_FORMS[sex])
    if "Side" in values:
        forms["side"] = SIDE_TEXT[values["Side"]]
        forms["pyelo"] = PYELO_TEXT[values["Side"]]
    if "Phimosis degree" in values:
        forms["degree"] = DEGREE_TEXT[values["Phimosis degree"]]
    if "Balanoposthitis" in values:
        forms["fimoze"] = BALANOPOSTHITIS_TEXT[bool(values["Balanoposthitis"])]
    if "Surgery type" in values:
        forms["op"] = SURGERY_TEXT[values["Surgery type"]]["nom"]
        forms["op_gen"] = SURGERY_TEXT[values["Surgery type"]]["gen"]
    if "Operation date" in values:
        forms["op_date"] = values["Operation date"].strftime("%d.%m.%Y")
    if "Catheter size (Fr)" in values:
        forms["fr"] = values["Catheter size (Fr)"]
    if "Catheter reason" in values:
        forms["reason"] = REASON_TEXT[values["Catheter reason"]]
    if "Side" in values and ("Stone size (mm)" in values or "Right stone size (mm)" in values):
        words = STONE_WORDS[values["Side"]]
        sizes = {}
        if values["Side"] == "Both":
            sizes["size_right"] = f"Ø {values['Right stone size (mm)']:g} mm"
            sizes["size_left"] = f"Ø {values['Left stone size (mm)']:g} mm"
        else:
            sizes["size"] = f"Ø {values['Stone size (mm)']:g} mm"
        forms["stone_pain"] = words["pain"]
        forms["stone_epi"] = words["epi"]
        forms["stone_dx"] = words["dx"].format(**sizes)
    return forms


# ---------------------------------------------------------------
# 2a. URINARY TRACT INFECTION (N39.0): helpers, checks and text builders
# ---------------------------------------------------------------


def condition_met(actual, wanted):
    """Does a modifier value satisfy a show_if condition? "*" means: anything filled in."""
    if wanted == "*":
        return actual not in (None, "", [], False)
    options = wanted if isinstance(wanted, list) else [wanted]
    if isinstance(actual, list):  # a multi-select: any chosen option counts
        return any(a in options for a in actual)
    return actual in options


def blank(value):
    return value is None or (isinstance(value, str) and not value.strip())


def lv_list(items):
    """['a', 'b', 'c'] -> 'a, b un c' (Latvian list)."""
    items = [i for i in items if i]
    if len(items) <= 1:
        return "".join(items)
    return ", ".join(items[:-1]) + " un " + items[-1]


def lv_list_or(items):
    """['a', 'b', 'c'] -> 'a, b vai c'."""
    items = [i for i in items if i]
    if len(items) <= 1:
        return "".join(items)
    return ", ".join(items[:-1]) + " vai " + items[-1]


def fmt_date(value):
    return value.strftime("%d.%m.%Y")


def clean_rows(df):
    """Turn the table from st.data_editor into a list of plain dicts (empty rows dropped)."""
    rows = []
    for record in df.to_dict("records"):
        row = {}
        for key, value in record.items():
            if value is None or (not isinstance(value, (str, bool)) and pd.isna(value)):
                value = None
            elif hasattr(value, "date") and callable(value.date):
                value = value.date()  # Timestamp / datetime -> date
            row[key] = value
        if any(not blank(v) and v is not False for v in row.values()):
            rows.append(row)
    return rows


# ----- lab values -----

def lab_series(rows, test):
    """Sorted (date, value) pairs of one test."""
    points = [(r["Date"], r["Value"]) for r in rows
              if r.get("Test") == test and r.get("Value") is not None and r.get("Date") is not None]
    return sorted(points)


def lab_direction(points):
    """Compare the last value with the first one. Needs at least two values."""
    if len(points) < 2:
        return None
    first, last = points[0][1], points[-1][1]
    if last < first:
        return "Decreasing"
    if last > first:
        return "Increasing"
    return "Stable"


def lab_summary(rows, tests):
    parts = []
    for test in tests:
        label, unit = LAB_TESTS[test]
        points = lab_series(rows, test)
        if not points:
            continue
        if len(points) == 1:
            parts.append(f"{label} {points[0][1]:g} {unit}")
        else:
            parts.append(f"{label} {points[0][1]:g} → {points[-1][1]:g} {unit}")
    return "; ".join(parts)


# ----- antibiotics -----

def abx_rows(v):
    return sorted(v.get("Antibiotics") or [], key=lambda r: r["Start date"])


def abx_groups(rows):
    """Group the rows that start on the same day: [(date, [rows]), ...] oldest first."""
    groups = []
    for row in rows:
        if groups and groups[-1][0] == row["Start date"]:
            groups[-1][1].append(row)
        else:
            groups.append((row["Start date"], [row]))
    return groups


def abx_phrase(row, with_details):
    """Wording after 'ar ...' / 'uz ...'. The first course has no dose, a change has dose and frequency."""
    meta = MEDICATION_CATALOG[row["Medication"]]
    if with_details:
        text = meta.get("acc_dose", meta["acc"] + " {dose}").format(dose=(row.get("Dose") or "").strip())
    else:
        text = meta["acc"]
    parts = [text, ROUTE_WORD[row["Route"]]]
    if with_details and row.get("Frequency"):
        parts.append(row["Frequency"])
    return " ".join(" ".join(parts).split()).replace(" )", ")")


def abx_discharge_lines(v):
    lines = []
    for row in abx_rows(v):
        if not row.get("Continue at discharge"):
            continue
        meta = MEDICATION_CATALOG[row["Medication"]]
        text = (f"{meta['form']} {meta['nom']} {(row.get('Dose') or '').strip()} "
                f"{ROUTE_WORD[row['Route']]} {row['Frequency']} līdz {fmt_date(row['End date'])}.")
        lines.append("    o " + " ".join(text.split()))
    return lines


# ----- safety checks -----

def n39_checks(v):
    """Returns (blockers, cautions). Blockers stop the routine discharge text."""
    blockers, cautions = [], []

    status = v.get("Clinical status")
    if status in ("Deteriorated", "Incomplete assessment"):
        blockers.append(f"Clinical status at discharge is '{status}'.")
    if v.get("Temperature trend") in ("Persistent fever", "Recurrent fever"):
        blockers.append(f"Temperature: {v['Temperature trend'].lower()}.")
    if v.get("Inflammatory markers trend") == "Increasing":
        blockers.append("Inflammatory markers are increasing.")
    if v.get("Renal function trend") == "Deteriorated":
        blockers.append("Renal function has deteriorated.")
    if v.get("Hemodynamic status") == "Unstable":
        blockers.append("Hemodynamic status is unstable.")
    if v.get("Oral intake") == "Unable to tolerate":
        blockers.append("The patient cannot tolerate oral intake.")
    if v.get("Discharge readiness") == "Requires reassessment":
        blockers.append("Discharge readiness: requires reassessment.")
    if v.get("Complications status") == "Ongoing":
        blockers.append("A complication (for example obstruction) is still ongoing.")
    if v.get("Urine drainage") == "None":
        blockers.append("No urine drainage through the catheter: assess for obstruction.")
    if v.get("Voiding after removal") == "Urinary retention":
        blockers.append("Urinary retention after catheter removal.")
    if v.get("Antibiotic renal adjustment") == "Requires review":
        blockers.append("Antibiotic renal dosing needs review.")
    if v.get("Antibiotic review") == "Requires prescriber review":
        blockers.append("The antibiotic plan requires prescriber review.")

    admission, discharge = v.get("Admission date"), v.get("Discharge date")
    if admission and discharge and discharge < admission:
        blockers.append("The discharge date is before the admission date.")

    for number, row in enumerate(v.get("Antibiotics") or [], start=1):
        name = row.get("Medication") or f"row {number}"
        for field in ("Medication", "Route", "Start date"):
            if blank(row.get(field)):
                blockers.append(f"Antibiotics, row {number}: {field} is missing.")
        if row.get("Continue at discharge"):
            for field in ("Dose", "Frequency", "End date"):
                if blank(row.get(field)):
                    blockers.append(f"Antibiotics, row {number} ({name}) continues at discharge: {field} is missing.")
            if discharge and row.get("End date") and row["End date"] < discharge:
                cautions.append(f"{name}: the end date is before the discharge date.")
        if admission and row.get("Start date") and row["Start date"] < admission:
            cautions.append(f"{name}: the start date is before the admission date.")

    labs = v.get("Labs") or []
    directions = [d for d in (lab_direction(lab_series(labs, t)) for t in INFLAMMATORY_TESTS) if d]
    if v.get("Inflammatory markers trend") == "Decreasing":
        if not directions:
            cautions.append("Inflammatory markers are called 'decreasing', but no CRO or leukocyte "
                            "values (two or more per test) were entered to support this.")
        elif any(d != "Decreasing" for d in directions):
            cautions.append("Inflammatory markers are called 'decreasing', but the entered CRO / "
                            "leukocyte values do not all decrease.")

    if status == "Improved" and (v.get("Pain trend") in ("Unchanged", "Increased")
                                 or v.get("Urinary symptoms trend") == "Persistent"):
        cautions.append("The status is 'improved', but pain or urinary symptoms have not improved. "
                        "The text will not say that the complaints subsided.")
    if v.get("Antibiotic allergy status") == "Documented allergy":
        cautions.append("An allergy is documented: confirm that none of the antibiotics listed is affected.")
    elif v.get("Antibiotic allergy status") == "Unknown":
        cautions.append("Antibiotic allergy status is unknown.")
    if v.get("Treatment tolerance") == "Adverse reaction":
        cautions.append("An adverse reaction to treatment is recorded.")
    if (v.get("Urine specimen") == "Not taken") != (v.get("Urine culture") == "Not performed"):
        cautions.append("Urine specimen and urine culture do not match (one says none was taken or done).")
    if v.get("Urine culture") == "Possibly contaminated sample":
        cautions.append("The urine culture sample is possibly contaminated: consider a repeat sample.")
    if v.get("Voiding after removal") in ("Not tested", "Increased residual urine"):
        cautions.append("Voiding after catheter removal: " + v["Voiding after removal"].lower() + ".")
    if v.get("Urine drainage") == "Not documented":
        cautions.append("Urine drainage through the catheter is not documented.")
    if v.get("Obstruction drained") == "No":
        blockers.append("Urostasis with infection that has not been drained: drainage is needed before discharge.")
    if v.get("Sepsis signs") == "Yes":
        blockers.append("Signs of sepsis: admission is needed.")
    if v.get("Able to take oral medication") == "No":
        blockers.append("The patient cannot take oral medication: admission for IV treatment is needed.")
    if v.get("Urine appearance") == "With clots":
        cautions.append("Clots in the urine: check for clot retention and the source of bleeding.")
    if v.get("_department") == "NMC" and not any(r.get("Continue at discharge") for r in v.get("Antibiotics") or []):
        cautions.append("No antibiotic is marked 'Continue at discharge' for a patient sent home.")
    return blockers, cautions


# ----- text builders -----

def n39_suprapubic(v):
    return v.get("Has catheter") and v.get("Catheter route") == "Suprapubic (epicistostoma)"


def n39_catheter_noun(v):
    """Catheter wording by route: gen (katetra), acc (katetru), nom (katetrs), short (without Foley)."""
    if n39_suprapubic(v):
        return {"gen": "epicistostomas katetra", "acc": "epicistostomas katetru",
                "nom": "epicistostomas katetrs", "short": "epicistostomas katetrs"}
    return {"gen": "urīnpūšļa Foley katetra", "acc": "urīnpūšļa Foley katetru",
            "nom": "urīnpūšļa Foley katetrs", "short": "urīnpūšļa katetrs"}


def n39_diagnosis(v):
    context = CONTEXT_TEXT[v["Infection context"]]
    if n39_suprapubic(v):
        context = context.replace("urīnpūšļa katetra", "epicistostomas katetra")
    first = COMPLEXITY_TEXT[v["Infection complexity"]] + context
    if v.get("Diagnosis certainty") == "Possible":
        first = "Iespējama " + first[0].lower() + first[1:]
    lines = [first + "."]
    site = v["Infection site"]
    if site == "Pyelonephritis":
        lines.append(PYELO_TEXT[v["Pyelonephritis side"]])
    elif site in SITE_TEXT:
        lines.append(SITE_TEXT[site])
    complications = v.get("Complications") or []
    if complications:
        lines.append("Komplikācijas: " + ", ".join(COMPLICATION_TEXT.get(c, c) for c in complications) + ".")
    if not blank(v.get("Comorbidities")):
        lines.append("Blakusslimības: " + v["Comorbidities"].strip())
    return "\n".join(lines)


def n39_history(v):
    """Devices in the history: stents and catheter."""
    has_stents, has_catheter = v.get("Has stents"), v.get("Has catheter")
    catheter = n39_catheter_noun(v)["short"]
    if has_stents:
        kind = "urīnvada DJ" if v.get("Stent type") == "DJ" else "urīnvada"
        if v["Stent count"] == "2":
            stents = f"implantēti 2 {kind} stenti"
        else:
            stents = "implantēts " + {"Right": "labās", "Left": "kreisās"}[v["Stent side"]] + f" puses {kind} stents"
        when = ""
        if v.get("Stent insertion date"):
            when = fmt_date(v["Stent insertion date"]) + " "
        elif not blank(v.get("Stent age")):
            when = f"pirms aptuveni {v['Stent age'].strip()} "
        text = f"Anamnēzē {when}{stents}"
        if v.get("Stent indication") in STENT_INDICATION_TEXT:
            text += f" ({STENT_INDICATION_TEXT[v['Stent indication']]})"
        if has_catheter:
            text += f" un ievietots {catheter}"
        return text + "."
    if has_catheter:
        return f"Anamnēzē ievietots {catheter}."
    return ""


def n39_background(v):
    """Duration, onset and fever after the complaints."""
    parts = []
    duration, onset = v.get("Symptom duration"), v.get("Symptom onset")
    if duration in DURATION_TEXT:
        text = f"Sūdzības {DURATION_TEXT[duration]}"
        if onset in ONSET_TEXT:
            text += f", sākušās {ONSET_TEXT[onset]}"
        parts.append(text + ".")
    elif onset in ONSET_TEXT:
        parts.append(f"Sūdzības sākušās {ONSET_TEXT[onset]}.")
    if v.get("Fever status") in FEVER_STATUS_TEXT:
        parts.append(FEVER_STATUS_TEXT[v["Fever status"]])
    return " ".join(parts)


def n39_other_history(v, devices_mentioned):
    history = v.get("Relevant history") or []
    if not history:
        return ""
    lead = "Anamnēzē arī: " if devices_mentioned else "Anamnēzē: "
    return lead + ", ".join(HISTORY_TEXT.get(h, h) for h in history) + "."


def n39_catheter_sentence(v):
    fr = v["Catheter Fr"]
    noun = n39_catheter_noun(v)
    if v["Catheter action"] == "Changed":
        text = f"NMC etapā veikta {noun['gen']} maiņa, ievietojot CH/Fr {fr} katetru."
    elif v["Catheter action"] == "Inserted":
        text = f"NMC etapā ievietots {noun['nom']} CH/Fr {fr}."
    else:
        text = f"Saglabāts iepriekš ievietotais {noun['nom']} CH/Fr {fr}."
    indication = v.get("Catheter indication")
    if indication and indication != "Not specified":
        text += f" Indikācija: {CATHETER_INDICATION_TEXT[indication]}."
    return text


def n39_findings(v):
    """Urine appearance and imaging."""
    parts = []
    if v.get("Urine appearance") in URINE_APPEARANCE_TEXT:
        parts.append(URINE_APPEARANCE_TEXT[v["Urine appearance"]])
    if v.get("Imaging") in IMAGING_TEXT:
        name = IMAGING_TEXT[v["Imaging"]]
        if v.get("Hydronephrosis") == "None":
            parts.append(f"{name}: bez urostāzes.")
        elif v.get("Hydronephrosis") in HYDRO_TEXT:
            text = f"{name}: {HYDRO_TEXT[v['Hydronephrosis']]} urostāze"
            if v.get("Obstruction drained") == "Yes":
                text += ", veikta nieres drenāža"
            parts.append(text + ".")
    return " ".join(parts)


def n39_stent_words(v):
    """('abi urīnvada stenti' style words) for the stent paragraphs."""
    if v["Stent count"] == "2":
        return {"in_place": "abi urīnvada stenti atrodas vietā", "removed": "urīnvada stenti izņemti",
                "care": "Stentu"}
    side = {"Right": "labās", "Left": "kreisās"}[v["Stent side"]]
    return {"in_place": f"{side} puses urīnvada stents atrodas vietā",
            "removed": f"{side} puses urīnvada stents izņemts", "care": "Stenta"}


def n39_opening(f, v):
    symptoms = lv_list([SYMPTOM_TEXT.get(s, s) for s in v["Symptoms"]])
    opening = (f"{f['Pacients']} akūtā kārtā {f['iestajies']} P. Stradiņa Klīniskās universitātes "
               f"slimnīcas Neatliekamās medicīnas centrā (NMC) ar sūdzībām par {symptoms}.")
    history = n39_history(v)
    for part in (n39_background(v), history, n39_other_history(v, bool(history))):
        if part:
            opening += " " + part
    return opening


def n39_first_sentences(v):
    """What happened at NMC: catheter, specimen, urine and imaging."""
    sentences = []
    if v.get("Has catheter"):
        sentences.append(n39_catheter_sentence(v))
    sentences.append(SPECIMEN_TEXT[v["Urine specimen"]])
    findings = n39_findings(v)
    if findings:
        sentences.append(findings)
    return sentences


def n39_epicrisis(sex, v):
    f = SEX_FORMS[sex]
    paragraphs = [n39_opening(f, v)]

    # NMC phase: catheter, specimen, findings, first antibiotics, admission to the ward
    sentences = n39_first_sentences(v)
    groups = abx_groups(abx_rows(v))
    sentences.append("Uzsākta antibakteriālā terapija ar "
                     + lv_list([abx_phrase(r, False) for r in groups[0][1]]) + ".")
    sentences.append(f"{f['Pacients']} {f['stacionets']} 4. Uroloģijas nodaļā turpmākai ārstēšanai.")
    paragraphs.append(" ".join(sentences))

    # Later changes of antibiotics
    for day, rows in groups[1:]:
        paragraphs.append(f"{fmt_date(day)}. antibakteriālā terapija mainīta uz "
                          + lv_list([abx_phrase(r, True) for r in rows]) + ".")

    # Course
    if v["Clinical status"] == "Improved":
        status_text = "vispārējais stāvoklis uzlabojies"
        complaints_better = (v.get("Pain trend") in ("Resolved", "Reduced")
                             or v.get("Urinary symptoms trend") in ("Resolved", "Improved")) \
            and v.get("Pain trend") not in ("Unchanged", "Increased") \
            and v.get("Urinary symptoms trend") != "Persistent"
        if complaints_better:
            status_text += ", sūdzības mazinājušās"
    else:
        status_text = "stabils"
    inflammatory = INFLAMMATORY_TEXT[v["Inflammatory markers trend"]]
    summary = lab_summary(v.get("Labs") or [], INFLAMMATORY_TESTS)
    if summary:
        inflammatory += f" ({summary})"
    culture = CULTURE_TEXT[v["Urine culture"]]
    if v["Urine culture"] == "Positive" and not blank(v.get("Culture organism")):
        culture += f" ({v['Culture organism'].strip()})"
    course = ("Stacionēšanas laikā veikta ārstēšana un novērošana. "
              f"{f['pacienta'].capitalize()} stāvoklis izrakstīšanas brīdī: {status_text}. "
              f"Iekaisuma rādītāju dinamika: {inflammatory}. "
              f"Urīna uzsējuma rezultāts: {culture}.")
    if v["Urine culture"] == "Result not yet known":
        course += " Nepieciešams nodrošināt rezultāta pārbaudi un turpmāko rīcību atbilstoši atradei."
    if v.get("Blood culture") in BLOOD_CULTURE_TEXT:
        course += f" Asins uzsējuma rezultāts: {BLOOD_CULTURE_TEXT[v['Blood culture']]}."
    paragraphs.append(course)

    # Catheter and stents at discharge
    paragraphs += n39_device_paragraphs(f, v, "Stacionēšanas laikā")

    # Closing
    if v["Clinical status"] == "Improved":
        lead = f"Uzlabojoties {f['pacienta']} vispārējam stāvoklim"
    else:
        lead = f"Saglabājoties stabilam {f['pacienta']} stāvoklim"
    paragraphs.append(f"{lead} un izvērtējot turpmākās ārstēšanas iespējas, {f['pacients']} "
                      f"{f['izrakstits']} turpmākai aprūpei ģimenes ārsta un urologa uzraudzībā. "
                      "Sniegtas rekomendācijas.")
    return "\n".join(paragraphs)


def n39_device_paragraphs(f, v, when):
    paragraphs = []
    noun = n39_catheter_noun(v)
    if v.get("Has catheter"):
        if v["Catheter at discharge"] == "In place":
            paragraphs.append(f"{f['Pacients']} {f['izrakstits']} ar {noun['acc']} CH/Fr {v['Catheter Fr']}. "
                              f"Urīna attece izrakstīšanas brīdī: {DRAINAGE_TEXT[v['Urine drainage']]}.")
        else:
            paragraphs.append(f"Pirms izrakstīšanas {noun['short']} izņemts. "
                              f"Patstāvīgas urinācijas rezultāts: {VOIDING_TEXT[v['Voiding after removal']]}.")
    if v.get("Has stents"):
        words = n39_stent_words(v)
        if v["Stents at discharge"] == "In place":
            paragraphs.append(f"Izrakstīšanas brīdī {words['in_place']}. {words['care']} turpmākā aprūpe un "
                              "izņemšanas vai maiņas termiņi saskaņojami ar ārstējošo urologu.")
        else:
            paragraphs.append(f"{when} {words['removed']}. Turpmākā uroloģiskā novērošana "
                              "atbilstoši ārstējošā urologa norādījumiem.")
    return paragraphs


def n39_nmc_epicrisis(sex, v):
    f = SEX_FORMS[sex]
    paragraphs = [n39_opening(f, v)]
    sentences = n39_first_sentences(v)
    rows = abx_rows(v)
    if rows:
        sentences.append("Nozīmēta antibakteriālā terapija ar "
                         + lv_list([abx_phrase(r, True) for r in abx_groups(rows)[0][1]]) + ".")
    paragraphs.append(" ".join(sentences))
    paragraphs += n39_device_paragraphs(f, v, "NMC etapā")
    paragraphs.append("Dotajā brīdī nav indikāciju neatliekamai stacionēšanai uroloģijas nodaļā. Vispārēji "
                      "apmierinošā stāvoklī un ar rekomendācijām izraksta tālākai ambulatorai ārstēšanai "
                      "ģimenes ārsta uzraudzībā.")
    return "\n".join(paragraphs)


def n39_recommendations(v, nmc=False):
    lines = []
    counter = [0]

    def numbered(text):
        counter[0] += 1
        return f"{counter[0]}. {text}"

    if v["Urology follow-up"] == "On a specific date":
        when = f"kontrole pie urologa {fmt_date(v['Urology follow-up date'])}"
    else:
        when = "kontrole pie urologa (datums vēl nav noteikts)"
    lines.append(numbered(f"Ģimenes ārsta uzraudzība un {when}."))

    antibiotics = abx_discharge_lines(v)
    if antibiotics:
        lines.append(numbered("Antibakteriālā terapija:"))
        lines.extend(antibiotics)
        lines.append("    o Terapiju precizēt atbilstoši urīna uzsējuma rezultātiem, nieru funkcijai "
                     "un ārstējošā ārsta norādījumiem.")
    lines.append("{blocks}")  # the blocks you drag in end up here
    chronic = (v.get("Chronic medication changes") or "").strip().rstrip(".")
    lines.append(numbered("Turpināt iepriekš nozīmēto pastāvīgo terapiju" + (f"; izmaiņas: {chronic}." if chronic else ".")))

    lines.append(numbered(f"Šķidruma uzņemšana {v['Fluid intake'].strip()}, ja nav medicīnisku kontrindikāciju."))
    lines.append(numbered("Saudzējošs režīms, izvairīties no pārmērīgas fiziskas slodzes."))

    lines += tests_lines(v, "numbered", numbered)  # the tests ticked in the Follow-up tests box
    if v["Urine culture"] == "Result not yet known":
        who = CULTURE_CHECKER_TEXT[v["Culture checked by"]]
        lines.append(numbered(f"Urīna uzsējuma rezultātu pārbauda {who}, nepieciešamības gadījumā "
                              "sazinoties ar pacientu terapijas korekcijai."))

    if v.get("Has catheter"):
        if v["Catheter at discharge"] == "In place":
            heading = "Epicistostomas katetra aprūpe:" if n39_suprapubic(v) else "Urīnpūšļa katetra aprūpe:"
            lines += [
                heading,
                "• Nodrošināt regulāru katetra un urīna savākšanas maisa aprūpi.",
                "• Uzturēt noslēgtu drenāžas sistēmu un maisu novietot zem urīnpūšļa līmeņa.",
                "• Sekot urīna attecei un katetra caurlaidībai.",
            ]
            if n39_suprapubic(v):
                lines.append("• Stomas vietu ikdienā apkopt un pārsiet.")
                lines += ["• Epicistostomas katetra maiņa ik pēc 4–6 nedēļām ģimenes ārsta praksē.",
                          "• Urīna savākšanas maisa maiņa ik pēc 3–7 dienām."]
            else:
                lines += ["• " + line for line in CATHETER_CARE_LINES]
        else:
            lines += [
                "Urinācijas novērošana:",
                "• Sekot patstāvīgai urinācijai un urīnpūšļa iztukšošanās kvalitātei.",
                "• Urīna aiztures vai nespējas urinēt gadījumā nekavējoties vērsties pēc medicīniskās palīdzības.",
            ]
    if v.get("Has stents") and v["Stents at discharge"] == "In place":
        if v["Stent plan"] == "On a specific date":
            plan_text = "plānota " + fmt_date(v["Stent plan date"])
        else:
            plan_text = STENT_PLAN_TEXT[v["Stent plan"]]
        lines += [
            "Urīnvada stentu aprūpe:" if v["Stent count"] == "2" else "Urīnvada stenta aprūpe:",
            f"• Kontrole pie urologa un stentu izņemšana vai maiņa {plan_text}."
            if v["Stent count"] == "2" else
            f"• Kontrole pie urologa un stenta izņemšana vai maiņa {plan_text}.",
            "• Neatlikt noteikto kontroles vizīti.",
        ]

    if nmc:
        lines.append(numbered("Ja 48–72 stundu laikā stāvoklis neuzlabojas vai saglabājas paaugstināta ķermeņa "
                              "temperatūra, nekavējoties vērsties NMC."))
    signs = [WARNING_TEXT[s] for s in WARNING_TEXT if s in (v.get("Warning signs") or [])]
    if signs:
        text = lv_list_or(signs)
        lines.append(numbered(f"{text[0].upper()}{text[1:]} gadījumā nekavējoties vērsties NMC."))
    return "\n".join(lines)


def n39_parts(sex, v, blocks=()):
    return {
        "clinical_diagnosis": n39_diagnosis(v),
        "epicrisis": n39_epicrisis(sex, v),
        "recommendations": add_blocks(n39_recommendations(v), list(blocks)),
    }


BUILDERS = {"n39": n39_parts}
CHECKS = {"N39.0": n39_checks}


# ---------------------------------------------------------------
# 2c. KIDNEY STONES (N20): checks and text builders
# ---------------------------------------------------------------


def make_counter():
    """numbered("text") gives "1. text", then "2. ...", and so on."""
    count = [0]

    def numbered(text):
        count[0] += 1
        return f"{count[0]}. {text}"
    return numbered


def n20_stone_words(v):
    """(diagnosis line, words after 'vizualizē', pain area) for the stone(s)."""
    side, place = v["Side"], STONE_PLACE[v["Stone location"]]
    if side == "Both":
        dx = (f"Abu pušu konkrementi {place['both']} (labajā pusē Ø {v['Right stone size (mm)']:g} mm, "
              f"kreisajā pusē Ø {v['Left stone size (mm)']:g} mm)")
        epi = ("iesprūdušus konkrementus " if place["impacted"] else "konkrementus ") + place["Both"]
    else:
        side_word = "Labās" if side == "Right" else "Kreisās"
        dx = f"{side_word} puses konkrements Ø {v['Stone size (mm)']:g} mm {place['dx']}"
        epi = ("iesprūdušu konkrementu " if place["impacted"] else "konkrementu ") + place[side]
    pain = f"{place['area']}, {STONE_WORDS[side]['pain']}"
    return dx, epi, pain


def n20_urostasis(v):
    grade = v["Urostasis grade"]
    return "bez urostāzes" if grade == "No urostasis" else f"ar {grade} pakāpes urostāzi"


def n20_checks(v):
    blockers, cautions = [], []
    if v.get("Admission") == "Planned":
        if v.get("DJ stent placed") and v.get("Stent removal") == "At the outpatient urology review":
            cautions.append("A DJ stent stays in: if possible give a concrete removal date, so the stent is not forgotten.")
        if "With complications" in (v.get("Intraoperative course"), v.get("Postoperative course")):
            cautions.append("Complications are recorded: check that the recommendations cover them.")
    elif v.get("_department") == "NMC":
        names = {"CRO elevated": "raised CRO", "Leukocytosis": "leukocytosis",
                 "Urine infection signs": "infection signs in the urine"}
        signs = [names[n] for n in names if v.get(n)]
        if signs:
            blockers.append("Signs of infection (" + ", ".join(signs) + ") together with a stone: an obstructed, "
                            "infected kidney needs admission and drainage. Write that discharge from "
                            "4. Urology department as an acute admission.")
        if v.get("Pain controlled") == "No":
            blockers.append("Pain is not controlled.")
    else:  # acute admission in 4. Urology with drainage
        if v.get("Clinical status") in ("Deteriorated", "Incomplete assessment"):
            blockers.append(f"Clinical status at discharge is '{v['Clinical status']}'.")
        if v.get("Temperature trend") in ("Persistent fever", "Recurrent fever"):
            blockers.append(f"Temperature: {v['Temperature trend'].lower()}.")
        if v.get("Inflammatory markers trend") == "Increasing":
            blockers.append("Inflammatory markers are increasing.")
        abx_blockers, abx_cautions = abx_row_checks(v)
        blockers += abx_blockers
        cautions += abx_cautions
        if v.get("Urostasis grade") == "No urostasis":
            cautions.append("A drainage was done, but no urostasis is recorded.")
        if v.get("Drainage") == "DJ stent" and not v.get("Definitive treatment date"):
            cautions.append("The patient goes home with a DJ stent: give the date of the definitive treatment, "
                            "so the stent is not forgotten.")
        if v.get("Urine culture") == "Possibly contaminated sample":
            cautions.append("The urine culture sample is possibly contaminated: consider a repeat sample.")
    return blockers, cautions


def n20_nmc_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    dx_line, stone_epi, pain = n20_stone_words(v)
    epicrisis = (
        f"{f['Pacients']} {f['iestajies']} akūtā kārtā PSKUS NMC ar sūdzībām par sāpēm {pain}. "
        f"NMC etapā vispusīgi {f['izmekleta']}; veic CT vēdera dobumam, kur vizualizē {stone_epi} "
        f"{n20_urostasis(v)}. Dinamikā vispārējais stāvoklis uzlabojas, sāpes mazinās. Dotajā brīdī nav "
        "indikāciju neatliekamai stacionēšanai un ķirurģiskai terapijai. Vispārēji apmierinošā stāvoklī un ar "
        "rekomendācijām izraksta tālākai ambulatorai terapijai ģimenes ārsta uzraudzībā."
    )
    numbered = make_counter()
    lines = [
        numbered("Atrasties ģimenes ārsta uzraudzībā"),
        numbered("Lietot daudz šķidruma, vismaz 2-2.5 l/DN"),
        numbered("Aktīvs dzīvesveids!"),
    ]
    if blocks:  # the medications you drag in
        lines.append(numbered("Terapijā lietot:"))
        lines.extend("    " + b for b in blocks)
    lines.append(numbered("Ambulatori urologa konsultācija pēc mēneša."))
    lines += tests_lines(v, "numbered", numbered)
    lines.append(numbered("Stāvokļa pasliktināšanās gadījumā vērsties pēc palīdzības NMC."))
    return {
        "clinical_diagnosis": "Nierakmeņu slimība\n" + dx_line,
        "epicrisis": epicrisis,
        "recommendations": "\n".join(lines),
    }


def n20_planned_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    side = v["Side"]

    # Clinical diagnosis
    diagnosis = "Nierakmeņu slimība. " + STONE_LOCATION_TEXT[v["Stone location"]][side] + " konkrementi."
    if v.get("Renal colic"):
        diagnosis += " " + COLIC_TEXT[side]
    if v.get("Urinary tract infection") == "Yes":
        diagnosis += " Urīnceļu infekcija."

    # The operation, as a sentence and as a short name
    procedure = PROCEDURE_TEXT[v["Procedure"]]
    with_parts, name_parts = [], []
    for done in v.get("Lithotripsy") or []:
        with_parts.append(LITHO_TEXT[done][0])
        name_parts.append(LITHO_TEXT[done][1])
    if v.get("DJ stent placed"):
        stent_words = STENTS_TEXT if side == "Both" else STENT_TEXT
        with_parts.append(stent_words[0])
        name_parts.append(stent_words[1])
    summary = f"Veikta {PROCEDURE_SIDE[side]} {procedure}"
    if with_parts:
        summary += " ar " + lv_list(with_parts)
    name = f"{PROCEDURE_SIDE[side]} {procedure}"
    name = name[0].upper() + name[1:]
    if name_parts:
        name += ", " + ", ".join(name_parts)
    day = fmt_date(v["Procedure date"])

    if v["Intraoperative course"] == "Without complications":
        intraop = "bez komplikācijām"
    else:
        intraop = "ar komplikācijām: " + v["Intraoperative complications"].strip().rstrip(".")
    if v["Postoperative course"] in POSTOP_TEXT:
        postop = POSTOP_TEXT[v["Postoperative course"]]
    else:
        postop = "ar komplikācijām: " + v["Postoperative complications"].strip().rstrip(".")

    epicrisis = "\n".join([
        f"{f['Pacients']} plānveida kārtā {f['iestajies']} P. Stradiņa Klīniskās universitātes slimnīcas "
        "4. Uroloģijas nodaļā endoskopiskai ārstēšanai iepriekš diagnosticētas nierakmeņu slimības dēļ. "
        f"Veikti nepieciešamie pirmsoperācijas izmeklējumi, izvērtēts {f['pacienta']} vispārējais "
        "veselības stāvoklis un anestēzijas risks.",
        f"{summary} ({day}). Operācija noritēja {intraop}. Pēcoperācijas periods {postop}. "
        f"{f['pacienta'].capitalize()} vispārējais stāvoklis izrakstīšanas brīdī: "
        f"{DISCHARGE_STATUS_TEXT[v['Discharge status']]}.",
        f"{f['Pacients']} {f['izrakstits']} turpmākai ambulatorai ārstēšanai ar rekomendācijām.",
        f"Izmeklējumi: {v['Investigations'].strip().rstrip('.')}.",
        f"Operācijas: {name} ({day}).",
    ])

    # Recommendations (lines that do not apply are left out)
    recs = [
        "• Turpināt ģimenes ārsta un ārstējošā urologa uzraudzību.",
        f"• Šķidruma uzņemšana {v['Fluid intake'].strip()}, ja nav medicīnisku kontrindikāciju. "
        "Saudzējošs režīms, izvairīties no pārmērīgas fiziskas slodzes.",
    ]
    if v.get("DJ stent placed"):
        if v["Stent removal"] == "On a specific date":
            when = "plānota " + fmt_date(v["Stent removal date"])
        else:
            when = "plānota ambulatorās urologa vizītes laikā"
        stent = "DJ stentu" if side == "Both" else "DJ stenta"
        recs.append(f"• {stent} izņemšana vai maiņa {when}; neatlikt noteikto vizīti.")
    if v.get("Discharged with catheter"):
        recs += ["• " + line for line in CATHETER_CARE_LINES]
        if v.get("Catheter removal date"):
            recs.append(f"• Katetra evakuācija plānota {fmt_date(v['Catheter removal date'])}.")
    recs.extend("• " + b for b in blocks)  # the medications you drag in
    if v["Urology follow-up"] == "On a specific date":
        recs.append(f"• Kontrole pie ārstējošā urologa {fmt_date(v['Urology follow-up date'])}.")
    else:
        recs.append("• Kontrole pie ārstējošā urologa (datums vēl nav noteikts).")
    recs += tests_lines(v, "bullet", after_op=True)
    recs.append("• Stāvokļa pasliktināšanās, paaugstinātas ķermeņa temperatūras, drebuļu, stipru sāpju "
                "vai urīna atteces traucējumu gadījumā nekavējoties vērsties NMC.")
    return {"clinical_diagnosis": diagnosis, "epicrisis": epicrisis, "recommendations": "\n".join(recs)}


BUILDERS.update({"n20_nmc": n20_nmc_parts, "n20_planned": n20_planned_parts})
CHECKS["N20"] = n20_checks


# ---------------------------------------------------------------
# 2d. BLADDER TUMOURS (C67, TURU): checks and text builder
# ---------------------------------------------------------------


C67_SMOKING = ("Atmest smēķēšanu – smēķēšana palielina urīnpūšļa vēža recidīva un progresijas risku "
               "(nepieciešamības gadījumā – ģimenes ārsta palīdzība smēķēšanas atmešanā).")
NMIBC_ADJUVANT_TEXT = {
    "No further instillations": "Papildu intravezikāla terapija nav nepieciešama.",
    "Intravesical chemotherapy course": "Plānota intravezikāla ķīmijterapija (indukcijas kurss 1 reizi nedēļā 6–8 "
                                        "nedēļas, uzturošā terapija līdz 1 gadam) – pēc ārstējošā urologa norādījumiem.",
    "BCG induction + maintenance": "Plānota intravezikāla BCG terapija: indukcijas kurss 1 reizi nedēļā 6 nedēļas, "
                                   "uzturošā terapija pa 3 instilācijām ar nedēļas intervālu 3., 6. un 12. mēnesī "
                                   "(augsta riska gadījumā – līdz 3 gadiem) – pēc ārstējošā urologa norādījumiem.",
    "Discuss radical cystectomy": "Onkoloģiskajā konsīlijā apspriest radikālas cistektomijas iespēju (ļoti augsta "
                                  "riska urīnpūšļa vēzis).",
}


def nmibc_risk(v):
    """EAU NMIBC risk group (2021 table, WHO 2004/2016 grade). None if the histology is not known or not NMIBC."""
    if v.get("Histology") != "Known":
        return None
    stage = v.get("Tumour stage")
    if stage not in ("pTa", "pT1", "pTis (CIS)"):
        return None
    hg = v.get("Tumour grade") == "High grade (HG)"
    cis = stage == "pTis (CIS)" or bool(v.get("Concomitant CIS"))
    rf = sum(bool(x) for x in (v.get("Age over 70 (C67)"), v.get("Tumour number") not in (None, "Single"),
                               v.get("Largest tumour over 3 cm")))
    if v.get("Variant histology / LVI"):
        return "very high"
    if stage == "pT1" and hg and ((cis and rf >= 1) or (not cis and rf == 3)):
        return "very high"
    if stage == "pTa" and hg and cis and rf == 3:
        return "very high"
    if cis:
        return "high"
    if stage == "pT1" and hg:
        return "high"
    if (stage == "pTa" and hg and rf >= 2) or (stage == "pT1" and not hg and rf >= 2) or \
            (stage == "pTa" and not hg and rf == 3):
        return "high"
    primary = not v.get("Recurrent tumour")
    if primary and not hg and ((stage == "pTa" and rf <= 1) or (stage == "pT1" and rf == 0)):
        return "low"
    return "intermediate"


def c67_checks(v):
    if v.get("C67 operation") == "Radical cystectomy":
        return c67_rc_checks(v)
    blockers, cautions = [], []
    risk = nmibc_risk(v)
    plan = v.get("Adjuvant plan (NMIBC)")
    if risk:
        expected = {"low": ("No further instillations",),
                    "intermediate": ("Intravesical chemotherapy course", "BCG induction + maintenance"),
                    "high": ("BCG induction + maintenance", "Discuss radical cystectomy"),
                    "very high": ("Discuss radical cystectomy",)}[risk]
        advice = {"low": "single immediate instillation only, no further adjuvant treatment (strong)",
                  "intermediate": "intravesical chemotherapy (up to 1 year) or 1 year of full-dose BCG",
                  "high": "full-dose BCG for 1–3 years, and discuss immediate radical cystectomy (strong)",
                  "very high": "discuss immediate radical cystectomy; BCG for 1–3 years if unfit or declining "
                               "(strong)"}[risk]
        if plan not in expected:
            cautions.append(f"EAU NMIBC risk group (calculated): {risk}. EAU: {advice}.")
    if v.get("Intravesical treatment") == "Not performed" and v.get("TURU type") == "TURU":
        cautions.append("No single immediate instillation: EAU (strong) offers it to low-risk and to intermediate-risk "
                        "patients with few recurrences – omit only with perforation, extensive resection or bleeding.")
    if v.get("Detrusor muscle in specimen") == "Absent" and not (v.get("Tumour stage") == "pTa" and
                                                                v.get("Tumour grade") == "Low grade (LG)"):
        cautions.append("No detrusor muscle in the specimen: EAU (strong) – second TURB within 2–6 weeks (except Ta LG "
                        "and primary CIS).")
    if v.get("Variant histology / LVI"):
        cautions.append("Variant histology or LVI: EAU treats these as very high risk – discuss radical cystectomy.")
    if v.get("Voiding after catheter removal") == "Urinary retention":
        blockers.append("Urinary retention after catheter removal.")
    if v.get("Histology") == "Pending" and v.get("Discharge paper") == "Final":
        cautions.append("Histology is still pending, but the discharge paper is marked as final.")
    if v.get("Histology") == "Known" and v.get("Discharge paper") == "Temporary":
        cautions.append("Histology is already known, but the discharge paper is marked as temporary.")
    if v.get("Histology") == "Pending" and v.get("Oncology council") == "Not applicable":
        cautions.append("Histology is still pending and no oncology council plan is set: "
                        "make sure someone follows up the result.")
    if v.get("Postoperative course") == "With complications":
        cautions.append("Complications are recorded: check that the recommendations cover them.")
    if v.get("Tumour stage") == "Other / not malignant" and blank(v.get("Histology details")):
        blockers.append("Histology is 'other / not malignant': describe it in 'Histology details'.")
    advanced = (v.get("Tumour stage") == "pT2"
                or v.get("N stage") in ("cN1", "cN2", "cN3")
                or v.get("M stage") in ("cM1a", "cM1b"))
    if advanced and v.get("Oncology council") == "Not applicable":
        cautions.append("Muscle-invasive, node-positive or metastatic disease, but no oncology council plan is set.")
    return blockers, cautions


def c67_stage_text(v, with_nm=True):
    """'pT1 cN0 cM0 HG + CIS' style text, or '' if histology is unknown or not a carcinoma.
    N and M come from imaging, not histology, so the histology sentence leaves them out."""
    if v.get("Histology") != "Known" or v.get("Tumour stage") in (None, "Other / not malignant"):
        return ""
    text = v["Tumour stage"]
    for part in ("N stage", "M stage"):  # nodes and metastases, only when assessed
        if with_nm and v.get(part) not in (None, "Not assessed"):
            text += " " + v[part]
    grade = GRADE_TEXT.get(v.get("Tumour grade"), "")
    if grade:
        text += " " + grade
    if v.get("Concomitant CIS"):
        text += " + CIS"
    return text


def c67_diagnosis(v):
    stage = c67_stage_text(v)
    if stage:
        first = f"Urīnpūšļa CA {stage}."
    elif v["Tumour number"] == "Single":
        first = "Urīnpūšļa jaunveidojums."
    else:
        first = "Urīnpūšļa jaunveidojumi."
    return f"{first}\nSt. pēc {TURU_SHORT[v['TURU type']]} ({fmt_date(v['Procedure date'])})."


def c67_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    single = v["Tumour number"] == "Single"
    verification = lv_list([VERIFICATION_TEXT[x] for x in v["Tumour verification"]])
    admission = "akūtā kārtā" if v["Admission"] == "Acute" else "plānveida kārtā"
    if single:
        tumours = f"{verification} verificētu urīnpūšļa jaunveidojumu"
    else:
        tumours = f"{verification} verificētiem {TUMOUR_NUMBER_TEXT[v['Tumour number']]} urīnpūšļa jaunveidojumiem"
    paragraphs = [f"{f['Pacients']} {f['stacionets']} P. Stradiņa KUS 4. uroloģijas nodaļā {admission} "
                  f"sakarā ar {tumours}."]

    # Operation
    location = lv_list([TUMOUR_LOCATION_TEXT[x] for x in v["Tumour location"]])
    if single:
        found = f"konstatēts urīnpūšļa {location} jaunveidojums, veikta tā rezekcija un hemostāze."
    else:
        found = f"konstatēti urīnpūšļa {location} jaunveidojumi, veikta to rezekcija un hemostāze."
    operation = (f"{fmt_date(v['Procedure date'])} veikta {TURU_TEXT[v['TURU type']]}. "
                 f"Operācijas laikā {found}")
    if not blank(v.get("Operation details")):
        details = v["Operation details"].strip()
        operation += " " + (details if details.endswith(".") else details + ".")
    paragraphs.append(operation)

    # Course after the operation
    course = []
    if v["Intravesical treatment"] in INSTILLATION_TEXT:
        course.append(f"Agrīnā pēcoperācijas periodā veikta {INSTILLATION_TEXT[v['Intravesical treatment']]}.")
    elif v["Intravesical treatment"] == "Other (type below)":
        course.append(f"Agrīnā pēcoperācijas periodā veikta {v['Intravesical treatment (other)'].strip().rstrip('.')}.")
    if v["Postoperative course"] in C67_POSTOP_TEXT:
        course.append(f"Pēcoperācijas periods noritēja {C67_POSTOP_TEXT[v['Postoperative course']]}.")
    else:
        course.append("Pēcoperācijas periods noritēja ar komplikācijām: "
                      f"{v['Postoperative complications'].strip().rstrip('.')}.")
    removal = v["Catheter removal"]
    planned = v.get("Planned catheter removal (optional)")
    if removal == "Discharged with catheter":
        sentence = f"{f['Pacients']} {f['izrakstits']} ar urīna ilgkatetru"
        if planned:
            sentence += f", kura evakuācija plānota {fmt_date(planned)}"
        course.append(sentence + ".")
    else:
        when = CATHETER_DAY_TEXT.get(removal) or fmt_date(v["Catheter removed on"])
        course.append(f"Urīna ilgkatetrs evakuēts {when}, pēc kā "
                      f"{C67_VOIDING_TEXT[v['Voiding after catheter removal']]}.")
    paragraphs.append(" ".join(course))

    # Histology and closing
    if v["Histology"] == "Pending":
        histology = "Histoloģiskās izmeklēšanas rezultāts izrakstīšanas brīdī vēl nav zināms"
    else:
        details = (v.get("Histology details") or "").strip().rstrip(".")
        histology = "Histoloģiskās izmeklēšanas rezultāts: " + ", ".join(p for p in (c67_stage_text(v, with_nm=False), details) if p)
    paragraphs.append(f"{histology}. Uzlabojoties vispārējam stāvoklim, {f['pacients']} {f['izrakstits']} "
                      "turpmākai aprūpei ģimenes ārsta uzraudzībā. Sniegtas rekomendācijas.")

    # Recommendations
    recs = []
    if v["Discharge paper"] == "Temporary":
        recs.append(TEMPORARY_PAPER_TEXT)
    recs += [
        "Atrasties ģimenes ārsta uzraudzībā.",
        "Diēta: vispusīga. Lietot šķidrumu aptuveni 2,0 L diennaktī, ja nav medicīnisku kontrindikāciju.",
        f"Režīms: saudzējošs {WEEKS_TEXT[v['Activity restriction']]} no operācijas dienas. Atturēties no "
        "pārmērīgas fiziskas slodzes, pirts un baseina apmeklējuma.",
    ]
    if removal == "Discharged with catheter":
        recs += CATHETER_CARE_LINES
        if planned:
            recs.append(f"Katetra evakuācija plānota {fmt_date(planned)}.")
    else:
        recs.append("Regulāri, ik pēc 1–2 stundām, iztukšot urīnpūsli.")
    if NMIBC_ADJUVANT_TEXT.get(v.get("Adjuvant plan (NMIBC)")):
        recs.append(NMIBC_ADJUVANT_TEXT[v["Adjuvant plan (NMIBC)"]])
    recs.append(f"Vizīte pie urologa ambulatori: {WEEKS_TEXT[v['Urology visit']]} pēc operācijas.")
    recs += tests_lines(v, "plain", after_op=True)  # e.g. cystoscopy 3 months after the operation
    if v.get("Smoker (C67)"):
        recs.append(C67_SMOKING)
    council = v["Oncology council"]
    if council == "Will be discussed after histology":
        recs.append(f"{f['pacienta'].capitalize()} gadījums tiks izskatīts onkoloģiskajā konsīlijā pēc "
                    "histoloģiskās atbildes saņemšanas.")
    elif council == "Already discussed":
        recs.append(f"{f['pacienta'].capitalize()} gadījums izskatīts onkoloģiskajā konsīlijā "
                    f"{fmt_date(v['Oncology council date'])}.")
    recs.append("Ja parādās izteikta hematūrija ar asins recekļiem, nav iespējams urinēt, rodas drudzis vai "
                "būtiski pasliktinās vispārējais stāvoklis, nekavējoties vērsties pēc neatliekamās "
                "medicīniskās palīdzības.")
    return {"clinical_diagnosis": c67_diagnosis(v), "epicrisis": "\n".join(paragraphs),
            "recommendations": "\n".join(recs)}


BUILDERS["c67"] = c67_parts
CHECKS["C67"] = c67_checks


# ---------------------------------------------------------------
# 2c-b. RADICAL CYSTECTOMY (C67) in 4. Urology
# ---------------------------------------------------------------
# EAU MIBC: RC for T2–T4a N0M0 (and very high-risk / BCG-unresponsive NMIBC) within 3 months of diagnosis unless
# neoadjuvant therapy is given; neoadjuvant cisplatin-based chemotherapy (or perioperative gemcitabine/cisplatin +
# durvalumab) for eligible patients; standard LND (obturator, internal and external iliac) – extended LND does not
# improve survival; sexual-organ-preserving techniques only in selected patients; no orthotopic neobladder with
# tumour in the urethra / at the urethral margin; ERAS, no bowel preparation; LMWH from day 1 for at least 4 weeks;
# adjuvant cisplatin chemotherapy for pT3/4 and/or pN+ without neoadjuvant treatment, adjuvant nivolumab for high-risk
# disease (not eligible for / declined cisplatin; EMA: PD-L1 >= 1%); adjuvant RT can be offered for pT3b–4, N+ or R1
# (weak). Follow-up: CT every 6 months to year 3, then yearly; upper tract imaging in multifocal disease, CIS or
# positive ureteral margins; vitamin B12 yearly after bowel diversion; watch for acidosis, renal function, stoma
# problems, stones and ureteroenteric strictures (up to 20%). Campbell-Walsh-Wein: conduit urine is always
# colonised – treat only symptomatic infection; stents are usually removed after 10–14 days (department: ~4 weeks).

RC_PREVIOUS = {  # (diagnosis line, epicrisis words)
    "TURU": ("St. pēc TURU.", "TURU"), "Re-TURU": ("St. pēc re-TURU.", "re-TURU"),
    "BCG instillations": ("St. pēc BCG intravezikālas terapijas.", "BCG intravezikāla terapija"),
    "Intravesical chemotherapy": ("St. pēc intravezikālas ķīmijterapijas.", "intravezikāla ķīmijterapija"),
    "Neoadjuvant cisplatin chemotherapy": ("St. pēc neoadjuvantas cisplatīnu saturošas ķīmijterapijas.",
                                           "neoadjuvanta cisplatīnu saturoša ķīmijterapija"),
    "Perioperative durvalumab (NIAGARA)": ("St. pēc neoadjuvantas ķīmijimūnterapijas (gemcitabīns, cisplatīns, "
                                           "durvalumabs).", "neoadjuvanta ķīmijimūnterapija ar durvalumabu"),
    "Pelvic radiotherapy": ("St. pēc staru terapijas mazajam iegurnim.", "staru terapija mazajam iegurnim"),
}
RC_INDICATION = {"Muscle-invasive bladder cancer": "muskuli invadējošu urīnpūšļa vēzi",
                 "Very high-risk / BCG-unresponsive NMIBC": "ļoti augsta riska / BCG nereaģējošu neinvazīvu "
                                                            "urīnpūšļa vēzi",
                 "Palliative (bleeding, pain, fistula)": "urīnpūšļa vēzi"}
RC_OPERATION = {  # (as done, after "St. pēc ...")
    "Radical cystoprostatectomy": ("radikāla cistoprostatektomija", "radikālas cistoprostatektomijas"),
    "Nerve-sparing cystoprostatectomy": ("nervus saudzējoša radikāla cistoprostatektomija",
                                         "nervus saudzējošas radikālas cistoprostatektomijas"),
    "With hysterectomy, adnexa and anterior vaginal wall": (
        "radikāla cistektomija ar histerektomiju, piedēkļu un maksts priekšējās sienas rezekciju",
        "radikālas cistektomijas ar histerektomiju, piedēkļu un maksts priekšējās sienas rezekciju"),
    "With anterior vaginal wall (no uterus)": ("radikāla cistektomija ar maksts priekšējās sienas rezekciju",
                                               "radikālas cistektomijas ar maksts priekšējās sienas rezekciju"),
    "Organ-sparing (uterus and vagina preserved)": ("orgānus saudzējoša radikāla cistektomija",
                                                    "orgānus saudzējošas radikālas cistektomijas"),
    "Palliative cystectomy": ("paliatīva cistektomija", "paliatīvas cistektomijas"),
}
RC_APPROACH = {"Open": ("", ""), "Robot-assisted": ("robotasistēta ", "robotasistētas "),
               "Laparoscopic": ("laparoskopiska ", "laparoskopiskas ")}
RC_DIVERSION = {"Ileal conduit": ("ileum conduit izveide", "ileum conduit izveides"),
                "Orthotopic neobladder": ("ortotopiskas ileālas neocistis izveide",
                                          "ortotopiskas ileālas neocistis izveides"),
                "Ureterocutaneostomy": ("ureterokutanostomija", "ureterokutanostomijas"),
                "No diversion – nephrostomies": (None, None)}
RC_LND = {"Standard bilateral (obturator, internal and external iliac)": (
              "abpusēja iliakāla un obturatoru limfadenektomija", "abpusējas iliakālas un obturatoru limfadenektomijas"),
          "Extended bilateral": ("abpusēja paplašināta iegurņa limfadenektomija",
                                 "abpusējas paplašinātas iegurņa limfadenektomijas"),
          "Not performed": (None, None)}
RC_ANESTHESIA = {"Epidural + general": "epidurālā un intubācijas anestēzijā", "General": "intubācijas anestēzijā"}
RC_DIVERSION_OK = {"Ileal conduit": "Ileum conduit funkcija apmierinoša.",
                   "Orthotopic neobladder": "Neocistis funkcija apmierinoša.",
                   "Ureterocutaneostomy": "Ureterokutanostomas funkcionē apmierinoši.",
                   "No diversion – nephrostomies": "Nefrostomas funkcionē apmierinoši."}
RC_VISIT = {"4–6 weeks": "4–6 nedēļas", "3 months": "3 mēnešus", "6 months": "6 mēnešus"}
RC_WARNING = ("Ja paaugstinās ķermeņa temperatūra, parādās drebuļi, sāpes vēderā vai sānos, slikta dūša, vemšana, vēdera "
              "uzpūšanās un gāzu neiziešana, samazinās vai izbeidzas urīna izdale pa stomu (nefrostomām), brūce kļūst "
              "apsārtusi vai no tās izdalās strutas vai urīns, parādās kājas tūska vai elpas trūkums, nekavējoties "
              "vērsties NMC.")
RC_STOMA_CARE = {
    "Ileal conduit": [
        "Kopt ileum conduit urostomu: urostomas plāksni un maisu mainīt pēc nepieciešamības (parasti ik pēc 2–3 dienām), "
        "ādu ap stomu tīrīt ar siltu ūdeni un nosusināt, plāksnes atveri pielāgot stomas izmēram; naktī pievienot "
        "nakts urīnsavācējmaisu. Urostomā normāli redzamas gļotas.",
        "Stomas aprūpes līdzekļus izrakstīs ģimenes ārsts; nepieciešamības gadījumā – stomas aprūpes māsas "
        "konsultācija.",
    ],
    "Ureterocutaneostomy": [
        "Kopt ureterokutanostomu: plāksni un maisu mainīt pēc nepieciešamības (parasti ik pēc 2–3 dienām), ādu ap "
        "stomu tīrīt ar siltu ūdeni un nosusināt; naktī pievienot nakts urīnsavācējmaisu.",
        "Stomas aprūpes līdzekļus izrakstīs ģimenes ārsts; nepieciešamības gadījumā – stomas aprūpes māsas "
        "konsultācija.",
    ],
    "Orthotopic neobladder": [
        "Kopt urīnpūšļa katetru; katetra laikā neocistis skalot ar fizioloģisko šķīdumu (gļotu evakuācijai) pēc "
        "nodaļā apmācītās metodes.",
        "Pēc katetra evakuācijas urinēt pēc pulksteņa ik pēc 2–3 stundām (arī naktī), urinējot atslābināt iegurņa "
        "pamatni un viegli sasprindzināt vēdera presi; pakāpeniski pagarināt intervālus; veikt iegurņa pamatnes "
        "muskulatūras (Kēgela) vingrinājumus.",
        "Lietot pietiekami šķidruma un sāls – neocistis var izraisīt sāls un šķidruma zudumu.",
    ],
}


def _rc_operation_names(v, sex):
    kind = v.get("Cystectomy type (woman)") if sex == "Woman" else v.get("Cystectomy type (man)")
    nom, gen = RC_OPERATION[kind]
    pre_nom, pre_gen = RC_APPROACH[v["Approach (RC)"]]
    nom, gen = pre_nom + nom, pre_gen + gen
    parts_nom, parts_gen = [nom], [gen]
    if v.get("Urethrectomy (RC)"):
        parts_nom.append("uretrektomija"); parts_gen.append("uretrektomijas")
    d_nom, d_gen = RC_DIVERSION[v["Urinary diversion"]]
    if d_nom:
        parts_nom.append(d_nom); parts_gen.append(d_gen)
    l_nom, l_gen = RC_LND[v["Lymph node dissection (RC)"]]
    if l_nom:
        parts_nom.append(l_nom); parts_gen.append(l_gen)
    return lv_list(parts_nom), lv_list(parts_gen), kind


def _rc_stage(v):
    if v.get("Histology (RC)") == "Known":
        stage = v["pT (RC)"] + v["pN (RC)"].replace("p", "", 1)
        nodes = (v.get("Nodes positive / removed (RC)") or "").strip()
        if nodes:
            stage += f" ({nodes})"
        return stage + f", {v['Margins (RC)']}"
    return _rc_clinical(v)


def _rc_clinical(v):
    return v["Clinical T (RC)"] + v["Clinical N (RC)"].replace("c", "", 1) + v["Clinical M (RC)"].replace("c", "", 1)


def _rc_units(n):
    n = int(n)
    return f"{n} ER masas vienību" if n % 10 == 1 and n % 100 != 11 else f"{n} ER masas vienības"


def c67_rc_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    nom, gen, kind = _rc_operation_names(v, sex)
    day = fmt_date(v["Procedure date"])
    previous = v.get("Previous treatment (RC)") or []
    diversion = v["Urinary diversion"]
    palliative = kind == "Palliative cystectomy" or v["Indication (RC)"].startswith("Palliative")

    # Clinical diagnosis
    lines = [f"Ca vesicae urinariae {_rc_stage(v)}" + (", HG" if v.get("Histology (RC)") != "Known" else "") + "."]
    lines += [RC_PREVIOUS[x][0] for x in previous if x in RC_PREVIOUS]
    if v.get("Persistent hematuria (RC)"):
        lines.append("Persistējoša makrohematūrija.")
    if v.get("Anaemia (RC)"):
        lines.append("Anēmija.")
    neph = [x for x in previous if x.startswith("Nephrostomy")]
    if len(neph) == 1:
        lines.append("Labās puses nefrostoma." if neph[0] == "Nephrostomy right" else "Kreisās puses nefrostoma.")
    elif neph or diversion == "No diversion – nephrostomies":
        lines.append("Abpusējas nefrostomas.")
    lines.append(f"St. pēc {gen} ({day}).")
    if not blank(v.get("Other diagnoses (optional)")):
        lines.append(v["Other diagnoses (optional)"].strip().rstrip(".") + ".")

    # Epicrisis
    reason = "histoloģiski verificētu " + RC_INDICATION[v["Indication (RC)"]]
    if v.get("Persistent hematuria (RC)"):
        reason += " ar persistējošu makrohematūriju"
    p1 = [f"{f['Pacients']} {f['stacionets']} P. Stradiņa KUS 4. uroloģijas nodaļā "
          f"{'akūtā' if v.get('Admission') == 'Acute' else 'plānveida'} kārtā "
          f"{'paliatīvai ' if palliative else ''}operatīvai ārstēšanai sakarā ar {reason} ({_rc_clinical(v)})."]
    prev_words = [RC_PREVIOUS[x][1] for x in previous if x in RC_PREVIOUS]
    if prev_words:
        p1.append(f"Anamnēzē – {', '.join(prev_words)}.")
    p1.append(f"{f['Pacients']} tiek {f['gatavots']} operācijai.")
    pre = _num(v, "Preoperative transfusion (RC, units)")
    if pre:
        p1.append(("Sakarā ar anēmiju pirmsoperācijas periodā " if v.get("Anaemia (RC)") else
                   "Pirmsoperācijas periodā ") + f"{'saņēmusi' if sex == 'Woman' else 'saņēmis'} {_rc_units(pre)}.")
    intraop, postop = c61_courses(v)
    p2 = [f"{day} {RC_ANESTHESIA[v['Anesthesia (RC)']]} tiek veikta {nom}. Operācija noritēja {intraop}."]
    p3 = [f"Pēcoperācijas periods {postop}."]
    post = _num(v, "Postoperative transfusion (RC, units)")
    if post:
        p3.append(f"Agrīnā pēcoperācijas periodā {'saņēmusi' if sex == 'Woman' else 'saņēmis'} {_rc_units(post)}.")
    p3.append("Atjaunojusies zarnu peristaltika un vēdera izeja." if v["Bowel recovery (RC)"].startswith("Peristalsis")
              else "Pēcoperācijas periodā zarnu parēze, kas izzuda konservatīvas terapijas rezultātā.")
    p3.append(RC_DIVERSION_OK[diversion])
    if v["Drain (RC)"] == "Removed in the ward":
        d = _num(v, "Drain removed on day (RC)")
        p3.append(f"Drena evakuēta {int(d)}. pēcoperācijas dienā." if d else "Drena evakuēta.")
    else:
        p3.append(f"{f['Pacients']} {f['izrakstits']} ar funkcionējošu drenu.")
    stents = v.get("Ureteral stents (RC)") if diversion != "No diversion – nephrostomies" else "No stents"
    if stents == "Discharged with stents":
        date = v.get("Stent removal date (RC)")
        p3.append("Uretera katetri (UK) funkcionē, atstāti" +
                  (f"; to evakuācija plānota {fmt_date(date)}" if date else "") + ".")
    elif stents == "Removed in the ward":
        p3.append("Uretera katetri (UK) evakuēti nodaļā.")
    if diversion == "Orthotopic neobladder":
        date = v.get("Neobladder catheter removal (RC)")
        p3.append(f"{f['Pacients']} {f['izrakstits']} ar urīnpūšļa katetru" +
                  (f", kura evakuācija plānota {fmt_date(date)}" if date else "") + ".")
    if v["Histology (RC)"] == "Pending":
        date = v.get("Histology expected (RC)")
        p3.append("Patohistoloģiskās izmeklēšanas rezultāts izrakstīšanas brīdī vēl nav zināms" +
                  (f" (būs pieejams no {fmt_date(date)})" if date else "") + ".")
    else:
        details = (v.get("Histology details (RC)") or "").strip().rstrip(".")
        p3.append("Patohistoloģiskās izmeklēšanas rezultāts: " + ", ".join(x for x in (details, _rc_stage(v)) if x) + ".")
    p3.append(f"Uzlabojoties vispārējam stāvoklim, {f['pacients']} tiek {f['izrakstits']} turpmākai aprūpei ģimenes "
              "ārsta uzraudzībā.")
    epicrisis = "\n".join(" ".join(p) for p in (p1, p2, p3))

    # Recommendations
    recs = []
    if v.get("Discharge paper") == "Temporary":
        recs.append("Pagaidu izraksts no stacionārā pacienta medicīniskās kartes, galīgais izraksts tiks nosūtīts pa "
                    "pastu, tiklīdz tiks saņemts patohistoloģiskā izmeklējuma slēdziens.")
    recs.append("Atrasties ģimenes ārsta uzraudzībā.")
    recs.append(("Diēta – atbilstoši cukura diabēta rekomendācijām" if v.get("Diabetes diet (RC)") else "Diēta – vispusīga")
                + ", lietot šķidrumu ap 2,0 L/dnn. Ēst regulāri pa mazām porcijām 4–5 reizes dienā. Nelietot asus, "
                  "treknus, kairinošus, vēdera pūšanos veicinošus ēdienus.")
    recs.append("Režīms – saudzējošs 1 mēnesi no operācijas dienas; 6 nedēļas necelt smagumus virs 5 kg (pēcoperācijas "
                "trūces profilaksei); atturēties no pārmērīgas fiziskas slodzes, pirts vai baseina apmeklējuma "
                "3 mēnešus no operācijas dienas.")
    if diversion in RC_STOMA_CARE:
        recs += RC_STOMA_CARE[diversion]
    elif diversion == "No diversion – nephrostomies":
        recs.append("Kopt nefrostomas: pārsiešanu veikt vismaz 2 reizes nedēļā vai biežāk pēc nepieciešamības, ievades "
                    "vietu apstrādāt ar dezinficējošu šķīdumu; urīnsavācējmaisus mainīt vismaz 1 reizi nedēļā; maisus "
                    "nostiprināt, lai nefrostomas netiktu izrautas.")
        nxt = v.get("Next nephrostomy change")
        recs.append("Nākamā nefrostomu maiņa paredzēta PSKUS 4. uroloģijas nodaļas DS " +
                    (fmt_date(nxt) if nxt else "pēc 3 mēnešiem (datumu precizēt)") +
                    ". Ierasties ar ģimenes ārsta nosūtījumu!")
    sutures = v.get("Skin sutures")
    recs.append("Ambulatori turpināt pēcoperācijas brūces apdari, šuves vietas apstrādājot ar Betadine 100 mg/ml vai "
                "brilijantzaļo šķīdumu. Pārsēja maiņa pēc nepieciešamības." +
                (" Diegu evakuācija 14. pēcoperācijas dienā ambulatorā kārtā." if sutures == "Non-absorbable" else
                 " Diegu evakuācija nav nepieciešama (intrakutāna uzsūcošā šuve)."))
    if v["Drain (RC)"] == "Discharged with the drain":
        recs.append("Kopt drenu, sekot izdalījumu daudzumam un raksturam; par drenas evakuāciju lems ārstējošais urologs.")
    meds, advice = split_advice(blocks)
    if meds:
        recs.append("Turpināt iepriekš nozīmēto medikamentozo terapiju, papildus nozīmē:")
        recs += [f"({i}) {b}" for i, b in enumerate(meds, start=1)]
    else:
        recs.append("Turpināt iepriekš nozīmēto medikamentozo terapiju.")
    recs += advice
    if stents == "Discharged with stents":
        date, time = v.get("Stent removal date (RC)"), (v.get("Stent removal time (RC)") or "").strip()
        recs.append("Stacionēšanās " + (fmt_date(date) if date else "pēc 2–4 nedēļām (datumu precizēt)") +
                    (f" plkst. {time}" if time and date else "") +
                    " P. Stradiņa KUS 48. uroloģijas nodaļas dienas stacionārā ar ģimenes ārsta nosūtījumu plānotai UK "
                    "evakuācijai.")
    if diversion == "Orthotopic neobladder" and v.get("Neobladder catheter removal (RC)"):
        recs.append(f"Urīnpūšļa katetra evakuācija plānota {fmt_date(v['Neobladder catheter removal (RC)'])} (pirms tam "
                    "– cistogrāfija neocistis hermētiskuma pārbaudei).")
    recs.append(f"Vizīte pie urologa ambulatorā kārtā {RC_VISIT[v['Urology visit (RC)']]} no operācijas dienas.")
    recs += tests_lines(v, "plain", after_op=True)
    council = v.get("Oncology council")
    if council == "Will be discussed after histology":
        recs.append("Ievērot onkologu konsīlija rekomendācijas – gadījums tiks izskatīts pēc patohistoloģiskās atbildes "
                    "saņemšanas, lēmums sekos.")
    elif council == "Already discussed":
        recs.append(f"Ievērot onkologu konsīlija ({fmt_date(v['Oncology council date'])}) rekomendācijas.")
    if v.get("Smoker (C67)"):
        recs.append(C67_SMOKING)
    recs.append(RC_WARNING)
    return {"clinical_diagnosis": "\n".join(lines), "epicrisis": epicrisis, "recommendations": "\n".join(recs)}


def c67_rc_checks(v):
    blockers, cautions = [], []
    chosen = v.get("_blocks") or []
    previous = v.get("Previous treatment (RC)") or []
    diversion = v.get("Urinary diversion")
    known = v.get("Histology (RC)") == "Known"
    pt, pn = v.get("pT (RC)"), v.get("pN (RC)")
    palliative = "Palliative" in str(v.get("Cystectomy type (man)") or v.get("Cystectomy type (woman)") or "") or \
        str(v.get("Indication (RC)")).startswith("Palliative")
    neoadj = any(x in previous for x in ("Neoadjuvant cisplatin chemotherapy", "Perioperative durvalumab (NIAGARA)"))
    if "fraxiparine-4w" not in chosen:
        cautions.append("EAU MIBC (strong): LMWH for at least 4 weeks after radical cystectomy – choose the 4-week "
                        "Fraxiparine block (or document why not).")
    if "fraxiparine" in chosen and "fraxiparine-4w" in chosen:
        cautions.append("Two Fraxiparine courses are chosen: keep the 4-week one.")
    if v.get("Indication (RC)") == "Muscle-invasive bladder cancer" and not neoadj and not palliative:
        cautions.append("No neoadjuvant treatment recorded: EAU (strong) offers cisplatin-based chemotherapy (or "
                        "gemcitabine/cisplatin + durvalumab) to eligible MIBC patients – note the reason if not given "
                        "(cisplatin-ineligible, patient choice).")
    if known and (pt in ("pT3a", "pT3b", "pT4a", "pT4b") or pn in ("pN1", "pN2", "pN3")):
        if neoadj:
            cautions.append("High-risk pathology after neoadjuvant treatment: EAU (strong) – adjuvant nivolumab "
                            "(EMA: PD-L1 >= 1%) via the oncology council.")
        else:
            cautions.append("pT3/4 and/or pN+ without neoadjuvant treatment: EAU (strong) – adjuvant cisplatin-based "
                            "chemotherapy, or nivolumab if cisplatin-ineligible; oncology council.")
        if v.get("Oncology council") == "Not applicable":
            cautions.append("Advanced pathology but no oncology council plan is set.")
    if known and v.get("Margins (RC)") == "R1":
        cautions.append("Positive margin (R1): EAU (weak) – adjuvant radiotherapy to the cystectomy bed can be offered; "
                        "oncology council.")
    if v.get("Histology (RC)") == "Pending" and v.get("Discharge paper") == "Final":
        cautions.append("Histology is still pending, but the discharge paper is marked as final.")
    if known and v.get("Discharge paper") == "Temporary":
        cautions.append("Histology is already known, but the discharge paper is marked as temporary.")
    if v.get("Ureteral stents (RC)") == "Discharged with stents" and not v.get("Stent removal date (RC)") and \
            diversion != "No diversion – nephrostomies":
        cautions.append("Stents left in: give the date for UK removal in 48.")
    if diversion == "Orthotopic neobladder" and v.get("Urethrectomy (RC)"):
        blockers.append("Neobladder with urethrectomy is not possible – check the diversion.")
    if diversion == "Orthotopic neobladder":
        cautions.append("Neobladder: EAU (strong) – not with tumour in the urethra / positive urethral margin; check the "
                        "frozen section and the final histology.")
    if v.get("Lymph node dissection (RC)") == "Not performed" and not palliative:
        cautions.append("No lymph node dissection: EAU (strong) – standard LND is part of radical cystectomy.")
    if v.get("Lymph node dissection (RC)") == "Extended bilateral":
        cautions.append("EAU (strong): standard LND is enough – extended LND does not improve survival.")
    if "amoxiclav-7-rc" in chosen:
        cautions.append("Amoksiklav: conduit urine is always colonised – treat only a symptomatic infection "
                        "(EAU / Campbell-Walsh-Wein); check the indication.")
    if v.get("_sex") == "Woman" and v.get("Cystectomy type (woman)") == "Organ-sparing (uterus and vagina preserved)":
        cautions.append("Organ-sparing cystectomy: EAU – only without tumour near the preserved organs (no trigone / "
                        "bladder neck tumour, no CIS there).")
    if "With complications" in (v.get("Intraoperative course"), v.get("Postoperative course")):
        cautions.append("Complications are recorded: check that the recommendations cover them.")
    return blockers, cautions


BUILDERS["c67_rc"] = c67_rc_parts


# ---------------------------------------------------------------
# 2u. EPIDIDYMO-ORCHITIS (N45), testicular / epididymal abscess – NMC or 4. Urology
# ---------------------------------------------------------------
# EAU 2026 (strong): midstream urine and first-void urine for culture and NAAT; antibiotics against C. trachomatis and
# Enterobacterales in young sexually active men, Enterobacterales only in older men without sexual risk; ceftriaxone
# 500 mg once if gonorrhoea is likely; follow national rules on reporting and partner treatment. Sudden severe pain:
# urgent exploration for torsion (do not wait for ultrasound – IUSTI 2024). Supportive: rest, scrotal elevation,
# analgesia. Review at 3 days if not better; ultrasound for persistent swelling (abscess, infarction, tumour).
# Campbell-Walsh-Wein: abscess or pyocele needs drainage – usually orchiectomy for a testicular abscess,
# epididymectomy / drainage for an epididymal abscess; older men – look for outlet obstruction; an "orchitis" that
# does not resolve may be a tumour.

EO_SIDE = {  # (adjective "kreisās puses", locative "kreisajā", testis nominative)
    "Right": ("Labās puses", "labās puses", "labajā", "labais"),
    "Left": ("Kreisās puses", "kreisās puses", "kreisajā", "kreisais"),
    "Both": ("Abpusējs", "abpusēja", "abās", "abi"),
}
EO_SITE = {"Epididymitis": "epididimīts", "Epididymo-orchitis": "epididimoorhīts", "Orchitis": "orhīts"}
EO_DURATION = {"1 day": "1 dienu", "2–3 days": "2–3 dienas", "4–7 days": "4–7 dienas",
               "More than a week": "ilgāk par nedēļu"}
EO_SYMPTOMS = {"Fever": "temperatūras pacelumiem", "Dysuria / frequency": "dizūriju",
               "Urethral discharge": "izdalījumiem no urīnizvadkanāla",
               "Scrotal redness / swelling": "sēklinieku maisiņa apsārtumu un tūsku"}
EO_RISK = {"Indwelling catheter / epicystostomy": "urīnpūšļa katetrs / epicistostoma",
           "Recent urological procedure": "nesen veikta uroloģiska manipulācija",
           "Bladder outlet obstruction (BPH)": "LPH ar infravezikālu obstrukciju", "Diabetes": "cukura diabēts",
           "Immunosuppression": "imūnsupresija"}
EO_US = {"Epididymitis (hyperaemia)": "epididimīts (piedēklis palielināts, hiperēmisks)", "Orchitis": "orhīts",
         "Reactive hydrocele": "reaktīva hidrocēle", "Abscess of the epididymis": "piedēkļa abscess",
         "Abscess of the testis": "sēklinieka abscess", "Pyocele": "piocēle",
         "Testicular infarction / ischaemia": "sēklinieka išēmija / infarkts",
         "Normal testicular blood flow": "sēklinieka asinsplūsma saglabāta"}
EO_LABS = {"No inflammatory changes": "Laboratoriski bez iekaisuma izmaiņām.",
           "Raised CRP / leukocytosis": "Laboratoriski – paaugstināts CRO, leikocitoze.",
           "Raised CRP and creatinine": "Laboratoriski – paaugstināts CRO, leikocitoze, paaugstināts kreatinīna "
                                        "līmenis."}
EO_OP = {  # (after "veic ...", after "St. pēc ...")
    "Orchiectomy (scrotal)": ("orhektomiju", "orhektomijas"),
    "Inguinal orchifuniculectomy": ("orhofunikulektomiju", "orhofunikulektomijas"),
    "Epididymectomy": ("epididimektomiju", "epididimektomijas"),
    "Abscess drainage / revision": ("sēklinieku maisiņa revīziju, abscesa atvēršanu un drenāžu",
                                    "sēklinieku maisiņa revīzijas, abscesa atvēršanas un drenāžas"),
    "Scrotal exploration": ("sēklinieku maisiņa revīziju", "sēklinieku maisiņa revīzijas"),
}
EO_VISIT = {"Urologist in 2 weeks": "Kontrole pie urologa ambulatori pēc 2 nedēļām.",
            "Family doctor in 1 week": "Kontrole pie ģimenes ārsta pēc 1 nedēļas.",
            "4. Urology (sutures)": "Kontroles vizīte PSKUS 4. nodaļā 10.–14. pēcoperācijas dienā (diegu "
                                    "evakuācija).",
            "Not needed": ""}
EO_SUPPORT = ("Pirmās dienas ievērot gultas režīmu, sēklinieku maisiņu turēt paceltu (cieši pieguļoša apakšveļa vai "
              "suspensorijs); aukstas kompreses caur drānu 15–20 minūtes vairākas reizes dienā.")
EO_STI = ("Dzimumpartneri(-us) informēt par nepieciešamību izmeklēties un ārstēties; atturēties no dzimumdzīves līdz "
          "ārstēšanas beigām un partnera ārstēšanai (vai lietot prezervatīvu).")
EO_WARNING = ("Ja 72 stundu laikā stāvoklis neuzlabojas, pastiprinās sāpes vai tūska, paaugstinās ķermeņa temperatūra, "
              "parādās drebuļi vai strutaini izdalījumi, nekavējoties vērsties NMC.")


def _eo_abscess(v):
    us = v.get("Scrotal US (EO)") or []
    return [x for x in ("Abscess of the testis", "Abscess of the epididymis", "Pyocele") if x in us]


def _eo_sti(v):
    return v.get("Likely source (EO)") in ("Sexually transmitted", "Both possible")


def _eo_diagnosis(v, ward):
    adj = EO_SIDE[v["Side (EO)"]]
    site = v["Inflammation site (EO)"]
    lines = []
    culture = (v.get("Culture result (EO)") or "").strip().rstrip(".") if ward else ""
    enteric = v.get("Likely source (EO)") in ("Enteric / urinary (UTI, obstruction, catheter, instrumentation)",
                                              "Both possible")
    if enteric and (culture or "Leukocyturia" in (v.get("Urine tests (EO)") or [])):
        lines.append("Komplicēta urīnceļu infekcija" + (f" ({culture})" if culture else "") + ".")
    abscess = _eo_abscess(v)
    if v.get("Likely source (EO)") == "Mumps (viral)":
        main = f"{adj[0]} epidēmiskā parotīta orhīts."
    elif "Abscess of the testis" in abscess:
        main = f"{adj[0]} strutains orhīts ar abscedēšanos."
    elif "Abscess of the epididymis" in abscess:
        main = f"{adj[0]} {EO_SITE[site]} ar piedēkļa abscesu."
    else:
        main = f"{adj[0]} {EO_SITE[site]}."
    lines.append(main)
    if "Pyocele" in abscess:
        lines.append(f"{adj[0]} piocēle.")
    if ward and v.get("Treatment (EO)") == "Surgery":
        nom, gen = EO_OP[v["Operation (EO)"]]
        lines.append(f"St. pēc {adj[1]} {gen} ({fmt_date(v['Procedure date'])}).")
    if not blank(v.get("Other diagnoses (optional)")):
        lines.append(v["Other diagnoses (optional)"].strip().rstrip(".") + ".")
    return "\n".join(lines)


def _eo_nmc_sentences(f, v):
    adj = EO_SIDE[v["Side (EO)"]]
    both = v["Side (EO)"] == "Both"
    complaints = [f"sāpēm {'abās sēklinieku maisiņa pusēs' if both else adj[2] + ' sēkliniekā'}"]
    complaints += [EO_SYMPTOMS[x] for x in (v.get("Symptoms (EO)") or []) if x in EO_SYMPTOMS]
    s = [f"{f['Pacients']} {f['iestajies']} PSKUS NMC ar sūdzībām par {lv_list(complaints)} "
         f"{EO_DURATION[v['Symptom duration (EO)']]}."]
    risks = [EO_RISK[x] for x in (v.get("Risk factors (EO)") or []) if x in EO_RISK]
    if risks:
        s.append(f"Anamnēzē – {lv_list(risks)}.")
    site = v["Inflammation site (EO)"]
    organ = {"Epididymitis": ("sēklinieka piedēklis", "palielināts, palpatori sāpīgs"),
             "Epididymo-orchitis": ("sēklinieks un piedēklis", "palielināti, palpatori sāpīgi"),
             "Orchitis": ("sēklinieks", "palielināts, palpatori sāpīgs")}[site]
    if both:
        exam = f"abpusēji {organ[0].replace('sēklinieks', 'sēklinieki').replace('piedēklis', 'piedēkļi')} " \
               f"{organ[1].replace('palielināts', 'palielināti').replace('sāpīgs', 'sāpīgi')}"
    else:
        exam = f"{adj[3]} {organ[0]} {organ[1]}"
    s.append(f"Konsultē urologs. NMC etapā vispusīgi {f['izmekleta']}. Objektīvi: {exam}.")
    us = [EO_US[x] for x in (v.get("Scrotal US (EO)") or []) if x in EO_US]
    if us:
        s.append("Sēklinieku maisiņa USS ar doplerogrāfiju: " + ", ".join(us) + ".")
    if v.get("Labs (EO)") in EO_LABS:
        s.append(EO_LABS[v["Labs (EO)"]])
    tests = v.get("Urine tests (EO)") or []
    taken = []
    if "Midstream urine culture" in tests:
        taken.append("urīna uzsējums")
    if any(x.startswith("First-void") for x in tests):
        taken.append("pirmās porcijas urīns NAAT izmeklējumam (C. trachomatis, N. gonorrhoeae, M. genitalium)")
    if "Leukocyturia" in tests:
        s.append("Urīna analīzē – leikocitūrija.")
    if taken:
        s.append("Paņemts " + lv_list(taken) + ".")
    if not blank(v.get("Other NMC findings (optional)")):
        text = v["Other NMC findings (optional)"].strip()
        s.append(text if text.endswith(".") else text + ".")
    return s


def _eo_common_recs(v, blocks, bullet):
    recs = []
    if blocks:
        recs.append(bullet + "Nozīmē lietot:")
        recs += ["    o " + b for b in blocks]
    if v.get("Likely source (EO)") == "Mumps (viral)":
        recs.append(bullet + "Antibakteriāla terapija nav nepieciešama (vīrusa izraisīts orhīts); simptomātiska terapija.")
    tests = v.get("Urine tests (EO)") or []
    taken = (["urīna uzsējuma"] if "Midstream urine culture" in tests else []) + \
        (["NAAT"] if any(x.startswith("First-void") for x in tests) else [])
    known = v.get("_department") == "4" and not blank(v.get("Culture result (EO)"))
    if taken and not (known and taken == ["urīna uzsējuma"]):
        joined = lv_list(taken)
        recs.append(bullet + f"{joined[0].upper() + joined[1:]} rezultātus var aplūkot Datamed.lv pēc 3–5 dienām; "
                    "antibakteriālo terapiju koriģēt atbilstoši rezultātiem.")
    if _eo_sti(v):
        recs.append(bullet + EO_STI)
    if v.get("Likely source (EO)").startswith("Enteric") and "Bladder outlet obstruction (BPH)" not in \
            (v.get("Risk factors (EO)") or []) and "Indwelling catheter / epicystostomy" not in (v.get("Risk factors (EO)") or []):
        recs.append(bullet + "Ambulatori urologa konsultācija – izvērtēt urīnpūšļa iztukšošanos (infravezikālu obstrukciju) "
                    "kā iespējamu infekcijas cēloni.")
    return recs


def n45_nmc_parts(sex, v, blocks=()):
    f = SEX_FORMS["Man"]
    s = _eo_nmc_sentences(f, v)
    if v.get("Ceftriaxone given at NMC"):
        s.append("Ievadīts Sol. Ceftriaxoni 500 mg i/m vienreizēji.")
    s.append("Uzsākta antibakteriāla terapija." if v.get("Likely source (EO)") != "Mumps (viral)" else
             "Uzsākta simptomātiska terapija.")
    s.append("Dotajā brīdī nav indikāciju neatliekamai stacionēšanai. Vispārēji apmierinošā stāvoklī un ar "
             "rekomendācijām izraksta tālākai ambulatorai ārstēšanai.")
    numbered = make_counter()
    recs = [numbered("Atrasties ģimenes ārsta uzraudzībā."), numbered(EO_SUPPORT)]
    for line in _eo_common_recs(v, blocks, ""):
        recs.append(line if line.startswith("    o ") else numbered(line))
    if EO_VISIT.get(v["Control visit (EO)"]):
        recs.append(numbered(EO_VISIT[v["Control visit (EO)"]]))
    recs += tests_lines(v, "numbered", numbered)
    recs.append(numbered(EO_WARNING))
    recs = [r.replace("    o ", "    ") for r in recs]
    return {"clinical_diagnosis": _eo_diagnosis(v, False), "epicrisis": " ".join(s), "recommendations": "\n".join(recs)}


def n45_ward_parts(sex, v, blocks=()):
    f = SEX_FORMS["Man"]
    adj = EO_SIDE[v["Side (EO)"]]
    s = _eo_nmc_sentences(f, v)
    s.append("Uzsākta antibakteriāla terapija. Stacionēts 4. uroloģijas nodaļā.")
    p2 = []
    abx = (v.get("IV antibiotic (EO)") or "").strip().rstrip(".")
    surgery = v.get("Treatment (EO)") == "Surgery"
    if not surgery:
        p2.append("Nodaļā turpināta i/v antibakteriāla terapija" + (f" ({abx})" if abx else "") +
                  " un simptomātiska terapija. Dinamikā iekaisuma rādītāji mazinās, sāpes un tūska mazinās.")
    else:
        if abx:
            p2.append(f"Nodaļā turpināta i/v antibakteriāla terapija ({abx}).")
        if v.get("After failed conservative therapy"):
            p2.append("Uz konservatīvas terapijas fona stāvoklis neuzlabojas, palielinās tūska un iekaisuma rādītāji. "
                      "Indicēta operatīva ārstēšana.")
        nom, gen = EO_OP[v["Operation (EO)"]]
        anesthesia = {"Spinal": "spinālā anestēzijā", "General": "vispārējā anestēzijā"}[v["Anesthesia (EO)"]]
        postop = ("bez komplikācijām" if v.get("Postoperative course (EO)") != "With complications" else
                  "ar komplikācijām: " + (v.get("Postoperative complications (EO)") or "").strip().rstrip("."))
        p2.append(f"{fmt_date(v['Procedure date'])} {anesthesia} veic {adj[1]} {nom}. Operācija bez sarežģījumiem. "
                  f"Pēcoperācijas periods {postop}.")
        drain = v.get("Drain (EO)")
        if drain == "Removed in the ward":
            p2.append("Drena evakuēta.")
        elif drain == "Discharged with drain":
            p2.append(f"{f['Pacients']} {f['izrakstits']} ar drenu.")
        p2.append("Dinamikā iekaisuma rādītāji mazinās.")
    culture = (v.get("Culture result (EO)") or "").strip().rstrip(".")
    if culture:
        p2.append(f"Uzsējumā aug {culture}. Pielāgota antibakteriāla terapija.")
    hist = v.get("Histology (EO)") if surgery else None
    if hist == "Pending":
        p2.append("Histoloģiskās izmeklēšanas rezultāts izrakstīšanas brīdī vēl nav zināms.")
    elif hist == "Known":
        p2.append("Histoloģiski: " + v["Histology result (EO)"].strip().rstrip(".") + ".")
    p2.append(f"Ar pozitīvu dinamiku {f['izrakstits']} ambulatori.")
    epicrisis = " ".join(s) + "\n" + " ".join(p2)

    weeks = WEEKS_TEXT[v["Activity restriction"]]
    recs = ["• Atrasties ģimenes ārsta uzraudzībā.",
            f"• Režīms – saudzējošs {weeks}" + (" no operācijas dienas" if surgery else "") +
            ", atturēties no pārmērīgas fiziskas slodzes, pirts vai baseina apmeklējuma 1 mēnesi" +
            (" no operācijas dienas." if surgery else "."),
            "• Diēta – vispusīga, lietot šķidrumu ap 2,0 L/dnn."]
    recs += _eo_common_recs(v, blocks, "• ")
    if surgery:
        recs.append("• Ambulatori turpināt pēcoperācijas brūces apdari, šuves vietas apstrādājot ar Braunol, pārsēja "
                    "maiņa pēc nepieciešamības.")
        recs.append("• Diegu evakuācija 10.–14. pēcoperācijas dienā ambulatori.")
        if v.get("Drain (EO)") == "Discharged with drain":
            recs.append("• Kopt drenu; par drenas evakuāciju lems ārstējošais urologs.")
        if hist == "Pending":
            recs.append("• Histoloģiskās izmeklēšanas rezultātu izvērtēs ārstējošais urologs.")
    else:
        recs.append("• " + EO_SUPPORT)
    if EO_VISIT.get(v["Control visit (EO)"]):
        recs.append("• " + EO_VISIT[v["Control visit (EO)"]])
    recs += tests_lines(v, "bullet")
    recs.append("• Stāvokļa pasliktināšanās gadījumā griezties NMC.")
    return {"clinical_diagnosis": _eo_diagnosis(v, True), "epicrisis": epicrisis, "recommendations": "\n".join(recs)}


def n45_checks(v):
    blockers, cautions = [], []
    nmc = v.get("_department") == "NMC"
    us = v.get("Scrotal US (EO)") or []
    chosen = v.get("_blocks") or []
    tests = v.get("Urine tests (EO)") or []
    source = v.get("Likely source (EO)")
    abscess = _eo_abscess(v)
    if "Testicular infarction / ischaemia" in us:
        (blockers if nmc else cautions).append("Testicular ischaemia / infarction on ultrasound: exclude torsion – "
                                              "urgent scrotal exploration; infarction may need orchiectomy.")
    if v.get("Onset (EO)") == "Sudden" and "Normal testicular blood flow" not in us:
        blockers.append("Sudden onset without documented testicular blood flow: torsion must be excluded – urgent "
                        "scrotal exploration (do not wait for ultrasound; salvage window ~6 hours).")
    elif v.get("Onset (EO)") == "Sudden":
        cautions.append("Sudden onset: Doppler flow does not fully exclude torsion (intermittent / partial) – explore if "
                        "in doubt, especially under 25 years.")
    if nmc:
        if abscess:
            blockers.append("Abscess or pyocele on ultrasound: admit for drainage / orchiectomy (Campbell-Walsh-Wein).")
        if v.get("Sepsis signs (EO)"):
            blockers.append("Signs of sepsis: admit for IV antibiotics.")
        if "Fever" in (v.get("Symptoms (EO)") or []) and v.get("Labs (EO)") in ("Raised CRP / leukocytosis",
                                                                                 "Raised CRP and creatinine"):
            cautions.append("Fever with raised CRP: consider admission (IV antibiotics), especially with diabetes, "
                            "immunosuppression or a catheter.")
    else:
        if abscess and v.get("Treatment (EO)") != "Surgery":
            cautions.append("Abscess / pyocele treated conservatively: usually needs drainage or orchiectomy – if "
                            "conservative, repeat ultrasound and watch closely.")
        if v.get("Treatment (EO)") == "Surgery" and v.get("Operation (EO)") in ("Orchiectomy (scrotal)",
                                                                              "Inguinal orchifuniculectomy",
                                                                              "Epididymectomy") \
                and v.get("Histology (EO)") == "Not sent":
            cautions.append("Send the specimen for histology – a tumour can present as orchitis / abscess.")
    if "Midstream urine culture" not in tests:
        cautions.append("EAU (strong): take a midstream urine culture.")
    if source in ("Sexually transmitted", "Both possible") and not any(x.startswith("First-void") for x in tests):
        cautions.append("Possible STI: EAU (strong) – first-void urine NAAT for chlamydia and gonorrhoea "
                        "(IUSTI: also M. genitalium).")
    if source in ("Sexually transmitted", "Both possible"):
        if not any(b in chosen for b in ("doxycycline", "levofloxacin", "moxifloxacin")):
            cautions.append("STI possible: choose an antibiotic active against C. trachomatis (doxycycline 14 days, or "
                            "levofloxacin if an enteric source is also possible).")
        if "Urethral discharge" in (v.get("Symptoms (EO)") or []) and not v.get("Ceftriaxone given at NMC") and nmc:
            cautions.append("Urethral discharge – gonorrhoea likely: EAU (strong) – ceftriaxone 500 mg i/m or i/v once "
                            "plus an anti-chlamydial drug.")
        cautions.append("STI: follow national rules on reporting and partner notification (EAU, strong).")
    if source and source.startswith("Enteric") and not any(b in chosen for b in ("levofloxacin", "cotrimoxazole-eo",
                                                                              "amoxiclav-eo", "cefuroxime-eo",
                                                                              "clindamycin-eo")):
        cautions.append("Enteric source: choose an antibiotic active against Enterobacterales (levofloxacin 10 days "
                        "first choice, or by culture).")
    if source == "Mumps (viral)" and any(b in chosen for b in _eo_abx + ["doxycycline"]):
        cautions.append("Mumps orchitis is viral: antibiotics are not needed. It is a notifiable infection.")
    if "Indwelling catheter / epicystostomy" in (v.get("Risk factors (EO)") or []):
        cautions.append("Catheter-associated: change (or remove) the catheter and take the culture from the new one; "
                        "a suprapubic catheter reduces the risk with long-term drainage.")
    if "levofloxacin" in chosen or "moxifloxacin" in chosen:
        cautions.append("Fluoroquinolone: EMA warnings – explain tendon / nerve / aortic risks and stop at the first "
                        "sign of tendon pain.")
    return blockers, cautions


BUILDERS["n45_nmc"] = n45_nmc_parts


# ---------------------------------------------------------------
# 2v. ACUTE BACTERIAL PROSTATITIS (N41.0), prostatic abscess – NMC or 4. Urology
# ---------------------------------------------------------------
# EAU 2026: no prostatic massage (strong); a gentle DRE is part of the examination; midstream urine dipstick and
# culture, blood culture and blood count (weak); TRUS in selected cases to exclude an abscess (weak); treat as a
# systemic UTI and manage complicating factors (strong). NICE / HSE: oral fluoroquinolone 14 days, review, up to
# 4 weeks; refer with retention, no response, or systemic illness (abscess). Prostatic abscess (systematic review
# 2020): < 2 cm may settle on antibiotics with close monitoring; otherwise TRUS-guided aspiration / drainage, or
# transurethral deroofing (unwell, multiloculated, outlet obstruction, failed aspiration). Campbell-Walsh-Wein:
# with retention a suprapubic tube was classically preferred (a short small-calibre urethral catheter is
# acceptable); PSA rises with prostatitis – do not interpret it for weeks; recurrent episodes – look for chronic
# bacterial prostatitis (Meares–Stamey test, EAU strong) and outlet obstruction.

AP_SYMPTOMS = {"Fever": "paaugstinātu temperatūru virs 38 °C", "Chills": "drebuļiem",
               "Dysuria / frequency": "dizūriskām sūdzībām", "Perineal / suprapubic pain": "sāpēm starpenē un virs "
                                                                                       "simfīzes",
               "Difficulty voiding / weak stream": "apgrūtinātu urināciju", "Urinary retention": "nespēju urinēt"}
AP_DURATION = {"1 day": "1 dienu", "2–3 days": "2–3 dienas", "4–7 days": "4–7 dienas", "About 1 week": "~1 nedēļu",
               "More than a week": "ilgāk par nedēļu"}
AP_RISK = {"Recurrent prostatitis": "recidivējošs prostatīts", "BPH / outlet obstruction": "LPH",
           "Diabetes": "cukura diabēts", "Immunosuppression": "imūnsupresija", "Indwelling catheter": "urīnpūšļa katetrs"}
AP_DRE = {"Enlarged, very tender, warm": "Per rectum: prostata palielināta, izteikti sāpīga, silta.",
          "Tender, fluctuating area (abscess?)": "Per rectum: prostata palielināta, izteikti sāpīga, palpējama "
                                                 "fluktuācija – jādomā par abscesu."}
AP_LABS = {"Mildly raised CRP": "Asins analīžu rezultātos nedaudz paaugstināti iekaisuma rādītāji.",
           "Raised CRP / leukocytosis": "Asins analīzēs – paaugstināts CRO, leikocitoze.",
           "Raised CRP and creatinine": "Asins analīzēs – paaugstināts CRO, leikocitoze, paaugstināts kreatinīna "
                                        "līmenis."}
AP_IMAGING = {"Ultrasound – no residual urine, no abscess": "USS – bez atlieku urīna.",
              "Ultrasound – residual urine": "USS – atlieku urīns.",
              "TRUS – prostatitis, no abscess": "TRUS – prostatīta aina, norādes par abscedēšanos nav.",
              "TRUS / CT / MRI – prostatic abscess": "Attēldiagnostikā – prostatas abscess"}
AP_ABSCESS_OP = {"TRUS-guided aspiration / drainage": ("TRUS kontrolē veikta prostatas abscesa punkcija un drenāža",
                                                       "TRUS vadītas prostatas abscesa punkcijas un drenāžas"),
                 "Transurethral drainage (deroofing)": ("veikta transuretrāla prostatas abscesa atvēršana un drenāža",
                                                        "transuretrālas prostatas abscesa drenāžas")}
AP_VISIT = {"Urologist in 2–3 weeks": "Kontrole pie urologa ambulatori pēc 2–3 nedēļām (ārstēšanas ilguma "
                                      "izvērtēšanai).",
            "Family doctor in 1 week": "Kontrole pie ģimenes ārsta pēc 1 nedēļas.", "Not needed": ""}
AP_DURATION_LINE = ("Antibakteriālo terapiju lietot visu nozīmēto kursu (kopā 2–4 nedēļas), arī ja sūdzības mazinās "
                    "agrāk; terapijas ilgumu pēc 2 nedēļām izvērtēs ārstējošais ārsts.")
AP_BOWEL = "Izvairīties no aizcietējumiem un ilgstošas sēdēšanas vai riteņbraukšanas līdz izveseļošanai."
AP_WARNING = ("Ja 48–72 stundu laikā stāvoklis neuzlabojas, saglabājas drudzis, parādās drebuļi, nespēja urinēt vai "
              "stipras sāpes starpenē, nekavējoties vērsties NMC.")


def _ap_abscess(v):
    return v.get("Imaging (AP)") == "TRUS / CT / MRI – prostatic abscess"


def _ap_retention(v):
    return v.get("Retention drainage (AP)") in ("Urethral catheter (small calibre)", "Suprapubic catheter") or \
        "Urinary retention" in (v.get("Symptoms (AP)") or [])


def _ap_sentences(f, v):
    symptoms = [AP_SYMPTOMS[x] for x in (v.get("Symptoms (AP)") or []) if x in AP_SYMPTOMS]
    s = [f"{f['Pacients']} {f['iestajies']} PSKUS NMC akūtā kārtā sakarā ar "
         f"{lv_list(symptoms) or 'dizūriskām sūdzībām'} {AP_DURATION[v['Symptom duration (AP)']]}."]
    context = {"After prostate biopsy": "Simptomi radušies pēc prostatas biopsijas.",
               "After catheter / instrumentation": "Simptomi radušies pēc urīnpūšļa katetrizācijas / uroloģiskas "
                                                   "manipulācijas."}.get(v.get("Context (AP)"))
    if context:
        s.append(context)
    risks = [AP_RISK[x] for x in (v.get("Risk factors (AP)") or []) if x in AP_RISK]
    if risks:
        s.append(f"Anamnēzē – {lv_list(risks)}.")
    s.append(f"NMC etapā vispusīgi {f['izmekleta']}.")
    if v.get("DRE (AP)") in AP_DRE:
        s.append(AP_DRE[v["DRE (AP)"]])
    if v.get("Labs (AP)") in AP_LABS:
        s.append(AP_LABS[v["Labs (AP)"]])
    tests = v.get("Urine tests (AP)") or []
    if "Urinalysis: leukocytes / nitrite" in tests:
        s.append("Urīna analīzē – leikocitūrija, nitrīti.")
    taken = [x for x, t in (("Midstream urine culture", "urīna uzsējums"), ("Blood culture", "asins uzsējums"),
                            ("First-void urine NAAT (STI)", "pirmās porcijas urīns NAAT izmeklējumam")) if x in tests]
    names = {"Midstream urine culture": "urīna uzsējums", "Blood culture": "asins uzsējums",
             "First-void urine NAAT (STI)": "pirmās porcijas urīns NAAT izmeklējumam"}
    if taken:
        s.append("Paņemts " + lv_list([names[x] for x in taken]) + ".")
    img = v.get("Imaging (AP)")
    if img in AP_IMAGING:
        text = AP_IMAGING[img]
        if _ap_abscess(v):
            size = _num(v, "Abscess size (cm) (AP)")
            text += (f" (~{_comma(size)} cm)." if size else ".")
        s.append(text)
    drain = v.get("Retention drainage (AP)")
    if drain in ("Urethral catheter (small calibre)", "Suprapubic catheter"):
        vol = _num(v, "Drained volume (AP)")
        s.append(("Sakarā ar urīna retenci ievieto Foley ilgkatetru" if drain.startswith("Urethral") else
                  "Sakarā ar urīna retenci veic perkutānu epicistostomiju") +
                 (f", izdalās {_comma(vol)} ml urīna." if vol else "."))
    if not blank(v.get("Other NMC findings (optional)")):
        text = v["Other NMC findings (optional)"].strip()
        s.append(text if text.endswith(".") else text + ".")
    return s


def _ap_diagnosis(v, ward):
    lines = []
    culture = (v.get("Culture result (AP)") or "").strip().rstrip(".") if ward else ""
    if ward:
        lines.append("Komplicēta urīnceļu infekcija" + (f" ({culture})" if culture else "") + ".")
    lines.append("Akūts bakteriāls prostatīts" + (" ar prostatas abscesu" if _ap_abscess(v) else "") + ".")
    if _ap_retention(v):
        lines.append("Urīna retence.")
    if "Recurrent prostatitis" in (v.get("Risk factors (AP)") or []):
        lines.append("Recidivējošs prostatīts anamnēzē.")
    if v.get("Context (AP)") == "After prostate biopsy":
        lines.append("Stāvoklis pēc prostatas biopsijas.")
    if ward and v.get("Abscess treatment (AP)") in AP_ABSCESS_OP and _ap_abscess(v):
        lines.append(f"St. pēc {AP_ABSCESS_OP[v['Abscess treatment (AP)']][1]} "
                     f"({fmt_date(v['Abscess procedure date (AP)'])}).")
    if not blank(v.get("Other diagnoses (optional)")):
        lines.append(v["Other diagnoses (optional)"].strip().rstrip(".") + ".")
    return "\n".join(lines)


def _ap_extra_recs(v):
    out = []
    if "STI risk (young, sexually active)" in (v.get("Risk factors (AP)") or []):
        out.append(EO_STI)
    if "Recurrent prostatitis" in (v.get("Risk factors (AP)") or []):
        out.append("Pēc akūtās epizodes izārstēšanas – urologa konsultācija hroniska bakteriāla prostatīta un "
                   "infravezikālas obstrukcijas izvērtēšanai.")
    return out


def n41_nmc_parts(sex, v, blocks=()):
    f = SEX_FORMS["Man"]
    s = _ap_sentences(f, v)
    first = (v.get("First IV dose at NMC (AP)") or "").strip().rstrip(".")
    if first:
        s.append(f"Ievadīta pirmā antibiotikas deva – {first}.")
    s.append("Dotajā brīdī nav indikāciju neatliekamai stacionēšanai uroloģijas nodaļā. Vispārēji apmierinošā stāvoklī "
             "un ar rekomendācijām izraksta tālākai ambulatorai ārstēšanai.")
    numbered = make_counter()
    recs = [numbered("Atrasties ģimenes ārsta un ārstējošā urologa uzraudzībā."),
            numbered("Lietot daudz šķidruma, vismaz 1,5–2,5 l/DN.")]
    if blocks:
        recs.append(numbered("Terapijā lietot:"))
        recs += ["    " + b for b in blocks]
    recs.append(numbered(AP_DURATION_LINE))
    if any(x in (v.get("Urine tests (AP)") or []) for x in ("Midstream urine culture", "Blood culture",
                                                             "First-void urine NAAT (STI)")):
        recs.append(numbered("Laboratorisko izmeklējumu rezultātus var aplūkot Datamed.lv; urīna uzsējuma rezultāti pēc "
                             "3–5 dienām – antibakteriālo terapiju koriģēt atbilstoši jutībai."))
    recs.append(numbered(AP_BOWEL))
    recs += [numbered(x) for x in _ap_extra_recs(v)]
    if AP_VISIT.get(v["Control visit (AP)"]):
        recs.append(numbered(AP_VISIT[v["Control visit (AP)"]]))
    recs += tests_lines(v, "numbered", numbered)
    recs.append(numbered(AP_WARNING))
    return {"clinical_diagnosis": _ap_diagnosis(v, False), "epicrisis": " ".join(s), "recommendations": "\n".join(recs)}


def n41_ward_parts(sex, v, blocks=()):
    f = SEX_FORMS["Man"]
    s = _ap_sentences(f, v)
    s.append("Konstatēta komplicēta urīnceļu infekcija, akūts prostatīts, sakarā ar ko stacionēts 4. uroloģijas nodaļā.")
    p2 = []
    abx = (v.get("IV antibiotic (AP)") or "").strip().rstrip(".")
    p2.append("Nodaļā nozīmēta i/v antibakteriāla terapija" + (f" ({abx})" if abx else "") + ".")
    if _ap_abscess(v):
        treat = v.get("Abscess treatment (AP)")
        if treat in AP_ABSCESS_OP:
            p2.append(f"{fmt_date(v['Abscess procedure date (AP)'])} {AP_ABSCESS_OP[treat][0]}.")
        else:
            p2.append("Prostatas abscess ārstēts konservatīvi ar antibakteriālu terapiju, dinamikā kontroles TRUS.")
    culture = (v.get("Culture result (AP)") or "").strip().rstrip(".")
    if culture:
        p2.append(f"Uzsējumā aug {culture}; antibakteriāla terapija pielāgota atbilstoši jutībai.")
    p2.append("Dinamikā normalizējas ķermeņa temperatūra, analīzēs pozitīva dinamika, sāpes mazinās."
              if v.get("Ward course (AP)") == "Fever resolved, CRP falling" else "Periods nodaļā bez īpatnībām.")
    if v.get("Retention drainage (AP)") in ("Urethral catheter (small calibre)", "Suprapubic catheter"):
        if v.get("Catheter at discharge (AP)") == "Discharged with catheter":
            date = v.get("Catheter removal date (AP)")
            p2.append(f"{f['Pacients']} {f['izrakstits']} ar " +
                      ("urīnpūšļa ilgkatetru" if v["Retention drainage (AP)"].startswith("Urethral") else
                       "epicistostomu") + (f", kura evakuācija plānota {fmt_date(date)}" if date else "") + ".")
        else:
            p2.append("Katetrs evakuēts, urinācija spontāna, bez atlieku urīna.")
    else:
        p2.append("Urinācija spontāna.")
    p2.append(f"Vispārējā apmierinošā stāvoklī {f['izrakstits']} tālākai terapijai ambulatori.")
    epicrisis = " ".join(s) + "\n" + " ".join(p2)

    recs = ["• Atrasties ģimenes ārsta un ārstējošā urologa uzraudzībā.",
            "• Diēta – vispusīga, lietot šķidrumu 1,5–2,5 L/dnn."]
    if blocks:
        recs.append("• Nozīmē lietot:")
        recs += ["    o " + b for b in blocks]
    recs.append("• " + AP_DURATION_LINE)
    if v.get("Catheter at discharge (AP)") == "Discharged with catheter" and \
            v.get("Retention drainage (AP)") in ("Urethral catheter (small calibre)", "Suprapubic catheter"):
        recs.append("• Kopt " + ("urīnpūšļa katetru" if v["Retention drainage (AP)"].startswith("Urethral") else
                                 "epicistostomu") + "; urīna savācējmaisa maiņa ik pēc 3–7 dienām.")
        date = v.get("Catheter removal date (AP)")
        recs.append("• Katetra evakuācija (urinācijas mēģinājums) " +
                    (f"plānota {fmt_date(date)}." if date else "pēc infekcijas mazināšanās – lems ārstējošais urologs."))
    if blank(v.get("Culture result (AP)")) and any(x in (v.get("Urine tests (AP)") or []) for x in
                                                   ("Midstream urine culture", "Blood culture")):
        recs.append("• Uzsējumu rezultātus var aplūkot Datamed.lv; antibakteriālo terapiju koriģēt atbilstoši jutībai.")
    recs.append("• " + AP_BOWEL)
    recs += ["• " + x for x in _ap_extra_recs(v)]
    if AP_VISIT.get(v["Control visit (AP)"]):
        recs.append("• " + AP_VISIT[v["Control visit (AP)"]])
    recs += tests_lines(v, "bullet")
    recs.append("• " + AP_WARNING)
    return {"clinical_diagnosis": _ap_diagnosis(v, True), "epicrisis": epicrisis, "recommendations": "\n".join(recs)}


def n41_checks(v):
    blockers, cautions = [], []
    nmc = v.get("_department") == "NMC"
    chosen = v.get("_blocks") or []
    risks = v.get("Risk factors (AP)") or []
    tests = v.get("Urine tests (AP)") or []
    fq = [b for b in chosen if b in ("cipro-ap-14", "cipro-ap-28", "levofloxacin-ap")]
    if nmc:
        if v.get("Sepsis signs (AP)"):
            blockers.append("Signs of sepsis: treat as a systemic UTI – admit for IV antibiotics (EAU).")
        if _ap_retention(v):
            blockers.append("Urinary retention with acute prostatitis: admit (drainage, IV antibiotics).")
        if _ap_abscess(v) or v.get("DRE (AP)") == "Tender, fluctuating area (abscess?)":
            blockers.append("Prostatic abscess (suspected or proven): admit – imaging and drainage as needed.")
        if any(x in risks for x in ("Diabetes", "Immunosuppression")) or v.get("Context (AP)") == "After prostate biopsy":
            cautions.append("Diabetes, immunosuppression or prostatitis after biopsy: higher risk of severe infection and "
                            "abscess – consider admission.")
    else:
        if _ap_abscess(v) and v.get("Abscess treatment (AP)") == "Antibiotics only":
            size = _num(v, "Abscess size (cm) (AP)")
            if size is None or size >= 2:
                cautions.append("Prostatic abscess of 2 cm or more (or size unknown) treated with antibiotics only: "
                                "usually drained (TRUS-guided aspiration or transurethral deroofing); repeat TRUS.")
    if v.get("DRE (AP)") == "Not performed":
        cautions.append("No DRE: a gentle rectal examination is part of the assessment (prostate tenderness, abscess); "
                        "EAU (strong): no prostatic massage.")
    if "Midstream urine culture" not in tests:
        cautions.append("EAU: take a midstream urine culture to guide treatment.")
    if "Fever" in (v.get("Symptoms (AP)") or []) and "Blood culture" not in tests:
        cautions.append("EAU: take a blood culture and a blood count in acute bacterial prostatitis.")
    if fq and (v.get("Context (AP)") in ("After prostate biopsy", "After catheter / instrumentation")
               or "Fluoroquinolone in the last 6 months" in risks):
        cautions.append("Fluoroquinolone after biopsy / instrumentation / recent fluoroquinolone use: EAU (strong) – "
                        "not empirical (ESBL and quinolone resistance); use it only by culture.")
    if fq:
        cautions.append("Fluoroquinolone: EMA warnings – explain tendon / nerve / aortic risks.")
    if not any(b in chosen for b in _ap_abx):
        cautions.append("Choose an oral antibiotic (total course 2–4 weeks).")
    if "STI risk (young, sexually active)" in risks:
        if "First-void urine NAAT (STI)" not in tests:
            cautions.append("STI risk: first-void urine NAAT for chlamydia / gonorrhoea.")
    if "Recurrent prostatitis" in risks:
        cautions.append("Recurrent prostatitis: after the acute episode look for chronic bacterial prostatitis – "
                        "Meares–Stamey 2- or 4-glass test (EAU, strong); CBP is treated with a fluoroquinolone for "
                        "4–6 weeks.")
    if v.get("Retention drainage (AP)") == "Urethral catheter (small calibre)" and not nmc:
        cautions.append("Urethral catheter in acute prostatitis: keep it short and small-calibre; a suprapubic catheter "
                        "is classically preferred if longer drainage is needed (Campbell-Walsh-Wein).")
    return blockers, cautions


BUILDERS["n41_nmc"] = n41_nmc_parts
BUILDERS["n41_ward"] = n41_ward_parts
CHECKS["N41.0"] = n41_checks
BUILDERS["n45_ward"] = n45_ward_parts
CHECKS["N45"] = n45_checks


# ---------------------------------------------------------------
# 2e. PROSTATE CANCER (C61): checks and text builders
# ---------------------------------------------------------------


def c61_gleason(option, style="dx"):
    """'3+4=7 (ISUP 2)' -> 'Gleason 3+4=7 (ISUP 2)' (diagnosis) or 'Gleason 3+4=7, ISUP 2' (text)."""
    score, isup = option.split(" (ISUP ")
    isup = isup.rstrip(")")
    if style == "dx":
        return f"Gleason {score} (ISUP {isup})"
    return f"Gleason {score}, ISUP {isup}"


def c61_pathology(v):
    """'pT3a pN0 R1' after prostatectomy, '' otherwise."""
    pn = v["Pathological N stage"].split(" ")[0]  # "pNx (no lymph ...)" -> "pNx"
    return f"{v['Pathological T stage']} {pn} {v['Surgical margin']}"


def c61_first_line(v):
    """The first line of the clinical diagnosis."""
    if v.get("Operation") == "Radical prostatectomy" and v.get("Histology") == "Known":
        text = "Prostatas CA " + c61_pathology(v)
        if v.get("Final Gleason score") not in (None, "Not stated"):
            text += ", " + c61_gleason(v["Final Gleason score"])
        return text + "."
    stage = " ".join(v[k] for k in ("Clinical T stage", "Prostate N stage", "Prostate M stage")
                     if v.get(k) not in (None, "Not stated", "Not assessed"))
    text = "Prostatas CA" + (" " + stage if stage else "")
    if v.get("Biopsy Gleason score") not in (None, "Not known"):
        text += ", " + c61_gleason(v["Biopsy Gleason score"])
    return text + "."


def c61_basis(v):
    """' (Gleason 3+4=7, ISUP 2; PSA 8.4 ng/ml)' or ''."""
    parts = []
    if v.get("Biopsy Gleason score") not in (None, "Not known"):
        parts.append(c61_gleason(v["Biopsy Gleason score"], "text"))
    if v.get("PSA (ng/ml)") not in (None, ""):
        parts.append(f"PSA {v['PSA (ng/ml)']:g} ng/ml")
    return f" ({'; '.join(parts)})" if parts else ""


def c61_courses(v):
    """Wording for the operation and the period after it."""
    if v["Intraoperative course"] == "Without complications":
        intraop = "bez komplikācijām"
    else:
        intraop = "ar komplikācijām: " + v["Intraoperative complications"].strip().rstrip(".")
    if v["Postoperative course"] in POSTOP_TEXT:
        postop = POSTOP_TEXT[v["Postoperative course"]]
    else:
        postop = "ar komplikācijām: " + v["Postoperative complications"].strip().rstrip(".")
    return intraop, postop


def c61_catheter_sentence(f, v):
    removal = v["Catheter removal"]
    if removal == "Discharged with catheter":
        sentence = f"{f['Pacients']} {f['izrakstits']} ar urīna ilgkatetru"
        if v.get("Planned catheter removal (optional)"):
            sentence += f", kura evakuācija plānota {fmt_date(v['Planned catheter removal (optional)'])}"
        return sentence + "."
    when = CATHETER_DAY_TEXT.get(removal) or fmt_date(v["Catheter removed on"])
    return (f"Urīna ilgkatetrs evakuēts {when}, pēc kā "
            f"{C67_VOIDING_TEXT[v['Voiding after catheter removal']]}.")


def c61_histology(v):
    if v["Histology"] == "Pending":
        return "Histoloģiskās izmeklēšanas rezultāts izrakstīšanas brīdī vēl nav zināms."
    details = (v.get("Prostate histology details") or "").strip().rstrip(".")
    parts = []
    if v.get("Operation") == "Radical prostatectomy":
        result = c61_pathology(v)
        if v.get("Final Gleason score") not in (None, "Not stated"):
            result += ", " + c61_gleason(v["Final Gleason score"], "text")
        parts.append(result)
    parts.append(details)
    return "Histoloģiskās izmeklēšanas rezultāts: " + ", ".join(p for p in parts if p) + "."


CATHETER_PLACE_TEXT = {"PSKUS outpatient clinic": " P. Stradiņa KUS poliklīnikā",
                       "Family doctor's practice": " ģimenes ārsta praksē", "Not stated": ""}
WOUND_SUTURE_TEXT = {"Absorbable": "Diegu evakuācija nav nepieciešama (intrakutānās šuves).",
                     "Non-absorbable": "Diegu evakuācija 10.–14. pēcoperācijas dienā ģimenes ārsta praksē."}


def c61_common_recs(f, v, blocks, extra_regime, regime=None, wound=(), numbered_blocks=False):
    """Recommendation lines shared by prostatectomy and TURP."""
    recs = []
    if v["Discharge paper"] == "Temporary":
        recs.append(TEMPORARY_PAPER_TEXT)
    recs += [
        "Atrasties ģimenes ārsta un ārstējošā urologa uzraudzībā.",
        "Diēta: vispusīga. Lietot šķidrumu aptuveni 2,0 L diennaktī, ja nav medicīnisku kontrindikāciju.",
        regime or (f"Režīms: saudzējošs {WEEKS_TEXT[v['Activity restriction']]} no operācijas dienas. Atturēties no "
                   "pārmērīgas fiziskas slodzes, smagumu celšanas, pirts un baseina apmeklējuma."),
    ] + extra_regime
    if v["Catheter removal"] == "Discharged with catheter":
        if v.get("Planned catheter removal (optional)"):  # a removal date: no 4–6 weekly change needed
            recs.append("Kopt urīnpūšļa ilgkatetru; " + CATHETER_CARE_LINES[1][0].lower() + CATHETER_CARE_LINES[1][1:])
        else:
            recs += CATHETER_CARE_LINES
        if v.get("Planned catheter removal (optional)"):
            place = CATHETER_PLACE_TEXT.get(v.get("Catheter removal place"), "")
            time = (v.get("Catheter removal time") or "").strip()
            recs.append(f"Katetra evakuācija plānota{place} {fmt_date(v['Planned catheter removal (optional)'])}"
                        + (f" plkst. {time}" if time else "") + ".")
    else:
        recs.append("Regulāri, ik pēc 1–2 stundām, iztukšot urīnpūsli.")
    if wound:
        recs += wound
    if blocks and numbered_blocks:
        recs.append("Turpināt iepriekš nozīmēto medikamentozo terapiju, papildus nozīmē:")
        recs += [f"({number}) {b.rstrip('.')};" for number, b in enumerate(blocks, start=1)]
        recs[-1] = recs[-1][:-1] + "."
    elif blocks:
        recs.append("Terapijā lietot:")
        recs += ["• " + b for b in blocks]
    visit = f"Vizīte pie urologa ambulatori: {WEEKS_TEXT[v['Urology visit']]} pēc operācijas"
    if v["Histology"] == "Pending":
        visit += ", lai izvērtētu histoloģisko atbildi"
    recs.append(visit + ".")
    return recs


def c61_council(f, v):
    council = v["Oncology council"]
    if council == "Will be discussed after histology":
        return [f"{f['pacienta'].capitalize()} gadījums tiks izskatīts onkoloģiskajā konsīlijā pēc "
                "histoloģiskās atbildes saņemšanas."]
    if council == "Already discussed":
        return [f"{f['pacienta'].capitalize()} gadījums izskatīts onkoloģiskajā konsīlijā "
                f"{fmt_date(v['Oncology council date'])}."]
    return []


def c61_rp_parts(sex, v, blocks=()):
    f = SEX_FORMS["Man"]  # prostate cancer: always male wording
    approach_nom, approach_gen = RP_TEXT[v["RP approach"]]
    lnd = v.get("Lymph node dissection")
    operation = approach_nom + (LND_TEXT[0] if lnd else "")
    operation_gen = approach_gen + (LND_TEXT[1] if lnd else "")
    day = fmt_date(v["Procedure date"])
    admission = "akūtā kārtā" if v["Admission"] == "Acute" else "plānveida kārtā"
    intraop, postop = c61_courses(v)

    diagnosis = f"{c61_first_line(v)}\nSt. pēc {operation_gen} ({day})."
    drain = v.get("RP drain") or "No drain"
    drain_course = ""
    if drain.startswith("Removed on day"):
        drain_course = f" Drenu evakuē {drain.split()[-1]}. pēcoperācijas dienā."
    elif drain == "Discharged with drain":
        drain_course = " Izrakstīšanas brīdī drena saglabāta."
    epicrisis = "\n".join([
        f"{f['Pacients']} {f['stacionets']} P. Stradiņa KUS 4. uroloģijas nodaļā {admission} operatīvai "
        f"ārstēšanai sakarā ar histoloģiski verificētu prostatas adenokarcinomu{c61_basis(v)}. Iestājoties sūdzību "
        f"nav, {f['pacients']} tiek {f['gatavots']} operatīvai terapijai.",
        f"{day} vispārējā anestēzijā veikta {operation}" + (", drenāža" if drain != "No drain" else "")
        + f". Operācija noritēja {intraop}.",
        f"Pēcoperācijas periods {postop}.{drain_course} {c61_catheter_sentence(f, v)}",
        f"{c61_histology(v)} Uzlabojoties vispārējam stāvoklim, {f['pacients']} {f['izrakstits']} turpmākai "
        "aprūpei ģimenes ārsta uzraudzībā. Sniegtas rekomendācijas.",
    ])
    rest = WEEKS_TEXT[v["Activity restriction"]]
    if v.get("Sauna and pool") == "3 months":
        regime = (f"Režīms – saudzējošs {rest} no operācijas dienas, atturēties no pārmērīgas fiziskas slodzes un "
                  "smagumu celšanas; pirts vai baseina apmeklējuma – 3 mēnešus no operācijas dienas.")
    else:
        regime = None
    wound = ["Ambulatori turpināt pēcoperācijas brūču apdari, šuvju vietas apstrādājot ar Betadine 100 mg/ml vai "
             "brilijantzaļo šķīdumu, pārsēja maiņa pēc nepieciešamības."]
    if v.get("Skin sutures") in WOUND_SUTURE_TEXT:
        wound.append(WOUND_SUTURE_TEXT[v["Skin sutures"]])
    if drain == "Discharged with drain":
        wound.append("Drenas evakuācija pēc ārstējošā urologa norādījumiem.")
    recs = c61_common_recs(f, v, list(blocks), [], regime=regime, wound=wound, numbered_blocks=True)
    recs.insert(recs.index(next(r for r in recs if r.startswith(("Urīnpūšļa katetra maiņa", "Kopt urīnpūšļa", "Regulāri")))),
                "Nozīmē Kēgela vingrinājumus iegurņa pamatnes muskulatūras nostiprināšanai, sākot pēc katetra "
                "evakuācijas (skat. izsniegtās norādes).")
    recs += tests_lines(v, "plain", after_op=True)  # e.g. PSA 3 months after the operation
    recs += c61_council(f, v)
    recs.append(C61_RP_WARNING)
    return {"clinical_diagnosis": diagnosis, "epicrisis": epicrisis, "recommendations": "\n".join(recs)}


def c61_turp_parts(sex, v, blocks=()):
    f = SEX_FORMS["Man"]
    day = fmt_date(v["Procedure date"])
    admission = "akūtā kārtā" if v["Admission"] == "Acute" else "plānveida kārtā"
    intraop, postop = c61_courses(v)

    diagnosis = f"{c61_first_line(v)}\nSt. pēc TURP ({day})."
    course = []
    if v.get("Bladder irrigation after surgery"):
        course.append("Pēcoperācijas periodā veikta urīnpūšļa skalošana.")
    course.append(f"Pēcoperācijas periods {postop}.")
    course.append(c61_catheter_sentence(f, v))
    epicrisis = "\n".join([
        f"{f['Pacients']} {f['stacionets']} P. Stradiņa KUS 4. uroloģijas nodaļā {admission} operatīvai "
        f"ārstēšanai sakarā ar verificētu prostatas vēzi{c61_basis(v)}. {f['Pacients']} tiek {f['gatavots']} "
        "operatīvai terapijai.",
        f"{day} {ANESTHESIA_TEXT[v['TURP anesthesia']]} veikta transuretrāla prostatas rezekcija (TURP). "
        f"Operācija noritēja {intraop}.",
        " ".join(course),
        f"{c61_histology(v)} Uzlabojoties vispārējam stāvoklim, {f['pacients']} {f['izrakstits']} turpmākai "
        "aprūpei ģimenes ārsta uzraudzībā. Sniegtas rekomendācijas.",
    ])
    recs = c61_common_recs(f, v, list(blocks), [
        "Izvairīties no aizcietējumiem un spiešanās defekācijas laikā.",
        "Atturēties no dzimumdzīves 4 nedēļas no operācijas dienas.",
    ])
    recs += tests_lines(v, "plain", after_op=True)
    recs.append("Turpināt prostatas vēža ārstēšanu un novērošanu pie ārstējošā urologa.")
    recs += c61_council(f, v)
    recs.append("Ja parādās izteikta hematūrija ar asins recekļiem, nav iespējams urinēt, rodas drudzis vai "
                "būtiski pasliktinās vispārējais stāvoklis, nekavējoties vērsties pēc neatliekamās "
                "medicīniskās palīdzības.")
    return {"clinical_diagnosis": diagnosis, "epicrisis": epicrisis, "recommendations": "\n".join(recs)}


def nmc_cancer_parts(f, v, blocks, history, dx_first):
    """NMC visit sent home, shared by prostate and kidney cancer.
    history = what follows 'Anamnēzē ...', dx_first = first line of the diagnosis."""
    reason = v["NMC reason"]
    complaint, reason_dx = NMC_REASON_TEXT[reason]
    if reason == "Pain":
        complaint = "sāpēm " + v["Pain location"].strip()
    sentences = [f"{f['Pacients']} {f['iestajies']} akūtā kārtā PSKUS NMC ar sūdzībām par {complaint}. "
                 f"Anamnēzē {history}."]
    if reason == "Urinary retention":
        text = f"Sterilos apstākļos ievietots urīnpūšļa Foley {v['NMC catheter Fr']} Fr katetrs"
        if v.get("Drained urine (ml)") not in (None, ""):
            text += f", izdalās {v['Drained urine (ml)']:g} ml urīna"
        sentences.append(text + ".")
    elif reason == "Hematuria":
        if v["Hematuria management"] == "Observation":
            sentences.append("NMC etapā novērots, urīns kļūst gaišāks. Hemodinamika stabila.")
        else:
            sentences.append("Ievietots trīsceļu urīnpūšļa katetrs, veikta urīnpūšļa skalošana, urīns kļūst "
                             "gaišāks. Hemodinamika stabila.")
    else:
        sentences.append("NMC etapā veikta analgēzija ar pozitīvu efektu.")
    sentences.append("Paņemtas asins un urīna analīzes. Dotajā brīdī nav indikāciju neatliekamai stacionēšanai "
                     "uroloģijas nodaļā. Vispārēji apmierinošā stāvoklī un ar rekomendācijām izraksta tālākai "
                     "ambulatorai ārstēšanai.")

    numbered = make_counter()
    recs = [
        numbered("Atrasties ģimenes ārsta un ārstējošā urologa uzraudzībā"),
        numbered("Lietot daudz šķidruma, vismaz 1.5-2.5 l/DN"),
    ]
    if v.get("Discharged with catheter"):
        text = ("Kopt urīnpūšļa katetru – tā maiņa ik pēc 4–6 nedēļām ģimenes ārsta praksē. Urīna "
                "savācējmaisa maiņa ik pēc 3–7 dienām, vai biežāk, ja tas ir bojāts.")
        if v.get("Catheter removal date"):
            text += f" Katetra evakuācija plānota {fmt_date(v['Catheter removal date'])}."
        recs.append(numbered(text))
    if blocks:
        recs.append(numbered("Terapijā lietot:"))
        recs += ["    " + b for b in blocks]
    recs.append(numbered(f"Ambulatori urologa konsultācija {CONSULT_TEXT[v['Urology consultation']]}."))
    recs += tests_lines(v, "numbered", numbered)
    recs.append(numbered("Stāvokļa pasliktināšanās gadījumā vērsties pēc palīdzības NMC."))
    return {"clinical_diagnosis": f"{dx_first}\n{reason_dx}",
            "epicrisis": " ".join(sentences), "recommendations": "\n".join(recs)}


def nmc_cancer_checks(v):
    blockers, cautions = [], []
    if v.get("Fever at NMC"):
        blockers.append("Fever: assess for urinary tract infection or urosepsis before sending the patient home.")
    if v.get("Hemodynamically stable") == "No":
        blockers.append("The patient is hemodynamically unstable.")
    if v.get("Urine clearing") == "No":
        blockers.append("The urine is not clearing (risk of clot retention).")
    if v.get("Pain controlled") == "No":
        blockers.append("Pain is not controlled.")
    if v.get("NMC reason") == "Urinary retention" and not v.get("Discharged with catheter"):
        cautions.append("Urinary retention, but no catheter at discharge: make sure the patient voids.")
    return blockers, cautions


def c61_nmc_parts(sex, v, blocks=()):
    f = SEX_FORMS["Man"]
    return nmc_cancer_parts(f, v, blocks, f"prostatas vēzis{c61_basis(v)}", c61_first_line(v))


def c61_checks(v):
    if v.get("_department") == "NMC":
        return nmc_cancer_checks(v)
    blockers, cautions = [], []

    if v.get("Voiding after catheter removal") == "Urinary retention":
        blockers.append("Urinary retention after catheter removal.")
    if v.get("Histology") == "Known" and v.get("Operation") == "TURP" and blank(v.get("Prostate histology details")):
        blockers.append("Describe the TURP histology in 'Prostate histology details'.")
    if v.get("Histology") == "Pending" and v.get("Discharge paper") == "Final":
        cautions.append("Histology is still pending, but the discharge paper is marked as final.")
    if v.get("Histology") == "Known" and v.get("Discharge paper") == "Temporary":
        cautions.append("Histology is already known, but the discharge paper is marked as temporary.")
    if (v.get("Operation") == "Radical prostatectomy" and v.get("Catheter removal") == "Discharged with catheter"
            and not v.get("Planned catheter removal (optional)")):
        cautions.append("The patient goes home with a catheter after prostatectomy: give the planned removal date.")
    adverse = (v.get("Pathological T stage") in ("pT3b", "pT4")
               or v.get("Pathological N stage") == "pN1"
               or v.get("Surgical margin") == "R1"
               or v.get("Prostate N stage") == "cN1"
               or str(v.get("Prostate M stage", "")).startswith("cM1"))
    if adverse and v.get("Oncology council") == "Not applicable":
        cautions.append("Adverse pathology or advanced disease, but no oncology council plan is set.")
    if v.get("Postoperative course") == "With complications":
        cautions.append("Complications are recorded: check that the recommendations cover them.")
    return blockers, cautions


BUILDERS.update({"c61_rp": c61_rp_parts, "c61_turp": c61_turp_parts, "c61_nmc": c61_nmc_parts})
CHECKS["C61"] = c61_checks


# ---------------------------------------------------------------
# 2f. KIDNEY CANCER (C64): checks and text builders
# ---------------------------------------------------------------


def c64_stage(v):
    """'cT1b cN0 cM0' (only what is known)."""
    return " ".join(v[k] for k in ("RCC clinical T stage", "RCC N stage", "RCC M stage")
                    if v.get(k) not in (None, "Not stated", "Not assessed"))


def c64_size(v):
    return f"Ø {v['Tumour size (cm)']:g} cm" if v.get("Tumour size (cm)") not in (None, "") else ""


def c64_clinical_line(v):
    """'Labās nieres jaunveidojums cT1b cN0 cM0 (Ø 4.5 cm).'"""
    text = f"{KIDNEY_SIDE_TEXT[v['Kidney side']].capitalize()} nieres jaunveidojums"
    if c64_stage(v):
        text += " " + c64_stage(v)
    if c64_size(v):
        text += f" ({c64_size(v)})"
    return text + "."


def c64_basis(v):
    """' (Ø 4.5 cm; cT1b cN0 cM0)' or ''."""
    parts = [p for p in (c64_size(v), c64_stage(v)) if p]
    return f" ({'; '.join(parts)})" if parts else ""


def c64_pathology(v):
    """'pT1b pN0 R0, ISUP 2'."""
    text = f"{v['RCC pT stage']} {v['RCC pN stage']} {v['RCC margin']}"
    if v.get("RCC ISUP grade") not in (None, "Not graded"):
        text += f", ISUP {v['RCC ISUP grade']}"
    return text


def c64_subtype(v):
    if v.get("RCC subtype") in ("Other RCC", "Benign"):
        return (v.get("Kidney histology details") or "").strip().rstrip(".")
    return RCC_SUBTYPE_TEXT[v["RCC subtype"]]


def c64_operation(v):
    """(name as done, name after 'St. pēc')."""
    adj, adj_gen = KIDNEY_APPROACH_TEXT[v["Kidney approach"]]
    side = KIDNEY_SIDE_TEXT[v["Kidney side"]]
    if v["Kidney operation"] == "Radical nephrectomy":
        nom, gen = f"{adj} {side} puses radikāla nefrektomija", f"{adj_gen} {side} puses radikālas nefrektomijas"
    else:
        nom, gen = f"{adj} {side} nieres rezekcija", f"{adj_gen} {side} nieres rezekcijas"
    extras = []
    if v.get("Adrenalectomy"):
        extras.append("ipsilaterālu adrenalektomiju")
    if v.get("Lymph node dissection"):
        extras.append("reģionālu limfadenektomiju")
    if extras:
        nom += " ar " + lv_list(extras)
        gen += " ar " + lv_list(extras)
    return nom, gen


def c64_surgery_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    op_nom, op_gen = c64_operation(v)
    day = fmt_date(v["Procedure date"])
    admission = "akūtā kārtā" if v["Admission"] == "Acute" else "plānveida kārtā"
    intraop, postop = c61_courses(v)
    side = KIDNEY_SIDE_TEXT[v["Kidney side"]]

    # Clinical diagnosis
    benign = v["Histology"] == "Known" and v.get("RCC subtype") == "Benign"
    if benign:
        first = f"{side.capitalize()} nieres jaunveidojums ({c64_subtype(v)})."
    elif v["Histology"] == "Known":
        first = f"{side.capitalize()} nieres CA {c64_pathology(v)}"
        if c64_subtype(v):
            first += f" ({c64_subtype(v)})"
        first += "."
    else:
        first = c64_clinical_line(v)
    diagnosis = f"{first}\nSt. pēc {op_gen} ({day})."

    # Epicrisis
    labs = []
    if v.get("Creatinine at discharge (µmol/l)") not in (None, ""):
        labs.append(f"kreatinīns {v['Creatinine at discharge (µmol/l)']:g} µmol/l")
    if v.get("eGFR at discharge") not in (None, ""):
        labs.append(f"GFĀ {v['eGFR at discharge']:g} ml/min/1,73 m²")
    renal = RENAL_FUNCTION_TEXT[v["Renal function at discharge"]] + (f" ({', '.join(labs)})" if labs else "") + "."
    if v["Histology"] == "Pending":
        histology = "Histoloģiskās izmeklēšanas rezultāts izrakstīšanas brīdī vēl nav zināms."
    elif benign:
        histology = f"Histoloģiskās izmeklēšanas rezultāts: {c64_subtype(v)}."
    else:
        parts = [c64_subtype(v), c64_pathology(v)]
        details = (v.get("Kidney histology details") or "").strip().rstrip(".")
        if v.get("RCC subtype") != "Other RCC":
            parts.append(details)
        histology = "Histoloģiskās izmeklēšanas rezultāts: " + ", ".join(p for p in parts if p) + "."
    epicrisis = "\n".join([
        f"{f['Pacients']} {f['stacionets']} P. Stradiņa KUS 4. uroloģijas nodaļā {admission} operatīvai "
        f"ārstēšanai sakarā ar radioloģiski verificētu {side} nieres jaunveidojumu{c64_basis(v)}. "
        f"{f['Pacients']} tiek {f['gatavots']} operatīvai terapijai.",
        f"{day} vispārējā anestēzijā veikta {op_nom}. Operācija noritēja {intraop}.",
        f"Pēcoperācijas periods {postop}. {renal}",
        f"{histology} Uzlabojoties vispārējam stāvoklim, {f['pacients']} {f['izrakstits']} turpmākai aprūpei "
        "ģimenes ārsta uzraudzībā. Sniegtas rekomendācijas.",
    ])

    # Recommendations
    recs = []
    if v["Discharge paper"] == "Temporary":
        recs.append(TEMPORARY_PAPER_TEXT)
    recs += [
        "Atrasties ģimenes ārsta un ārstējošā urologa uzraudzībā.",
        "Diēta: vispusīga. Lietot šķidrumu aptuveni 2,0 L diennaktī, ja nav medicīnisku kontrindikāciju.",
        f"Režīms: saudzējošs {WEEKS_TEXT[v['Activity restriction']]} no operācijas dienas. Atturēties no "
        "pārmērīgas fiziskas slodzes, smagumu celšanas, pirts un baseina apmeklējuma.",
        "Nelietot nefrotoksiskus medikamentus, t.sk. nesteroīdos pretiekaisuma līdzekļus (Ibuprofēns, "
        "Ibumetins, Diklofenaks u.c.).",
    ]
    if blocks:
        recs.append("Terapijā lietot:")
        recs += ["• " + b for b in blocks]
    visit = f"Vizīte pie urologa ambulatori: {WEEKS_TEXT[v['Urology visit']]} pēc operācijas"
    if v["Histology"] == "Pending":
        visit += ", lai izvērtētu histoloģisko atbildi"
    recs.append(visit + ".")
    recs += tests_lines(v, "plain", after_op=True)  # e.g. creatinine, CT by risk group
    if benign:
        recs.append("Turpmākā novērošana pēc ārstējošā urologa norādījumiem.")
    elif v["Histology"] == "Known":
        recs.append("Turpmākā novērošana atbilstoši recidīva riska grupai pēc ārstējošā urologa norādījumiem.")
    else:
        recs.append("Turpmākās novērošanas plānu ārstējošais urologs noteiks pēc histoloģiskās atbildes saņemšanas.")
    recs += c61_council(f, v)
    recs.append(C64_WARNING)
    return {"clinical_diagnosis": diagnosis, "epicrisis": epicrisis, "recommendations": "\n".join(recs)}


def c64_nmc_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    history = f"{KIDNEY_SIDE_TEXT[v['Kidney side']]} nieres jaunveidojums{c64_basis(v)}"
    return nmc_cancer_parts(f, v, blocks, history, c64_clinical_line(v))


def c64_checks(v):
    if v.get("_department") == "NMC":
        return nmc_cancer_checks(v)
    blockers, cautions = [], []
    if v.get("Renal function at discharge") == "Deteriorating":
        blockers.append("Renal function is deteriorating.")
    if (v.get("Histology") == "Known" and v.get("RCC subtype") in ("Other RCC", "Benign")
            and blank(v.get("Kidney histology details"))):
        blockers.append("Describe the histology in 'Kidney histology details'.")
    egfr = v.get("eGFR at discharge")
    if egfr not in (None, "") and egfr < 45:
        cautions.append("eGFR is below 45: EAU suggests nephrology referral during follow-up.")
    if v.get("Histology") == "Pending" and v.get("Discharge paper") == "Final":
        cautions.append("Histology is still pending, but the discharge paper is marked as final.")
    if v.get("Histology") == "Known" and v.get("Discharge paper") == "Temporary":
        cautions.append("Histology is already known, but the discharge paper is marked as temporary.")
    high = (v.get("Follow-up risk group") == "High"
            or v.get("RCC pN stage") == "pN1"
            or v.get("RCC margin") == "R1"
            or v.get("RCC pT stage") in ("pT3a", "pT3b", "pT3c", "pT4")
            or v.get("RCC M stage") == "cM1")
    if high and v.get("Oncology council") == "Not applicable":
        text = "High-risk or advanced disease, but no oncology council plan is set."
        if v.get("RCC subtype") == "Clear cell":
            text += " Adjuvant pembrolizumab may be considered for high-risk clear cell RCC (EAU, weak recommendation)."
        cautions.append(text)
    if v.get("Postoperative course") == "With complications":
        cautions.append("Complications are recorded: check that the recommendations cover them.")
    return blockers, cautions


BUILDERS.update({"c64_surgery": c64_surgery_parts, "c64_nmc": c64_nmc_parts})
CHECKS["C64"] = c64_checks


# ---------------------------------------------------------------
# 2g. BIG UPDATE: N10, N47, T83 builders, N20 acute admission, N39 at NMC
# ---------------------------------------------------------------


def abx_row_checks(v):
    """Checks of the antibiotics table rows (used outside N39)."""
    blockers, cautions = [], []
    for number, row in enumerate(v.get("Antibiotics") or [], start=1):
        name = row.get("Medication") or f"row {number}"
        for field in ("Medication", "Route", "Start date"):
            if blank(row.get(field)):
                blockers.append(f"Antibiotics, row {number}: {field} is missing.")
        if row.get("Continue at discharge"):
            for field in ("Dose", "Frequency", "End date"):
                if blank(row.get(field)):
                    blockers.append(f"Antibiotics, row {number} ({name}) continues at discharge: {field} is missing.")
    return blockers, cautions


def chosen_groups(v):
    """Groups of the recommendation blocks that were chosen for this diagnosis."""
    return {b["group"] for code_blocks in BLOCKS.values() for b in code_blocks if b["id"] in (v.get("_blocks") or [])}


# ----- N10 pyelonephritis -----

def n10_diagnosis(v):
    return f"{COMPLEXITY_TEXT[v['Infection complexity']]}.\n{PYELO_TEXT[v['Side']]}"


def n10_ward_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    side = SIDE_TEXT[v["Side"]]
    if v["Admission"] == "Acute":
        opening = (f"{f['Pacients']} {f['iestajies']} akūtā kārtā P. Stradiņa KUS NMC ar sūdzībām par sāpēm "
                   f"{side} un paaugstinātu ķermeņa temperatūru. NMC etapā {f['pacients']} tiek vispusīgi "
                   f"{f['izmekleta']}, konstatē urīnceļu infekciju, tiek {f['stacionets']} 4. uroloģijas nodaļā.")
    else:
        opening = (f"{f['Pacients']} plānveida kārtā {f['stacionets']} P. Stradiņa KUS 4. uroloģijas nodaļā "
                   f"pēc nosūtījuma sakarā ar sāpēm {side} un paaugstinātu ķermeņa temperatūru; ambulatori "
                   f"konstatēta urīnceļu infekcija. {f['Pacients']} tiek vispusīgi {f['izmekleta']}.")
    epicrisis = (f"{opening} Uzsākta konservatīva un antibakteriālā terapija. Uz saņemtās terapijas fona "
                 f"{f['pacienta']} stāvoklis ar pozitīvu dinamiku gan klīniski, gan laboratoriski, iekaisuma "
                 f"rādītāji mazinās. Uzlabojoties vispārējam stāvoklim {f['pacients']} tiek {f['izrakstits']} "
                 "turpmākai aprūpei ģimenes ārsta uzraudzībā, sniegtas rekomendācijas.")
    recs = "\n".join([
        "Atrasties ģimenes ārsta uzraudzībā.",
        "• Diēta – vispusīga, lietot šķidrumu 2.0 – 2.5 L/dnn, atturēties no saldinātu, "
        "kofeīnu vai alkoholu saturošu dzērienu lietošanas.",
        "• Režīms – saudzējošs, izvairīties no pārmērīgas fiziskas slodzes, pirts vai baseina apmeklējuma.",
        "• Regulāri, vismaz reizi 1 – 2 stundās iztukšot urīnpūsli.",
        "{blocks}",
        "• Kontrole pie urologa ambulatori pēc 3 mēnešiem.",
    ] + tests_lines(v, "bullet") + [  # the tests ticked in the Follow-up tests box
        "• Stāvokļa pasliktināšanās gadījumā griezties NMC.",
    ])
    return {"clinical_diagnosis": n10_diagnosis(v), "epicrisis": epicrisis,
            "recommendations": add_blocks(recs, list(blocks))}


def n10_nmc_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    iv = "NMC etapā ievadīta pirmā antibakteriālā preparāta deva i/v. " if v.get("Initial IV antibiotic at NMC") else ""
    epicrisis = (
        f"{f['Pacients']} {f['iestajies']} akūtā kārtā PSKUS NMC ar sūdzībām par sāpēm {SIDE_TEXT[v['Side']]} "
        f"un paaugstinātu ķermeņa temperatūru. NMC etapā vispusīgi {f['izmekleta']}: urīna analīzē iekaisuma "
        "izmaiņas, paņemts urīna uzsējums ar antibakteriālās jutības noteikšanu. "
        f"{N10_IMAGING_TEXT[v['NMC imaging']]}{iv}Dotajā brīdī nav indikāciju neatliekamai stacionēšanai "
        "uroloģijas nodaļā. Uzsākta perorāla antibakteriālā terapija. Vispārēji apmierinošā stāvoklī un ar "
        "rekomendācijām izraksta tālākai ambulatorai ārstēšanai ģimenes ārsta uzraudzībā."
    )
    numbered = make_counter()
    recs = [
        numbered("Atrasties ģimenes ārsta uzraudzībā"),
        numbered("Lietot šķidrumu 2.0 – 2.5 L/dnn"),
        numbered("Režīms – saudzējošs, izvairīties no pārmērīgas fiziskas slodzes, pirts vai baseina apmeklējuma."),
    ]
    if blocks:
        recs.append(numbered("Terapijā lietot:"))
        recs += ["    " + b for b in blocks]
    recs += [
        numbered("Urīna uzsējuma rezultātus var aplūkot Datamed.lv pēc 3–5 dienām; antibakteriālo terapiju "
                 "koriģēt atbilstoši jutībai."),
        numbered("Ja 48–72 stundu laikā stāvoklis neuzlabojas vai saglabājas paaugstināta ķermeņa temperatūra, "
                 "nekavējoties vērsties NMC."),
        numbered(f"Ambulatori urologa konsultācija {CONSULT_TEXT[v['Urology consultation']]}."),
    ]
    recs += tests_lines(v, "numbered", numbered)
    recs.append(numbered("Stāvokļa pasliktināšanās gadījumā vērsties pēc palīdzības NMC."))
    return {"clinical_diagnosis": n10_diagnosis(v), "epicrisis": epicrisis, "recommendations": "\n".join(recs)}


def n10_checks(v):
    blockers, cautions = [], []
    if v.get("_department") != "NMC":
        return blockers, cautions
    if v.get("Urostasis on imaging") == "Yes":
        blockers.append("Urostasis with pyelonephritis: an obstructed infected kidney needs admission and drainage.")
    if v.get("Sepsis signs") == "Yes":
        blockers.append("Signs of sepsis: admission is needed.")
    if v.get("Able to take oral medication") == "No":
        blockers.append("The patient cannot take oral medication: admission for IV treatment is needed.")
    if v.get("Pregnant") == "Yes":
        blockers.append("Pyelonephritis in pregnancy: admission is generally needed.")
    if v.get("Infection complexity") == "Complicated":
        cautions.append("Complicated pyelonephritis is usually treated in hospital: confirm that outpatient "
                        "treatment is appropriate.")
    if v.get("NMC imaging") == "None":
        cautions.append("No imaging was done: urinary tract obstruction has not been excluded.")
    groups = chosen_groups(v)
    if G_ORAL not in groups:
        cautions.append("No oral antibiotic is chosen in the recommendation blocks.")
    if G_IV in groups:
        cautions.append("An IV antibiotic is chosen for a patient sent home.")
    return blockers, cautions


# ----- N47 phimosis -----

def n47_parts(sex, v, blocks=()):
    f = SEX_FORMS["Man"]  # phimosis: always male wording
    degree = f"{DEGREE_TEXT[v['Phimosis degree']]}." + (" Frenulum breve." if v.get("Frenulum breve") else "")
    lines = [BALANOPOSTHITIS_TEXT[bool(v.get("Balanoposthitis"))], degree]

    if v.get("Surgery"):
        op, op_gen = SURGERY_TEXT[v["Surgery type"]]["nom"], SURGERY_TEXT[v["Surgery type"]]["gen"]
        if v.get("Frenuloplasty"):
            op, op_gen = op + " un frenuloplastika", op_gen + " un frenuloplastikas"
        diagnosis = "\n".join(lines) + f"\n\nSt. pēc {op_gen}"
        epicrisis = (
            f"{f['Pacients']} {f['stacionets']} P. Stradiņa KUS plānveida kārtā operatīvai terapijai sakarā ar "
            f"sūdzībām par grūtībām atvilkt priekšādiņu. {f['Pacients']} tiek {f['gatavots']} operatīvai terapijai. "
            f"{fmt_date(v['Operation date'])} {ANESTHESIA_N47_TEXT[v['Anesthesia']]} tiek veikta {op}. "
            "Pēcoperācijas periods apmierinošs. Uzlabojoties vispārējam stāvoklim, "
            f"{f['pacients']} tiek {f['izrakstits']} turpmākai ambulatorai aprūpei ģimenes ārsta uzraudzībā, "
            "sniegtas rekomendācijas."
        )
        recs = [
            "• Atrasties ģimenes ārsta uzraudzībā.",
            "• Diēta – vispusīga, lietot šķidrumu 2.0 – 2.5 L/dnn.",
            "• Režīms – saudzējošs 2 nedēļas no operācijas dienas, izvairīties no pārmērīgas fiziskas slodzes, "
            "pirts vai baseina apmeklējuma.",
            "• Atturēties no dzimumdzīves 3 nedēļas no operācijas dienas.",
        ]
        if blocks:
            recs.append("• Turpināt lietot iepriekš nozīmēto terapiju, papildus nozīmē:")
            recs += ["    o " + b for b in blocks]
        else:
            recs.append("• Turpināt lietot iepriekš nozīmēto terapiju.")
        recs += [
            "• Ambulatori turpināt pēcoperācijas brūces apdari, apmazgājot ar viegli sārtas krāsas kālija "
            "permanganāta šķīdumu un šuves vietas apstrādājot ar Betadine 100 mg/ml vai brilijantzaļo šķīdumu.",
            "• Ambulatori turpināt pēcoperācijas brūces apdari un pārsēja maiņu pēc nepieciešamības, diegu "
            "evakuācija 14. - 21. pēcoperācijas dienā ambulatori (vizīti atsevišķi pieteikt pa tālr. 67069280).",
            "• Kontrole pie urologa ambulatori 3 mēnešu pēc operācijas.",
        ]
        recs += tests_lines(v, "bullet", after_op=True)
        recs.append("• Stāvokļa pasliktināšanās gadījumā griezties NMC.")
    else:
        if v["Conservative reason"] == "Balanoposthitis":
            lines.append("Akūts balanopostīts.")
            epicrisis = (
                f"{f['Pacients']} {f['stacionets']} P. Stradiņa KUS akūtā kārtā sakarā ar priekšādiņas apsārtumu, "
                "tūsku un sāpēm. Konstatēts akūts balanopostīts uz fimozes fona. Uzsākta konservatīva terapija – "
                "lokāla antiseptiska apstrāde un antibakteriālā terapija. Uz terapijas fona iekaisuma pazīmes "
                f"mazinās. Uzlabojoties vispārējam stāvoklim, {f['pacients']} tiek {f['izrakstits']} turpmākai "
                "ambulatorai aprūpei ģimenes ārsta uzraudzībā, sniegtas rekomendācijas. Pēc iekaisuma "
                "mazināšanās ieteicama plānveida operatīvā ārstēšana."
            )
        else:
            lines.append("Parafimoze.")
            epicrisis = (
                f"{f['Pacients']} {f['stacionets']} P. Stradiņa KUS akūtā kārtā sakarā ar parafimozi – priekšādiņas "
                "iespīlējumu aiz dzimumlocekļa galviņas ar tūsku. Veikta parafimozes manuāla repozīcija. Pēc "
                f"manipulācijas tūska mazinās. Uzlabojoties vispārējam stāvoklim, {f['pacients']} tiek "
                f"{f['izrakstits']} turpmākai ambulatorai aprūpei ģimenes ārsta uzraudzībā, sniegtas rekomendācijas. "
                "Ieteicama plānveida operatīvā ārstēšana."
            )
        diagnosis = "\n".join(lines)
        recs = [
            "• Atrasties ģimenes ārsta uzraudzībā.",
            "• Ikdienas dzimumorgānu higiēna; priekšādiņu ar spēku neatvilkt.",
            "• Siltas vanniņas ar viegli sārtas krāsas kālija permanganāta šķīdumu 2 reizes dienā līdz "
            "iekaisuma mazināšanās.",
        ]
        if blocks:
            recs.append("• Terapijā lietot:")
            recs += ["    o " + b for b in blocks]
        recs += [
            "• Konsultācija pie urologa ambulatori par plānveida operatīvo ārstēšanu (cirkumcīziju) pēc "
            "iekaisuma mazināšanās.",
        ]
        recs += tests_lines(v, "bullet")
        recs += [
            "• Pastiprinātas tūskas, sāpju, urinācijas traucējumu, paaugstinātas ķermeņa temperatūras vai "
            "stāvokļa pasliktināšanās gadījumā griezties NMC.",
        ]
    return {"clinical_diagnosis": diagnosis, "epicrisis": epicrisis, "recommendations": "\n".join(recs)}


def n47_checks(v):
    blockers, cautions = [], []
    chosen = v.get("_blocks") or []
    if v.get("Surgery") and "amoxiclav-14" in chosen and not v.get("Balanoposthitis"):
        cautions.append("Routine antibiotics are generally not needed after clean genital surgery; Amoksiklav is "
                        "chosen without balanoposthitis.")
    if v.get("Surgery") and v.get("Anesthesia") == "Local" and "fraxiparine" in chosen:
        cautions.append("Fraxiparine is chosen after a short procedure under local anesthesia: check the indication.")
    return blockers, cautions


# ----- T83 catheter dysfunction -----


# ----- N47 paraphimosis at NMC -----
# Paraphimosis is a urological emergency: venous and lymphatic congestion can progress to arterial compromise and
# glans necrosis. Reduce early: analgesia (penile block or lidocaine gel), manual compression of the oedema (ice,
# compression bandage, multiple punctures – Dundee technique) and manual reduction; if that fails, dorsal slit
# (dorsumcision) under local anaesthesia. Elective circumcision is usually advised later (recurrence, cosmesis
# after a dorsal slit). Iatrogenic cause: foreskin left retracted after catheterisation or examination.

PF_DURATION = {"A few hours": "dažas stundas", "1 day": "1 dienu", "2 days": "2 dienas",
               "More than 2 days": "ilgāk par 2 dienām"}
PF_CAUSE = {"After intercourse / masturbation": "pēc dzimumakta", "After urinary catheterisation":
            "pēc urīnpūšļa katetrizācijas", "After examination / hygiene": "pēc apskates vai higiēnas procedūrām"}
PF_GLANS = {"Pink, viable": "dzimumlocekļa galviņa sārta, vitāla",
            "Dusky / cyanotic": "dzimumlocekļa galviņa cianotiska",
            "Necrotic areas": "dzimumlocekļa galviņā nekrozes perēkļi"}
PF_BEFORE = {"Local anaesthesia": "lokālā anestēzijā", "Manual compression of the oedema": "pēc tūskas manuālas "
             "kompresijas", "Puncture of the oedema (Dundee)": "pēc tūskas daudzkārtējas punkcijas ar adatu"}
PF_SUTURE = {"4-0 Vicryl interrupted": "4.0 Vicryl atsevišķām šuvēm",
             "4-0 Vicryl Rapide interrupted": "4.0 Vicryl Rapide atsevišķām šuvēm"}
PF_VISIT = {"4. Urology department": "PSKUS 4. nodaļā", "Urologist outpatient": "pie urologa ambulatori",
            "Family doctor": "pie ģimenes ārsta"}


def n47_nmc_parts(sex, v, blocks=()):
    f = SEX_FORMS["Man"]
    slit = v["Paraphimosis treatment"].startswith("Dorsumcision")
    before = [PF_BEFORE[x] for x in (v.get("Before reduction (PF)") or []) if x in PF_BEFORE]
    day = fmt_date(v["Procedure date"])
    lines = ["Parafimoze." + (" Balanopostīts." if v.get("Balanoposthitis") else "")]
    if slit:
        lines.append(f"Operācijas/invazīvās manipulācijas: Dorsumcīzija lokālā anestēzijā ({day}).")
    else:
        lines.append(f"Operācijas/invazīvās manipulācijas: Parafimozes manuāla repozīcija ({day}).")

    s = [f"{f['Pacients']} vēršas P. Stradiņa KUS NMC ar sūdzībām par tūsku dzimumloceklī "
         f"{PF_DURATION[v['Paraphimosis duration']]}" +
         (f", radusies {lv_list([PF_CAUSE[x] for x in v['Paraphimosis cause'] if x in PF_CAUSE])}"
          if any(x in PF_CAUSE for x in (v.get("Paraphimosis cause") or [])) else "") + "."]
    s.append(f"Apskates laikā konstatē parafimozi; {PF_GLANS[v['Glans appearance']]}.")
    if slit:
        s.append("Manuāla repozīcija neveiksmīga. Lokālā anestēzijā veic dorsumcīziju" +
                 (f", brūces malas apšuj ar {PF_SUTURE[v['Dorsumcision sutures']]}"
                  if v.get("Dorsumcision sutures") in PF_SUTURE else "") + ".")
    else:
        how = lv_list([b for b in before if b != "lokālā anestēzijā"])
        s.append(("Lokālā anestēzijā " if "lokālā anestēzijā" in before else "") +
                 (f"{how} " if how else "") + "veic parafimozes manuālu repozīciju – veiksmīgi; tūska mazinās.")
        s[-1] = s[-1][0].upper() + s[-1][1:]
    s.append("Apmierinošā stāvoklī ar rekomendācijām izrakstīts ambulatori.")

    recs = []
    if blocks:
        recs += ["• Lietot " + b[0].lower() + b[1:] if not b.startswith(("Tab.", "Caps.", "Ung.", "Inj.", "Sol."))
                 else "• Lietot " + b for b in blocks]
    if slit:
        recs.append("• Brūces apstrāde ar Sol. Braunoli un pārsiešana pēc nepieciešamības ar steriliem pārsējiem.")
        recs.append("• Atturēties no dzimumdzīves līdz brūces sadzīšanai (ap 3 nedēļām).")
    else:
        recs.append("• Priekšādiņu neatvilkt aiz dzimumlocekļa galviņas, līdz tūska pilnībā izzūd; pēc higiēnas, "
                    "katetrizācijas vai dzimumakta priekšādiņu vienmēr atgriezt vietā.")
        recs.append("• Ikdienas dzimumorgānu higiēna; pirmās dienas dzimumlocekli turēt paceltu tūskas mazināšanai.")
        recs.append("• Atturēties no dzimumdzīves, līdz tūska pilnībā izzudusi.")
    date = v.get("Control visit date (PF)")
    time = (v.get("Control visit time (PF)") or "").strip()
    place = PF_VISIT[v["Control visit (PF)"]]
    what = "kontroles vizīti" + (", šuvju izņemšanu" if slit and v.get("Suture removal at the visit") and
                                  v.get("Dorsumcision sutures") != "No sutures" else "")
    if date:
        recs.append(f"• Ierasties {fmt_date(date)}" + (f" plkst. {time}" if time else "") + f" {place} uz {what}.")
    else:
        recs.append(f"• Kontroles vizīte {place} pēc 10–14 dienām" +
                    (" (šuvju izņemšana)" if "šuvju" in what else "") + ".")
    if v.get("Plan circumcision"):
        recs.append("• Eventuāli plānojama cirkumcīzija – konsultācija pie urologa." if slit else
                    "• Atkārtotas parafimozes profilaksei ieteicama plānveida cirkumcīzija – konsultācija pie urologa.")
    recs += tests_lines(v, "bullet")
    recs.append("• Ja dzimumlocekļa galviņa kļūst tumša vai zilgana, " +
                ("parafimoze atkārtojas, " if not slit else "") + "pastiprinās tūska vai sāpes, nav iespējams "
                "urinēt, " + ("brūce asiņo " if slit else "") + ("vai " if slit else "") +
                "paaugstinās ķermeņa temperatūra, nekavējoties vērsties NMC.")
    return {"clinical_diagnosis": "\n".join(lines), "epicrisis": " ".join(s), "recommendations": "\n".join(recs)}


def n47_nmc_checks(v):
    blockers, cautions = [], []
    chosen = v.get("_blocks") or []
    if v.get("Glans appearance") == "Necrotic areas":
        blockers.append("Necrosis of the glans: urgent urology care and admission – do not send home.")
    if v.get("Glans appearance") == "Dusky / cyanotic":
        cautions.append("Dusky glans: arterial compromise is possible – check that the colour returned after the "
                        "reduction or slit before discharge.")
    if "amoxiclav-7-pf" in chosen and not v.get("Balanoposthitis"):
        cautions.append("Amoksiklav without balanoposthitis or skin breakdown: routine antibiotics are not needed "
                        "after reduction or a clean dorsal slit.")
    if "After urinary catheterisation" in (v.get("Paraphimosis cause") or []):
        cautions.append("Iatrogenic paraphimosis: remind the nurses / carers to return the foreskin after every "
                        "catheter care.")
    if v.get("Paraphimosis duration") == "More than 2 days" and v.get("Paraphimosis treatment") == "Manual reduction":
        cautions.append("Long-standing paraphimosis: oedema and fibrosis make reduction harder – check the glans and "
                        "consider a dorsal slit early if the reduction is incomplete.")
    return blockers, cautions


def n47_all_checks(v):
    return n47_nmc_checks(v) if v.get("_department") == "NMC" else n47_checks(v)


BUILDERS["n47_nmc"] = n47_nmc_parts

def t83_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    cystostomy = v["Catheter type"] == "Cystostomy"
    name = "epicistostomas katetra" if cystostomy else "urīnpūšļa ilgkatetra"
    diagnosis = (f"{REASON_TEXT[v['Catheter reason']]}{'Epicistostomas katetra' if cystostomy else 'Urīnpūšļa ilgkatetra'} "
                 "disfunkcija." + (" Urīna retence." if v.get("Urinary retention (diagnosis)") else ""))
    change = "epicistostomas katetra maiņu" if cystostomy else "urīnpūšļa katetra maiņu"
    epicrisis = (
        f"{f['Pacients']} {f['iestajies']} akūtā kārtā PSKUS NMC ar sūdzībām par {name} disfunkciju. "
        f"Sterilos apstākļos veic {change}. Ievieto Foley {v['Catheter size (Fr)']} Fr katetru, izdalās ap "
        f"{v['Drained volume (ml)']:g} ml {APPEARANCE_AT_CHANGE_TEXT[v['Urine appearance at change']]}. Tiek "
        "paņemtas asins un urīna analīzes, urīna uzsējums. Dotajā brīdī nav indikāciju neatliekamai "
        "stacionēšanai uroloģijas nodaļā. Vispārēji apmierinošā stāvoklī un ar rekomendācijām izraksta tālākai "
        "ambulatorai ārstēšanai."
    )
    if cystostomy:
        care = ("Kopt epicistostomas katetru un stomas vietu (ikdienas apstrāde un pārsiešana) – katetra maiņa ik "
                "pēc 4–6 nedēļām ģimenes ārsta praksē. Urīna savācējmaisa maiņa ik pēc 3–7 dienām, vai biežāk, ja "
                "tas ir bojāts.")
    else:
        care = ("Kopt urīnpūšļa katetru – tā maiņa ik pēc 4–6 nedēļām ģimenes ārsta praksē. Urīna savācējmaisa "
                "maiņa ik pēc 3–7 dienām, vai biežāk, ja tas ir bojāts.")
    numbered = make_counter()
    recs = [
        numbered("Atrasties ģimenes ārsta un ārstējošā urologa uzraudzībā"),
        numbered("Lietot daudz šķidruma, vismaz 1.5-2.5 l/DN"),
        numbered(care),
    ]
    if blocks:
        recs.append(numbered("Terapijā lietot:"))
        recs += ["    " + b for b in blocks]
    recs += [
        numbered("Laboratorisko izmeklējumu rezultātus var aplūkot Datamed.lv. Urīna uzsējuma rezultāti pēc "
                 "3-5. dienām"),
    ]
    recs += tests_lines(v, "numbered", numbered)
    recs.append(numbered("Stāvokļa pasliktināšanās gadījumā vērsties pēc palīdzības NMC."))
    return {"clinical_diagnosis": diagnosis, "epicrisis": epicrisis, "recommendations": "\n".join(recs)}


def t83_checks(v):
    blockers, cautions = [], []
    if v.get("Fever at NMC"):
        blockers.append("Fever with a urinary catheter: assess for catheter-associated infection or urosepsis.")
    if v.get("Urine appearance at change") == "With clots":
        cautions.append("Clots in the urine: check for clot retention and the source of bleeding.")
    if v.get("Drained volume (ml)") == 0:
        cautions.append("No urine drained after the change: check the catheter position.")
    return blockers, cautions


# ----- N20 acute admission with drainage -----

def n20_acute_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    side = v["Side"]
    dx_line, stone_epi, pain = n20_stone_words(v)
    drain_nom, drain_gen, drain_with = DRAIN_TEXT[v["Drainage"]][side]
    day = fmt_date(v["Drainage date"])

    diagnosis = (f"Nierakmeņu slimība. {dx_line}. Obstruktīvs pielonefrīts."
                 + (" Urosepse." if v.get("Urosepsis") else "") + f"\nSt. pēc {drain_gen} ({day}).")

    cro = f" (CRO {v['CRO (mg/L)']:g} mg/l)" if v.get("CRO (mg/L)") not in (None, "") else ""
    paragraphs = [
        f"{f['Pacients']} {f['iestajies']} akūtā kārtā PSKUS NMC ar sūdzībām par sāpēm {pain} un "
        f"paaugstinātu ķermeņa temperatūru. NMC etapā vispusīgi {f['izmekleta']}; veic CT vēdera dobumam, kur "
        f"vizualizē {stone_epi} {n20_urostasis(v)}. Laboratoriski iekaisuma izmaiņas{cro}."
    ]
    groups = abx_groups(abx_rows(v))
    paragraphs.append(
        f"{day} neatliekamā kārtā veikta {drain_nom}"
        + (", izdalās strutains urīns" if v.get("Purulent urine on drainage") else "") + ". "
        f"{f['Pacients']} {f['stacionets']} 4. uroloģijas nodaļā, uzsākta antibakteriālā terapija ar "
        + lv_list([abx_phrase(r, False) for r in groups[0][1]]) + "."
    )
    for change_day, rows in groups[1:]:
        paragraphs.append(f"{fmt_date(change_day)}. antibakteriālā terapija mainīta uz "
                          + lv_list([abx_phrase(r, True) for r in rows]) + ".")
    status = "uzlabojies" if v["Clinical status"] == "Improved" else "stabils"
    culture = CULTURE_TEXT[v["Urine culture"]]
    if v["Urine culture"] == "Positive" and not blank(v.get("Culture organism")):
        culture += f" ({v['Culture organism'].strip()})"
    course = (f"Uz terapijas fona {f['pacienta']} stāvoklis {status}, "
              f"{TEMPERATURE_COURSE_TEXT[v['Temperature trend']]}, iekaisuma rādītāju dinamika: "
              f"{INFLAMMATORY_TEXT[v['Inflammatory markers trend']]}. Urīna uzsējuma rezultāts: {culture}.")
    if v["Urine culture"] == "Result not yet known":
        course += " Nepieciešams nodrošināt rezultāta pārbaudi un turpmāko rīcību atbilstoši atradei."
    paragraphs.append(course)
    plan = v["Definitive stone treatment"]
    if plan in DEFINITIVE_TEXT:
        when = fmt_date(v["Definitive treatment date"]) if v.get("Definitive treatment date") else "pēc infekcijas izārstēšanas"
        definitive = f"Konkrementa definitīvā ārstēšana – {DEFINITIVE_TEXT[plan]} – plānota {when}."
    else:
        definitive = "Konkrementa definitīvās ārstēšanas taktiku noteiks ārstējošais urologs."
    paragraphs.append(f"{f['Pacients']} {f['izrakstits']} {drain_with} turpmākai ambulatorai ārstēšanai. {definitive}")

    recs = [
        "• Atrasties ģimenes ārsta un ārstējošā urologa uzraudzībā.",
        "• Lietot daudz šķidruma, vismaz 2–2.5 l/DN, ja nav medicīnisku kontrindikāciju.",
        "• Režīms – saudzējošs, izvairīties no pārmērīgas fiziskas slodzes.",
    ]
    antibiotics = abx_discharge_lines(v)
    if antibiotics:
        recs.append("• Antibakteriālā terapija:")
        recs += antibiotics
        recs.append("    o Terapiju precizēt atbilstoši urīna uzsējuma rezultātiem.")
    recs += ["• " + b for b in blocks]
    if v["Drainage"] == "DJ stent":
        recs += [
            "• DJ stents: iespējama bieža urinācija, neatliekama vēlme urinēt, diskomforts vēdera lejasdaļā vai "
            "sānos un neliels asiņu piejaukums urīnā.",
            "• DJ stenta izņemšana vai maiņa kopā ar konkrementa definitīvo ārstēšanu; neatlikt plānoto vizīti.",
        ]
    else:
        recs += [
            "• Nefrostomas aprūpe: izejas vietu apstrādāt un pārsiet ik pēc 2–3 dienām, sekot urīna attecei; "
            "urīna maisa maiņa ik pēc 3–7 dienām.",
            "• Nefrostomu nesaliekt un neizraut; atteces traucējumu gadījumā vērsties NMC.",
        ]
    if plan in DEFINITIVE_TEXT:
        recs.append(f"• Konkrementa definitīvā ārstēšana – {DEFINITIVE_TEXT[plan]} – plānota "
                    + (fmt_date(v["Definitive treatment date"]) if v.get("Definitive treatment date")
                       else "pēc ārstējošā urologa norādījumiem") + ".")
    else:
        recs.append("• Konkrementa definitīvās ārstēšanas taktiku noteiks ārstējošais urologs ambulatori.")
    if v["Urine culture"] == "Result not yet known":
        recs.append(f"• Urīna uzsējuma rezultātu pārbauda {CULTURE_CHECKER_TEXT[v['Culture checked by']]}, "
                    "nepieciešamības gadījumā koriģējot antibakteriālo terapiju.")
    recs += tests_lines(v, "bullet")
    recs.append("• Paaugstinātas ķermeņa temperatūras, drebuļu, stipru sāpju sānos, drenāžas vai urīna atteces "
                "traucējumu vai vispārējā stāvokļa pasliktināšanās gadījumā nekavējoties vērsties NMC.")
    return {"clinical_diagnosis": diagnosis, "epicrisis": "\n".join(paragraphs), "recommendations": "\n".join(recs)}


# ----- N39 at NMC (sent home) -----

def n39_nmc_parts(sex, v, blocks=()):
    return {
        "clinical_diagnosis": n39_diagnosis(v),
        "epicrisis": n39_nmc_epicrisis(sex, v),
        "recommendations": add_blocks(n39_recommendations(v, nmc=True), list(blocks)),
    }


BUILDERS.update({"n10_ward": n10_ward_parts, "n10_nmc": n10_nmc_parts, "n47": n47_parts, "t83": t83_parts,
                 "n20_acute": n20_acute_parts, "n39_nmc": n39_nmc_parts})
CHECKS.update({"N10": n10_checks, "N47": n47_all_checks, "T83": t83_checks})


# ---------------------------------------------------------------
# 2h. URETHRAL STRICTURE (N35): checks and text builder (EAU guideline based)
# ---------------------------------------------------------------


def n35_parts(sex, v, blocks=()):
    f = SEX_FORMS["Man"]
    op = v["Stricture operation"]
    op_nom, op_gen = STRICTURE_OP[op]
    if op == "Internal urethrotomy (DVIU)" and v.get("Urethrotomy instrument") == "Laser":
        op_nom, op_gen = "lāzeruretrotomija", "lāzeruretrotomijas"
    elif op == "Internal urethrotomy (DVIU)":
        op_nom, op_gen = "optiskā uretrotomija (pēc Sachse)", "optiskās uretrotomijas (pēc Sachse)"
    also = (v.get("Also performed (N35)") or "").strip().rstrip(".")
    day = fmt_date(v["Procedure date"])
    length = v.get("Stricture length (cm)")
    catheter_in = v["Catheter at discharge (N35)"] == "In place"
    comp = COMPENSATION.get(v.get("Stricture compensation"))

    # Clinical diagnosis (department format: "Meatus striktūra, st. pēc meatotomijas, urīna ilgkatetrs (date).")
    first = STRICTURE_DX[v["Stricture location"]]
    if comp:
        first = first.replace("striktūra", f"{comp[0]} striktūra")
    if v.get("Previous stricture treatment"):
        first = "Recidivējoša " + first[0].lower() + first[1:]
    if length not in (None, ""):
        first += f" ({length:g} cm)"
    diagnosis = (f"{first}, st. pēc {op_gen}" + (f" un {also}" if also else "")
                 + (", urīna ilgkatetrs" if catheter_in else "") + f" ({day}).")
    cause = v.get("Stricture cause")
    if cause in STRICTURE_CAUSE_DX:
        diagnosis += "\n" + STRICTURE_CAUSE_DX[cause]
    if not blank(v.get("Other diagnoses (optional)")):
        diagnosis += "\n" + v["Other diagnoses (optional)"].strip().rstrip(".") + "."

    # Epicrisis
    reason = STRICTURE_EPI[v["Stricture location"]]
    if comp:
        reason = reason.replace("striktūru", f"{comp[1]} striktūru")
    verified = [VERIFIED_BY[x] for x in (v.get("Stricture verified by") or []) if x in VERIFIED_BY]
    if verified:
        reason = f"{lv_list(verified)} verificētu {reason}"
    if length not in (None, ""):
        reason += f" (garums ap {length:g} cm)"
    if v.get("Previous stricture treatment"):
        reason += ", iepriekš ārstētu"
    admission = "akūtā kārtā" if v["Admission"] == "Acute" else "plānveida kārtā"
    if v["Postoperative course"] in N35_POSTOP:
        postop = N35_POSTOP[v["Postoperative course"]]
    else:
        postop = "ar komplikācijām: " + v["Postoperative complications"].strip().rstrip(".")
    sentences = [
        f"{f['Pacients']} {f['stacionets']} P. Stradiņa KUS 4. uroloģijas nodaļā {admission} sakarā ar {reason}.",
        f"{f['Pacients']} tiek {f['gatavots']} {'endoskopiskai manipulācijai' if op in ENDOSCOPIC else 'operatīvai terapijai'}"
        f", {day} {N35_ANESTHESIA[v['N35 anesthesia']]} tiek veikta {op_nom}" + (f" un {also}" if also else "")
        + f", tiek ievadīts Fr{v['Catheter Fr (N35)']} izmēra urīna ilgkatetrs.",
        f"Pēcoperācijas periods {postop}.",
    ]
    if v["Tissue for histology"] == "Pending":
        sentences.append("Audu materiāls nosūtīts histoloģiskai izmeklēšanai, rezultāts izrakstīšanas brīdī vēl nav zināms.")
    elif v["Tissue for histology"] == "Known":
        sentences.append(f"Histoloģiskās izmeklēšanas rezultāts: {v['Histology result (N35)'].strip().rstrip('.')}.")
    if not catheter_in:
        sentences.append(f"Urīna ilgkatetrs evakuēts, pēc kā {N35_VOIDING[v['Voiding after removal (N35)']]}.")
    sentences.append(f"Uzlabojoties vispārējam stāvoklim {f['pacients']} tiek {f['izrakstits']} turpmākai aprūpei "
                     "ģimenes ārsta uzraudzībā" + (" ar funkcionējošu urīna ilgkatetru" if catheter_in else "")
                     + ", sniegtas rekomendācijas.")
    investigations = (v.get("Investigations (N35)") or "").strip().rstrip(".") or "veikti ambulatorā kārtā, pieļauj operāciju"
    epicrisis = (" ".join(sentences) + f"\n\nIzmeklējumi: {investigations}.\n\nOperācijas: {op_nom}"
                 + (f", {also}" if also else "") + f" ({day}).")

    # Recommendations
    recs = ["• Atrasties ģimenes ārsta un ārstējošā urologa uzraudzībā.",
            "• Diēta – vispusīga, lietot šķidrumu 1.5 – 2.0 L/dnn, ja nav medicīnisku kontrindikāciju."]
    if op in URETHROPLASTY:
        recs += ["• Režīms – saudzējošs 4 nedēļas no operācijas dienas; 6 nedēļas izvairīties no riteņbraukšanas, "
                 "smagumu celšanas un ilgstošas sēdēšanas uz cietas virsmas.",
                 "• Atturēties no dzimumdzīves 4–6 nedēļas no operācijas dienas."]
    elif op in ("Meatotomy", "Meatoplasty"):
        recs += ["• Režīms – saudzējošs 2 nedēļas no operācijas dienas, izvairīties no pārmērīgas fiziskas slodzes, "
                 "pirts vai baseina apmeklējuma.",
                 "• Atturēties no dzimumdzīves 3 nedēļas no operācijas dienas.",
                 "• Brūces apstrāde: apmazgāt ar viegli sārtas krāsas kālija permanganāta šķīdumu un apstrādāt ar "
                 "Betadine 100 mg/ml; diegu evakuācija nav nepieciešama, ja lietoti uzsūcošie diegi."]
    else:
        rest = WEEKS_TEXT.get(v.get("Rest after endoscopic treatment") or "", "1 nedēļu")
        recs += [f"• Režīms – saudzējošs {rest} no operācijas dienas, izvairīties no pārmērīgas fiziskas slodzes, "
                 "necelt smagumus virs 10 kg."]
    if catheter_in:
        recs += ["• Kopt urīna ilgkatetru: uzturēt noslēgtu drenāžas sistēmu, maisu novietot zem urīnpūšļa līmeņa, "
                 "sekot urīna attecei; urīna savākšanas maisa maiņa ik pēc 3–7 dienām."]
        removal = v.get("Planned catheter removal (N35)")
        if removal:
            place = CATHETER_PLACE_N35.get(v.get("Catheter removal place (N35)"), "")
            time = (v.get("Catheter removal time (N35)") or "").strip()
            recs.append(f"• Urīna ilgkatetra evakuācija plānota{place} {fmt_date(removal)}"
                        + (f" plkst. {time}" if time else "") + ".")
        else:
            recs.append("• Urīna ilgkatetra evakuācija plānota pēc ārstējošā urologa norādījumiem.")
    if v.get("Self-dilation taught"):
        recs.append("• Turpināt urīnizvadkanāla pašdilatāciju (pašbužēšanu) saskaņā ar sniegtajām instrukcijām.")
    if cause == "Lichen sclerosus":
        recs.append("• Lichen sclerosus dēļ ilgstoša dzimumorgānu ādas novērošana pie urologa vai dermatologa.")
    if blocks:
        recs.append("• Terapijā lietot:")
        recs += ["    o " + b for b in blocks]
    if v["Tissue for histology"] == "Pending":
        recs.append("• Histoloģiskās izmeklēšanas rezultātu izvērtēs ārstējošais urologs.")
    if op in URETHROPLASTY:
        recs.append("• Kontrole pie urologa ambulatori pēc 3 mēnešiem ar simptomu un seksuālās funkcijas novērtējumu; "
                    "turpmāk pēc 12 un 24 mēnešiem.")
    else:
        recs.append("• Kontrole pie urologa ambulatori pēc 3 mēnešiem ar simptomu novērtējumu.")
    recs += tests_lines(v, "bullet", after_op=True)  # e.g. urethrogram before catheter removal, uroflowmetry
    recs.append("• Ja nav iespējams urinēt, urīna strūkla kļūst vāja, ir traucēta urīna attece pa katetru, parādās "
                "drudzis, izteikta hematūrija vai brūces asiņošana, nekavējoties vērsties NMC.")
    return {"clinical_diagnosis": diagnosis, "epicrisis": epicrisis, "recommendations": "\n".join(recs)}


def n35_checks(v):
    blockers, cautions = [], []
    op, location = v.get("Stricture operation"), v.get("Stricture location")
    length = v.get("Stricture length (cm)")
    if v.get("Voiding after removal (N35)") == "Urinary retention":
        blockers.append("Urinary retention after catheter removal.")
    if op in ("Meatotomy", "Meatoplasty") and location != "Meatus":
        blockers.append("Meatotomy and meatoplasty are operations for a meatal stricture: check the location.")
    if v.get("Tissue for histology") == "Known" and blank(v.get("Histology result (N35)")):
        blockers.append("Type the histology result.")
    if op in ENDOSCOPIC:
        removal = v.get("Planned catheter removal (N35)")
        if (v.get("Catheter at discharge (N35)") == "In place" and removal and v.get("Procedure date")
                and (removal - v["Procedure date"]).days > 3):
            cautions.append("EAU: remove the catheter within 72 hours after uncomplicated urethrotomy or dilation.")
        if v.get("Previous stricture treatment"):
            cautions.append("Recurrent stricture: EAU (strong) – do not repeat DVIU/dilatation more than twice if "
                            "urethroplasty is a viable option.")
        if location == "Penile urethra" and op == "Internal urethrotomy (DVIU)":
            cautions.append("EAU (strong): do not use DVIU for penile strictures – augmentation urethroplasty is the "
                            "standard.")
        if length not in (None, "") and length > 2:
            cautions.append("EAU (strong): do not use DVIU/dilatation as the only treatment for strictures longer "
                            "than 2 cm.")
        if location == "Membranous (posterior) urethra":
            cautions.append("EAU: DVIU/dilatation is meant for short bulbar strictures; membranous strictures carry a "
                            "risk of incontinence – check the indication.")
        if location == "Penile urethra" and op == "Urethral dilation":
            cautions.append("Penile stricture: EAU (strong) recommends augmentation urethroplasty.")
    if (op in URETHROPLASTY and v.get("Catheter at discharge (N35)") == "In place"
            and "rug_pericath" not in (v.get("_test_ids") or [])):
        cautions.append("EAU (strong recommendation): perform a urethrogram before catheter removal after urethroplasty "
                        "– tick it under Follow-up tests.")
    if (v.get("Stricture cause") == "Lichen sclerosus" or location == "Meatus") and v.get("Tissue for histology") == "Not sent":
        cautions.append("Meatal stricture / suspected lichen sclerosus: consider a biopsy for histological confirmation.")
    if v.get("Postoperative course") == "With complications":
        cautions.append("Complications are recorded: check that the recommendations cover them.")
    return blockers, cautions


N35_APPEARANCE = {"Clear": "dzidra", "Cloudy": "duļķaina", "Blood-tinged": "ar asiņu piejaukumu",
                  "With clots": "ar asins recekļiem"}


def n35_acute_parts(sex, v, blocks=()):
    """NMC: acute urinary retention due to a stricture, epicystostomy, sent home."""
    f = SEX_FORMS["Man"]
    day = fmt_date(v["Procedure date"])
    first = STRICTURE_DX[v["Stricture location"]]
    if v.get("Previous stricture treatment"):
        first = "Recidivējoša " + first[0].lower() + first[1:]
    length = v.get("Stricture length (cm)")
    if length not in (None, ""):
        first += f" ({length:g} cm)"
    diagnosis = f"{first}. Akūta urīna retence, st. pēc epicistostomijas ({day})."
    if v.get("Stricture cause") in STRICTURE_CAUSE_DX:
        diagnosis += "\n" + STRICTURE_CAUSE_DX[v["Stricture cause"]]

    history = STRICTURE_EPI[v["Stricture location"]].replace("striktūru", "striktūra").replace("sašaurinājumu", "sašaurinājums")
    sentences = [f"{f['Pacients']} {f['iestajies']} akūtā kārtā PSKUS NMC ar sūdzībām par nespēju urinēt. "
                 f"Anamnēzē {history}" + (", iepriekš ārstēta" if v.get("Previous stricture treatment") else "") + "."]
    if v.get("Failed urethral catheterisation"):
        sentences.append("Mēģinājums ievadīt urīnpūšļa katetru transuretrāli nesekmīgs urīnizvadkanāla sašaurinājuma dēļ.")
    guided = " USS kontrolē" if v.get("Ultrasound guided") else ""
    urine = f"izdalās {v['Drained volume (ml)']:g} ml"
    if v.get("Urine appearance") in N35_APPEARANCE:
        urine += f" {N35_APPEARANCE[v['Urine appearance']]}"
    sentences.append(f"{day} vietējā anestēzijā{guided} veikta epicistostomija, ievadīts Fr{v['Catheter Fr (N35)']} "
                     f"epicistostomas katetrs, {urine} urīna.")
    sentences.append("Paņemtas asins un urīna analīzes, urīna uzsējums. Dotajā brīdī nav indikāciju neatliekamai "
                     "stacionēšanai uroloģijas nodaļā. Vispārēji apmierinošā stāvoklī un ar rekomendācijām izraksta "
                     "tālākai ambulatorai ārstēšanai.")

    numbered = make_counter()
    recs = [
        numbered("Atrasties ģimenes ārsta un ārstējošā urologa uzraudzībā"),
        numbered("Lietot daudz šķidruma, vismaz 1.5-2.5 l/DN"),
        numbered("Kopt epicistostomas katetru: stomas vietu ikdienā apstrādāt un pārsiet, uzturēt noslēgtu drenāžas "
                 "sistēmu, maisu novietot zem urīnpūšļa līmeņa. Katetra maiņa ik pēc 4–6 nedēļām ģimenes ārsta praksē, "
                 "urīna savācējmaisa maiņa ik pēc 3–7 dienām."),
        numbered("Uretrālu katetrizāciju vai urīnizvadkanāla dilatāciju neveikt bez urologa konsultācijas."),
    ]
    if blocks:
        recs.append(numbered("Terapijā lietot:"))
        recs += ["    " + b for b in blocks]
    recs += [
        numbered("Laboratorisko izmeklējumu rezultātus var aplūkot Datamed.lv. Urīna uzsējuma rezultāti pēc 3-5. dienām"),
        numbered(f"Ambulatori urologa konsultācija {CONSULT_TEXT[v['Urology consultation']]}, lai izvērtētu striktūru "
                 "un plānotu definitīvo ārstēšanu."),
    ]
    recs += tests_lines(v, "numbered", numbered)  # e.g. urethrography before the definitive treatment
    recs += [
        numbered("Ja epicistostomas katetrs izkrīt, nav urīna atteces, parādās drudzis vai izteikta asiņošana, "
                 "nekavējoties vērsties NMC – izkritušā katetra kanāls var ātri slēgties."),
    ]
    return {"clinical_diagnosis": diagnosis, "epicrisis": " ".join(sentences), "recommendations": "\n".join(recs)}


def n35_all_checks(v):
    if v.get("_department") != "NMC":
        return n35_checks(v)
    blockers, cautions = [], []
    if v.get("Fever at NMC"):
        blockers.append("Fever with urinary retention: assess for urosepsis before sending the patient home.")
    if v.get("Creatinine raised (N35)") or (v.get("Drained volume (ml)") or 0) > 1000:
        cautions.append("Large retention volume or raised creatinine: risk of post-obstructive diuresis and renal "
                        "impairment - consider admission for monitoring.")
    if v.get("Urine appearance") == "With clots":
        cautions.append("Clots in the urine: check for clot retention and the source of bleeding.")
    return blockers, cautions


BUILDERS.update({"n35": n35_parts, "n35_acute": n35_acute_parts})
CHECKS["N35"] = n35_all_checks


# ---------------------------------------------------------------
# 2i. NEPHROSTOMY CHANGE (Z43.6) in 48. daily department
# ---------------------------------------------------------------


def neph_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    side = v["Nephrostomy side"]
    adjective, change, op_name = NEPH_SIDE[side]
    both = side == "Both"
    day = fmt_date(v["Procedure date"])
    ch = v["Nephrostomy Ch"]

    # Clinical diagnosis with co-diagnoses
    lines = [("Abpusējas nefrostomas" if both else f"{adjective.capitalize()} nefrostoma")
             + f", st. pēc nefrostom{'u' if both else 'as'} maiņas ({day})."]
    codx = [NEPHROSTOMY_CODIAGNOSES.get(c, c) for c in (v.get("Nephrostomy co-diagnoses") or [])]
    if codx:
        lines.append("Blakusdiagnozes:")
        lines += [c.rstrip(".") + "." for c in codx]

    # Epicrisis
    patient_acc = "Pacienti" if sex == "Woman" else "Pacientu"
    target = f"Ch{ch} izmēra nefrostom{'ām' if both else 'u'}"
    sentences = [
        f"{f['Pacients']} iestājās P. Stradiņa KUS 4. uroloģijas nodaļas 48. dienas stacionārā plānotai "
        + (f"abpusējai nefrostomu maiņai." if both else f"{adjective} nefrostomas maiņai."),
        f"Sterīlā veidā tiek veikta {change} uz {target}"
        + (", veikta skalošana" if v.get("Irrigation performed") else "")
        + (", funkcija apmierinoša." if v["Nephrostomy function"] == "Satisfactory" else ", funkcija pavājināta."),
    ]
    if v.get("Infection signs at change"):
        sentences.append("Izdalās duļķains urīns, paņemts urīna uzsējums.")
    sentences.append(("Nefrostomas fiksē" if both else "Nefrostomu fiksē")
                     + (" ar ādas šuvēm" if v.get("Fixed with skin sutures") else "") + ", uzliek sterilu pārsēju.")
    sentences.append(f"{patient_acc} apmierinošā stāvoklī izraksta tālākai terapijai ambulatori.")
    investigations = (v.get("Investigations (48)") or "").strip()
    epicrisis = " ".join(sentences) + f"\n\nIzmeklējumi: {investigations}\n\nOperācijas: {op_name} ({day})"

    # Recommendations
    nxt = v.get("Next nephrostomy change")
    when = fmt_date(nxt) if nxt else "pēc 3 mēnešiem (datumu precizēt)"
    recs = [
        f"• Atrasties {SUPERVISOR_TEXT[v['Supervising doctor']]} uzraudzībā!",
        "• Lietot šķidrumu vismaz 2l/DN",
        f"• Nākamā nefrostomas maiņa paredzēta PSKUS 4. uroloģijas nodaļas DS {when}. Ierasties ar ģimenes ārsta nosūtījumu!",
        "• Kopt nefrostomu:",
        "    o pārsiešanu veikt vismaz 2 reizes nedēļā, vai biežāk pēc nepieciešamības;",
        "    o nefrostomas ievades vietu apstrādā ar dezinficējošu šķīdumu (Braunoli, Prontosan, u.c.);",
        "    o nefrostomai pārsēju jāuzliek tā, lai tās ieejas vieta atrastos zem pārsēja (plākstera);",
        "    o nefrostomas urīnsavācējmaisu maina vismaz 1 reizi nedēļā, vai arī, ja tas ir bojāts.",
        "• Nelietot nefrotoksiskus medikamentus, pretsāpju zāles - Ibumetins, Diklofenaks. Izvēles pretsāpju "
        "medikamenti ir Analgin, Ultracod, Paracetamol, Tramadol, Promedol.",
    ]
    if v.get("Infection signs at change"):
        recs.append("• Urīna uzsējuma rezultātus var aplūkot Datamed.lv pēc 3–5 dienām; antibakteriālo terapiju "
                    "uzsākt atbilstoši jutībai, ja ir infekcijas simptomi.")
    recs += ["• " + b for b in blocks]
    recs += tests_lines(v, "bullet")
    recs.append("• Ja nefrostoma izkrīt, nav urīna atteces, parādās drudzis, drebuļi, stipras sāpes sānos vai izteikta "
                "asiņošana, nekavējoties vērsties NMC.")
    return {"clinical_diagnosis": "\n".join(lines), "epicrisis": epicrisis, "recommendations": "\n".join(recs)}


def neph_checks(v):
    blockers, cautions = [], []
    if v.get("Fever (48)"):
        blockers.append("Fever at the nephrostomy change: assess for infection or obstruction before sending the patient home.")
    if v.get("Nephrostomy function") == "Reduced":
        cautions.append("Reduced nephrostomy function: check the tube position (e.g. antegrade pyelography) before discharge.")
    if v.get("Infection signs at change"):
        cautions.append("Cloudy urine at the change: colonisation is common; treat only if there are infection symptoms.")
    nxt, day = v.get("Next nephrostomy change"), v.get("Procedure date")
    if not nxt:
        cautions.append("Give a concrete date for the next nephrostomy change.")
    elif day and (nxt - day).days > 92:
        cautions.append("The next change is more than 3 months away: long intervals raise the risk of encrustation and blockage.")
    if not v.get("Nephrostomy co-diagnoses"):
        cautions.append("No co-diagnosis (cause of the nephrostomy) is chosen.")
    return blockers, cautions


BUILDERS["nephrostomy"] = neph_parts
CHECKS["Z43.6"] = neph_checks


# ---------------------------------------------------------------
# 2k-b. NEPHROSTOMY FELL OUT (Z43.6): new nephrostomy by the radiologist at NMC, sent home or admitted to 4.
# ---------------------------------------------------------------
# A dislodged nephrostomy leaves the kidney obstructed again. The tract closes within hours to days, so a new
# percutaneous nephrostomy is placed under ultrasound guidance. EAU (urolithiasis / urological infections): an
# obstructed kidney with signs of infection is an emergency – urgent decompression and antibiotics, i.e. admission.
# Admit also with raised creatinine in a solitary kidney, a complicated or failed puncture, ongoing bleeding,
# uncontrolled pain or a poor general condition. Premedication used in the department: morphine 1 amp. i/m and
# diazepam 1 amp. i/m right before the procedure (opioid + benzodiazepine: monitor breathing, no driving that day).

NF_SIDE = {  # (after "sakarā ar", diagnosis, "St. pēc ...", "tiek izveidota ...", "fell out" sentence)
    "Right": ("izkritušu labās puses nefrostomu", "Izkritusi labās puses nefrostoma.", "labās puses nefrostomijas",
              "tiek izveidota labās puses nefrostoma", "Nefrostoma izkritusi"),
    "Left": ("izkritušu kreisās puses nefrostomu", "Izkritusi kreisās puses nefrostoma.", "kreisās puses nefrostomijas",
             "tiek izveidota kreisās puses nefrostoma", "Nefrostoma izkritusi"),
    "Both": ("izkritušām abpusējām nefrostomām", "Izkritušas abpusējas nefrostomas.", "abpusējas nefrostomijas",
             "tiek izveidotas abpusējas nefrostomas", "Nefrostomas izkritušas"),
}
NF_TIME = {"Less than 24 hours": "mazāk nekā pirms 24 stundām", "1–3 days": "pirms 1–3 dienām",
           "More than 3 days": "vairāk nekā pirms 3 dienām"}
NF_SIGNS = {"Fever 38 °C or more": "drudzi", "Chills": "drebuļiem", "Hypotension / tachycardia": None,
            "Flank pain not controlled": "stiprām sāpēm sānā"}
NF_BLOOD = {"No inflammatory changes": "Asins analīzēs bez iekaisuma izmaiņām.",
            "Raised CRP / leukocytosis": "Asins analīzēs – paaugstināts CRO / leikocitoze.",
            "Raised creatinine": "Asins analīzēs – paaugstināts kreatinīna līmenis.",
            "Raised CRP and creatinine": "Asins analīzēs – paaugstināts CRO un kreatinīna līmenis."}
NF_URINE = {"Clear": "Pa nefrostomu izdalās dzidrs urīns.",
            "Blood-tinged": "Pa nefrostomu izdalās urīns ar nelielu asiņu piejaukumu.",
            "Cloudy / purulent": "Pa nefrostomu izdalās duļķains urīns."}
NF_PREMED = ("Pirms manipulācijas veikta premedikācija – Sol. Morphini hydrochloridi 1 amp. (10 mg) i/m un "
             "Sol. Diazepami 1 amp. (10 mg) i/m.")
NF_CARE = [
    "• Kopt nefrostomu:",
    "    o pārsiešanu veikt vismaz 2 reizes nedēļā, vai biežāk pēc nepieciešamības;",
    "    o nefrostomas ievades vietu apstrādā ar dezinficējošu šķīdumu (Braunoli, Prontosan, u.c.);",
    "    o nefrostomai pārsēju jāuzliek tā, lai tās ieejas vieta atrastos zem pārsēja (plākstera);",
    "    o nefrostomas urīnsavācējmaisu maina vismaz 1 reizi nedēļā, vai arī, ja tas ir bojāts;",
    "    o urīnsavācējmaisu nostiprināt pie ķermeņa vai kājas un nenovietot augstāk par nieri; nefrostomu nelocīt "
    "un neraut – tā var izslīdēt.",
]
NF_PAIN = ("• Nelietot nefrotoksiskus medikamentus, pretsāpju zāles - Ibumetins, Diklofenaks. Izvēles pretsāpju "
           "medikamenti ir Analgin, Ultracod, Paracetamol, Tramadol, Promedol.")
NF_PLAN = {
    "Urologist: plan definitive treatment": "• Ambulatori urologa konsultācija – lemt par obstrukcijas definitīvu "
                                            "ārstēšanu (DJ stents, konkrementa evakuācija u.c.), lai izvairītos no "
                                            "ilgstošas nefrostomas.",
    "Oncologist: oncology plan": "• Turpināt onkologa uzraudzību un ārstēšanu; par nefrostomas turpmāko taktiku lemt "
                                 "kopā ar urologu.",
    "Long-term nephrostomy": "",
}
NF_WARNING = ("• Ja nefrostoma atkal izslīd vai izkrīt, nav urīna atteces, parādās drudzis, drebuļi, stipras sāpes sānos "
              "vai izteikta asiņošana, nekavējoties vērsties NMC – jo ātrāk, jo vieglāk nefrostomu atjaunot.")


def _nf_infected(v):
    return any(x in (v.get("Systemic signs (NF)") or []) for x in ("Fever 38 °C or more", "Chills",
                                                                   "Hypotension / tachycardia"))


def _nf_diagnosis(v):
    words = NF_SIDE[v["Nephrostomy side"]]
    lines = [words[1], f"St. pēc {words[2]} ({fmt_date(v['Procedure date'])})."]
    if v.get("_department") == "4" and _nf_infected(v):
        lines.insert(1, "Akūts obstruktīvs pielonefrīts.")
    codx = [NEPHROSTOMY_CODIAGNOSES.get(c, c) for c in (v.get("Nephrostomy co-diagnoses") or [])]
    if codx:
        lines.append("Blakusdiagnozes:")
        lines += [c.rstrip(".") + "." for c in codx]
    return "\n".join(lines)


def _nf_nmc_sentences(f, v):
    words = NF_SIDE[v["Nephrostomy side"]]
    s = [f"{words[4]} {NF_TIME[v['Time since it fell out']]}."]
    signs = [NF_SIGNS[x] for x in (v.get("Systemic signs (NF)") or []) if NF_SIGNS.get(x)]
    if not v.get("Systemic signs (NF)"):
        s.append("Drudzi un drebuļus noliedz.")
    if v.get("Solitary kidney"):
        s.append("Vienīgā funkcionējošā niere.")
    us = v.get("Hydronephrosis on US (NF)")
    s.append(f"NMC etapā vispusīgi {f['izmekleta']}." +
             {"Present": " USS – hidronefroze.", "Not present": " USS – izteiktu hidronefrozi nekonstatē."}.get(us, ""))
    if v.get("Blood tests (NF)") in NF_BLOOD:
        s.append(NF_BLOOD[v["Blood tests (NF)"]])
    if not blank(v.get("Other NMC findings (optional)")):
        text = v["Other NMC findings (optional)"].strip()
        s.append(text if text.endswith(".") else text + ".")
    if v.get("Premedication given"):
        s.append(NF_PREMED)
    result = v["Puncture result (NF)"]
    s.append(f"{fmt_date(v['Procedure date'])} ultrasonogrāfijas kontrolē {words[3]} (Ch{v['Nephrostomy Ch']})" +
             {"Without complications": ", manipulācija noritēja bez komplikācijām.",
              "Bleeding through the tube": "; pa nefrostomu sākotnēji neliela asiņošana, kas mazinās.",
              "Complicated / not successful": "; manipulācija noritēja ar sarežģījumiem."}[result])
    s.append(NF_URINE[v["Urine from new nephrostomy"]] +
             (" Paņemts urīna uzsējums." if v.get("Urine culture taken (NF)") else ""))
    both = v["Nephrostomy side"] == "Both"
    s.append(("Nefrostomas fiksē" if both else "Nefrostomu fiksē") +
             (" ar ādas šuvēm" if v.get("Fixed with skin sutures") else "") + ", uzliek sterilu pārsēju.")
    return s, signs


def _nf_recs(v, blocks):
    nxt = v.get("Next nephrostomy change")
    when = fmt_date(nxt) if nxt else "pēc 3 mēnešiem (datumu precizēt)"
    recs = [f"• Atrasties {SUPERVISOR_TEXT[v['Supervising doctor']]} uzraudzībā!",
            "• Lietot šķidrumu vismaz 2 l/DN, ja nav medicīnisku kontrindikāciju."]
    recs += NF_CARE
    recs.append(f"• Nākamā nefrostomas maiņa paredzēta PSKUS 4. uroloģijas nodaļas DS {when}. Ierasties ar ģimenes "
                "ārsta nosūtījumu!")
    if NF_PLAN.get(v.get("Definitive plan (NF)")):
        recs.append(NF_PLAN[v["Definitive plan (NF)"]])
    recs.append(NF_PAIN)
    if v.get("Premedication given") and v.get("_department") == "NMC":
        recs.append("• Manipulācijas dienā nevadīt transportlīdzekli un neatrasties vienam (premedikācija ar morfīnu un "
                    "diazepāmu).")
    if v.get("Urine culture taken (NF)") and blank(v.get("Urine culture result (NF)")):
        recs.append("• Urīna uzsējuma rezultātus var aplūkot Datamed.lv pēc 3–5 dienām; antibakteriālo terapiju "
                    "koriģēt atbilstoši jutībai, ja ir infekcijas simptomi.")
    if blocks:
        recs.append("• Terapijā lietot:")
        recs += ["    o " + b for b in blocks]
    recs += tests_lines(v, "bullet")
    recs.append(NF_WARNING)
    return recs


def nf_nmc_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    words = NF_SIDE[v["Nephrostomy side"]]
    s, signs = _nf_nmc_sentences(f, v)
    first = f"{f['Pacients']} {f['iestajies']} akūtā kārtā PSKUS NMC sakarā ar {words[0]}."
    s = [first] + s
    s.append("Pēc manipulācijas novērots, stāvoklis stabils. Dotajā brīdī nav indikāciju neatliekamai stacionēšanai. "
             "Vispārēji apmierinošā stāvoklī un ar rekomendācijām izraksta tālākai ambulatorai ārstēšanai.")
    if sex == "Woman":
        s[-1] = s[-1].replace("novērots", "novērota")
    return {"clinical_diagnosis": _nf_diagnosis(v), "epicrisis": " ".join(s),
            "recommendations": "\n".join(_nf_recs(v, blocks))}


def nf_ward_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    words = NF_SIDE[v["Nephrostomy side"]]
    s, signs = _nf_nmc_sentences(f, v)
    first = (f"{f['Pacients']} {f['stacionets']} P. Stradiņa KUS 4. uroloģijas nodaļā akūtā kārtā sakarā ar "
             f"{words[0]}" + (f", {lv_list(signs)}" if signs else "") + ".")
    p1 = [first] + s
    treat = v.get("Ward treatment (NF)") or []
    acts = []
    if "IV antibiotics" in treat:
        abx = (v.get("Antibiotic in the ward (NF)") or "").strip().rstrip(".")
        acts.append("nozīmēta i/v antibakteriālā terapija" + (f" ({abx})" if abx else ""))
    if "IV fluids" in treat:
        acts.append("infūzu terapija")
    if "Analgesia" in treat:
        acts.append("analgēzija")
    p2 = ["Nodaļā " + lv_list(acts) + "."] if acts else []
    result = (v.get("Urine culture result (NF)") or "").strip().rstrip(".")
    if result:
        p2.append(f"Urīna uzsējumā – {result}.")
    p2.append("Stāvoklis uzlabojas, normalizējas ķermeņa temperatūra, iekaisuma rādītāji mazinās."
              if v.get("Ward course (NF)") == "Fever resolved, CRP falling" else "Periods nodaļā bez īpatnībām.")
    p2.append("Nefrostoma funkcionē apmierinoši.")
    epicrisis = " ".join(p1) + "\n" + " ".join(p2) + \
        f"\nVispārējā apmierinošā stāvoklī {f['izrakstits']} tālākai ārstēšanai ambulatori."
    return {"clinical_diagnosis": _nf_diagnosis(v), "epicrisis": epicrisis,
            "recommendations": "\n".join(_nf_recs(v, blocks))}


def nf_checks(v):
    blockers, cautions = [], []
    nmc = v.get("_department") == "NMC"
    blood = v.get("Blood tests (NF)")
    crp = blood in ("Raised CRP / leukocytosis", "Raised CRP and creatinine")
    creat = blood in ("Raised creatinine", "Raised CRP and creatinine")
    result = v.get("Puncture result (NF)")
    signs = v.get("Systemic signs (NF)") or []
    if nmc:
        if _nf_infected(v):
            blockers.append("Obstructed kidney with fever, chills or unstable circulation: EAU – urgent drainage plus "
                            "antibiotics; admit to 4. Urology (use department 4).")
        if "Flank pain not controlled" in signs:
            blockers.append("Pain is not controlled: admit.")
        if result == "Complicated / not successful":
            blockers.append("Complicated or failed puncture: admit (bleeding, urinoma, no drainage).")
        if creat and v.get("Solitary kidney"):
            blockers.append("Raised creatinine in a solitary kidney: admit and follow renal function.")
        if crp:
            cautions.append("Raised CRP / leukocytosis without fever: infection of the obstructed kidney is possible – "
                            "consider admission; at least take a culture and recheck in 48–72 hours.")
        if creat and not v.get("Solitary kidney"):
            cautions.append("Raised creatinine: recheck in 48–72 hours (tests box); morphine accumulates in renal "
                            "impairment.")
        if result == "Bleeding through the tube":
            cautions.append("Bleeding through the new tube: observe until it clears; admit if it continues, clots block "
                            "the tube or Hb falls.")
        if v.get("Premedication given"):
            cautions.append("Morphine + diazepam: risk of respiratory depression and sedation – observe breathing and "
                            "SpO2 until fully awake before discharge; lower doses in the elderly, COPD or renal "
                            "impairment.")
    else:
        if _nf_infected(v) and "IV antibiotics" not in (v.get("Ward treatment (NF)") or []):
            cautions.append("Infected obstructed kidney: IV antibiotics are expected (EAU).")
        if _nf_infected(v) and not any(b in (v.get("_blocks") or []) for b in ("cefuroxime-7", "amoxiclav-7")):
            cautions.append("Obstructive pyelonephritis: continue with an oral antibiotic by culture to complete 7–14 "
                            "days in total (choose a block), unless the full course was given in the ward.")
    if v.get("Urine from new nephrostomy") == "Cloudy / purulent" and not v.get("Urine culture taken (NF)"):
        cautions.append("Cloudy / purulent urine from the kidney (possible pyonephrosis): take a culture.")
    if v.get("Time since it fell out") == "More than 3 days" and v.get("Hydronephrosis on US (NF)") == "Not present":
        cautions.append("No hydronephrosis after more than 3 days without the tube: check whether a nephrostomy is "
                        "still needed (the obstruction may have resolved).")
    if not v.get("Nephrostomy co-diagnoses"):
        cautions.append("No co-diagnosis (cause of the nephrostomy) is chosen.")
    if not v.get("Next nephrostomy change"):
        cautions.append("Give a concrete date for the next nephrostomy change.")
    if any(b for b in (v.get("_blocks") or []) if b in ("cefuroxime-7", "amoxiclav-7")) and not (
            _nf_infected(v) or crp or v.get("Urine from new nephrostomy") == "Cloudy / purulent"):
        cautions.append("Antibiotic without infection signs: nephrostomy urine is usually colonised – EAU: do not treat "
                        "asymptomatic bacteriuria.")
    return blockers, cautions


def z43_checks(v):
    return neph_checks(v) if v.get("_department") == "48" else nf_checks(v)


BUILDERS["nf_nmc"] = nf_nmc_parts
BUILDERS["nf_ward"] = nf_ward_parts
CHECKS["Z43.6"] = z43_checks


# ---------------------------------------------------------------
# 2l. BPH (N40): TURP discharge
# ---------------------------------------------------------------


def n40_parts(sex, v, blocks=()):
    f = SEX_FORMS["Man"]
    indications = [BPH_INDICATION_TEXT[i] for i in v["BPH indication"]]
    complicated = any(i in BPH_COMPLICATED for i in v["BPH indication"])
    day = fmt_date(v["Procedure date"])
    volume = v.get("Prostate volume (ml)")
    admission = "akūtā kārtā" if v["Admission"] == "Acute" else "plānveida kārtā"
    intraop, postop = c61_courses(v)

    dx_name = "Komplicēta LPH" if complicated else "LPH"
    diagnosis = f"{dx_name} ({', '.join(indications)})."
    if volume not in (None, ""):
        diagnosis += f" Prostatas tilpums ~{volume:g} ml."
    diagnosis += f"\nSt. pēc TURP ({day})."

    reason = f"iepriekš diagnosticētu {'komplicētu ' if complicated else ''}LPH ({', '.join(indications)})"
    if volume not in (None, ""):
        reason += f", prostatas tilpums ~{volume:g} ml"
    first = (f"{f['Pacients']} {f['stacionets']} P. Stradiņa KUS 4. uroloģijas nodaļā {admission} operatīvai "
             f"ārstēšanai sakarā ar {reason}.")
    if v.get("Preoperative catheter"):
        first += " Iestājoties ar urīnpūšļa ilgkatetru."
    first += f" {f['Pacients']} tiek {f['gatavots']} operatīvai terapijai."
    course = []
    if v.get("Bladder irrigation after surgery"):
        course.append("Pēcoperācijas periodā veikta urīnpūšļa caurskalošana.")
    course.append(f"Pēcoperācijas periods {postop}.")
    course.append(c61_catheter_sentence(f, v))
    if v["Histology"] == "Pending":
        histology = "Histoloģiskās izmeklēšanas rezultāts izrakstīšanas brīdī vēl nav zināms."
    else:
        histology = f"Histoloģiskās izmeklēšanas rezultāts: {v['Histology result (BPH)'].strip().rstrip('.')}."
    epicrisis = "\n".join([
        first,
        f"{day} {ANESTHESIA_TEXT[v['TURP anesthesia']]} veikta {TURP_TECHNIQUE_TEXT[v['TURP technique']]} "
        f"transuretrāla prostatas rezekcija (TURP). Operācija noritēja {intraop}.",
        " ".join(course),
        f"{histology} Uzlabojoties vispārējam stāvoklim, {f['pacients']} {f['izrakstits']} turpmākai aprūpei "
        "ģimenes ārsta uzraudzībā. Sniegtas rekomendācijas.",
    ])

    recs = []
    if v["Discharge paper"] == "Temporary":
        recs.append(TEMPORARY_PAPER_TEXT)
    recs += [
        "Atrasties ģimenes ārsta un ārstējošā urologa uzraudzībā.",
        "Diēta: vispusīga. Lietot šķidrumu 1,5–2,0 L diennaktī, ja nav medicīnisku kontrindikāciju; pirmās nedēļas "
        "izvairīties no alkohola un kofeīnu saturošiem dzērieniem.",
        f"Režīms: saudzējošs {WEEKS_TEXT[v['Activity restriction']]} no operācijas dienas – necelt smagumus, "
        "nebraukt ar velosipēdu, izvairīties no pirts un karstām vannām.",
        "Izvairīties no aizcietējumiem un spiešanās defekācijas laikā.",
        "Atturēties no dzimumdzīves 4 nedēļas no operācijas dienas. Pēc operācijas iespējama retrogrāda ejakulācija "
        "(sēklas šķidrums nonāk urīnpūslī) – tā nav bīstama.",
        "Pirmajās nedēļās iespējama bieža un neatliekama urinācija, dedzināšana urinējot un neliels asiņu piejaukums "
        "urīnā, īpaši 10.–14. dienā pēc operācijas, kad atdalās brūces krevele; šādā gadījumā palielināt šķidruma "
        "uzņemšanu.",
    ]
    if v["Catheter removal"] == "Discharged with catheter":
        recs += CATHETER_CARE_LINES
        if v.get("Planned catheter removal (optional)"):
            recs.append(f"Katetra evakuācija plānota {fmt_date(v['Planned catheter removal (optional)'])}.")
    else:
        recs.append("Regulāri iztukšot urīnpūsli, nenoturēt urīnu ilgstoši.")
    if v["Previous BPH medication"] == "Stop after surgery":
        recs.append("Iepriekš lietotos prostatas medikamentus (alfa blokatorus, 5-alfa reduktāzes inhibitorus) "
                    "pārtraukt, ja urologs nav norādījis citādi.")
    elif v["Previous BPH medication"] == "Continue until review":
        recs.append("Iepriekš lietotos prostatas medikamentus turpināt līdz kontrolei pie urologa.")
    if blocks:
        recs.append("Terapijā lietot:")
        recs += ["• " + b for b in blocks]
    visit = f"Vizīte pie urologa ambulatori: {WEEKS_TEXT[v['Urology visit']]} pēc operācijas"
    if v["Histology"] == "Pending":
        visit += "; histoloģisko atbildi izvērtēs ārstējošais urologs"
    recs.append(visit + ".")
    recs += tests_lines(v, "plain", after_op=True)  # e.g. IPSS, uroflowmetry, residual urine at the visit
    recs.append("Ja parādās izteikta hematūrija ar asins recekļiem, nav iespējams urinēt, katetrs nefunkcionē, rodas "
                "drudzis vai drebuļi, nekavējoties vērsties NMC.")
    return {"clinical_diagnosis": diagnosis, "epicrisis": epicrisis, "recommendations": "\n".join(recs)}


def n40_checks(v):
    blockers, cautions = [], []
    if v.get("Voiding after catheter removal") == "Urinary retention":
        blockers.append("Urinary retention after catheter removal (failed trial without catheter).")
    if v.get("Histology") == "Known":
        result = (v.get("Histology result (BPH)") or "").lower()
        if blank(result):
            blockers.append("Type the histology result.")
        elif any(w in result for w in ("karcinom", "adenokarcinom", "ca ", "gleason", "isup")):
            cautions.append("Incidental prostate cancer in the TURP specimen: plan PSA, staging and oncology "
                            "council (and consider coding as C61).")
    if v.get("Histology") == "Pending" and v.get("Discharge paper") == "Final":
        cautions.append("Histology is still pending, but the discharge paper is marked as final.")
    if v.get("Histology") == "Known" and v.get("Discharge paper") == "Temporary":
        cautions.append("Histology is already known, but the discharge paper is marked as temporary.")
    chosen = v.get("_blocks") or []
    if "urorec" in chosen and v.get("Previous BPH medication") == "Stop after surgery":
        cautions.append("Urorec is chosen, while previous prostate medicines are to be stopped: check the plan.")
    if "solifenacin" in chosen and v.get("Catheter removal") != "Discharged with catheter":
        cautions.append("Solifenacin after catheter removal: watch for urinary retention.")
    if any(b in chosen for b in ("fraxiparine", "fraxiparine-04")) and v.get("Voiding after catheter removal") == \
            "Voids independently, mild hematuria":
        cautions.append("Fraxiparine with ongoing hematuria after TURP: weigh the bleeding risk.")
    if v.get("Postoperative course") == "With complications":
        cautions.append("Complications are recorded: check that the recommendations cover them.")
    return blockers, cautions


BUILDERS["n40"] = n40_parts
CHECKS["N40"] = n40_checks


# ---------------------------------------------------------------
# 2n. HYDROCELE (N43): hydrocele plasty in 4. Urology
# ---------------------------------------------------------------

HYDROCELE_SIDE = {"Right": ("Labās", "labās"), "Left": ("Kreisās", "kreisās")}
HYDROCELE_TECHNIQUE = {  # (in the epicrisis, short form in the diagnosis)
    "Eversion (Winkelmann)": ("ar tunica vaginalis izvēršanu (pēc Winkelmann)", "pēc Winkelmann"),
    "Excision (Bergmann)": ("ar tunica vaginalis parietālās lapiņas ekscīziju (pēc Bergmann)", "pēc Bergmann"),
    "Plication (Lord)": ("ar tunica vaginalis plikāciju (pēc Lord)", "pēc Lord"),
}


def n43_parts(sex, v, blocks=()):
    f = SEX_FORMS["Man"]  # hydrocele: always male wording
    side_cap, side = HYDROCELE_SIDE[v["Hydrocele side"]]
    day = fmt_date(v["Procedure date"])
    technique, technique_short = HYDROCELE_TECHNIQUE[v["Hydrocele technique"]]
    admission = "akūtā kārtā" if v["Admission"] == "Acute" else "plānveida kārtā"
    intraop, postop = c61_courses(v)

    diagnosis = f"{side_cap} puses hidrocēle.\nSt. pēc {side} puses hidrocēles plastikas {technique_short} ({day})."

    course = [f"Pēcoperācijas periods {postop}."]
    drain = v["Hydrocele drain"]
    if drain == "Removed on day 1":
        course.append("Aspirācijas drena evakuēta 1. pēcoperācijas dienā.")
    if v.get("Scrotal hematoma"):
        course.append("Novērojama neliela sēklinieku maisiņa hematoma, dinamikā bez palielināšanās.")
    course.append("Brūce tīra, sausa.")
    closing = (f"Uzlabojoties vispārējam stāvoklim, {f['pacients']} {f['izrakstits']} turpmākai aprūpei ģimenes ārsta "
               "uzraudzībā" + (" ar aspirācijas drenu" if drain == "Discharged with drain" else "") + ". Sniegtas rekomendācijas.")
    epicrisis = "\n".join([
        f"{f['Pacients']} {f['stacionets']} P. Stradiņa KUS 4. uroloģijas nodaļā {admission} operatīvai terapijai "
        f"sakarā ar iepriekš diagnosticētu {side} puses hidrocēli. {f['Pacients']} tiek {f['gatavots']} operatīvai "
        "terapijai.",
        f"{day} {ANESTHESIA_N47_TEXT[v['N43 anesthesia']]} veikta {side} puses hidrocēles plastika {technique}. "
        f"Operācija noritēja {intraop}.",
        " ".join(course),
        closing,
    ])

    recs = [
        "• Atrasties ģimenes ārsta uzraudzībā.",
        "• Diēta – vispusīga.",
        f"• Režīms – saudzējošs {WEEKS_TEXT[v['Activity restriction']]} no operācijas dienas: necelt smagumus, "
        "izvairīties no pārmērīgas fiziskas slodzes, riteņbraukšanas, pirts un baseina apmeklējuma.",
        "• Pirmās 1–2 nedēļas valkāt suspensoriju vai cieši pieguļošu apakšveļu; pirmajās 48 stundās sēklinieku "
        "maisiņam vairākas reizes dienā uz 15–20 minūtēm likt aukstu kompresi (caur drānu).",
        "• Atturēties no dzimumdzīves 3 nedēļas no operācijas dienas.",
        "• Brūces aprūpe: pārsiešana katru otro dienu vai biežāk pēc nepieciešamības, brūci apstrādāt ar Braunol; "
        "dušā drīkst mazgāties 48 stundas pēc operācijas, brūci nemērcēt (vanna, baseins) līdz tās sadzīšanai.",
    ]
    if v["Skin sutures"] == "Absorbable":
        recs.append("• Ādas šuves uzsūcošās, to evakuācija nav nepieciešama.")
    else:
        recs.append("• Ādas šuvju evakuācija 10.–14. pēcoperācijas dienā ģimenes ārsta praksē.")
    if drain == "Discharged with drain":
        when = fmt_date(v["Drain removal date"]) if v.get("Drain removal date") else "pēc ārstējošā urologa norādījumiem"
        recs.append(f"• Aspirācijas drenas evakuācija plānota {when}; sekot izdalījumu daudzumam.")
    recs.append("• Pirmajās nedēļās iespējama sēklinieku maisiņa tūska un zilums, kas pakāpeniski izzūd; pilnīga "
                "tūskas izzušana var ilgt vairākas nedēļas.")
    if blocks:
        recs.append("• Terapijā lietot:")
        recs += ["    o " + b for b in blocks]
    recs.append(f"• Kontrole pie urologa ambulatori {WEEKS_TEXT[v['Urology visit']]} pēc operācijas.")
    recs += tests_lines(v, "bullet", after_op=True)
    recs.append("• Ja sēklinieku maisiņš strauji palielinās vai kļūst sasprindzināts, pastiprinās sāpes, brūce asiņo vai "
                "izdalās strutas, paaugstinās ķermeņa temperatūra, nekavējoties vērsties NMC.")
    return {"clinical_diagnosis": diagnosis, "epicrisis": epicrisis, "recommendations": "\n".join(recs)}


def n43_checks(v):
    blockers, cautions = [], []
    if not v.get("Scrotal ultrasound before surgery"):
        cautions.append("No scrotal ultrasound before surgery: in adults a hydrocele can hide a testicular tumour "
                        "(reactive hydrocele) – check that the testis was assessed.")
    if v.get("Scrotal hematoma"):
        cautions.append("Scrotal hematoma: make sure it is not growing before discharge; consider a scrotal ultrasound.")
    if v.get("Hydrocele drain") == "Discharged with drain" and not v.get("Drain removal date"):
        cautions.append("The patient goes home with a drain: give the planned removal date.")
    if "With complications" in (v.get("Intraoperative course"), v.get("Postoperative course")):
        cautions.append("Complications are recorded: check that the recommendations cover them.")
    if v.get("N43 anesthesia") == "Local" and any(b in (v.get("_blocks") or []) for b in ("fraxiparine", "fraxiparine-04")):
        cautions.append("Fraxiparine after a short procedure under local anesthesia: check the indication "
                        "(and the bleeding risk in the scrotum).")
    return blockers, cautions


BUILDERS["n43"] = n43_parts
CHECKS["N43"] = n43_checks


# ---------------------------------------------------------------
# 2o. PUJ STRICTURE (N13.0): pyeloplasty in 4. Urology
# ---------------------------------------------------------------
# There is no adult EAU guideline on PUJ obstruction (EAU covers it only in paediatric urology),
# so the follow-up follows common practice: DJ stent out after 4–6 weeks, ultrasound after that,
# diuretic renography (MAG3) about 3–6 months after surgery.

PUJ_CAUSE_DX = {"Intrinsic (congenital) stenosis": "Iedzimta ģenēze.",
                "Crossing lower-pole vessel": "Aberrants nieres apakšpola asinsvads.",
                "Secondary (after stones or surgery)": "Sekundāra ģenēze."}
PYELO_APPROACH = {"Laparoscopic": ("laparoskopiska", "laparoskopiskas"),
                  "Robot-assisted": ("robotasistēta", "robotasistētas"),
                  "Open": ("atvērta", "atvērtas")}
PYELO_TECHNIQUE = {"Dismembered (Anderson–Hynes)": "pēc Anderson–Hynes", "Non-dismembered (Y-V plasty)": "Y-V plastika"}


def n13_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    right = v["PUJ side"] == "Right"
    side_cap, side = ("Labās", "labās") if right else ("Kreisās", "kreisās")
    adj, adj_gen = PYELO_APPROACH[v["Pyeloplasty approach"]]
    technique = PYELO_TECHNIQUE[v["Pyeloplasty technique"]]
    day = fmt_date(v["Procedure date"])
    admission = "akūtā kārtā" if v["Admission"] == "Acute" else "plānveida kārtā"
    intraop, postop = c61_courses(v)
    stent = v["Pyeloplasty stent"] == "DJ stent"
    vessel = v.get("PUJ cause") == "Crossing lower-pole vessel" and v.get("Vessel transposition")
    split = v.get("Split renal function (%)")

    # Clinical diagnosis
    diagnosis = f"{side_cap} puses pieloureterālā segmenta striktūra, hidronefroze."
    if v.get("PUJ cause") in PUJ_CAUSE_DX:
        diagnosis += " " + PUJ_CAUSE_DX[v["PUJ cause"]]
    extras = (["aberrantā asinsvada transpozīciju"] if vessel else []) + (["DJ stenta ievietošanu"] if stent else [])
    diagnosis += (f"\nSt. pēc {side} puses {adj_gen} pieloplastikas ({technique})"
                  + (" ar " + lv_list(extras) if extras else "") + f" ({day}).")

    # Epicrisis
    reason = f"iepriekš diagnosticētu {side} puses pieloureterālā segmenta striktūru ar hidronefrozi"
    if split not in (None, ""):
        reason += f" ({side} nieres relatīvā funkcija {split:g} %)"
    drain = v.get("Pyeloplasty drain") or "No drain"
    course = [f"Pēcoperācijas periods {postop}."]
    if drain.startswith("Removed on day"):
        course.append(f"Drenu evakuē {drain.split()[-1]}. pēcoperācijas dienā.")
    elif drain == "Discharged with drain":
        course.append("Izrakstīšanas brīdī drena saglabāta.")
    course.append(c61_catheter_sentence(f, v))
    if v["Histology"] == "Pending":
        histology = "Histoloģiskās izmeklēšanas rezultāts izrakstīšanas brīdī vēl nav zināms."
    else:
        histology = f"Histoloģiskās izmeklēšanas rezultāts: {v['PUJ histology result'].strip().rstrip('.')}."
    epicrisis = "\n".join([
        f"{f['Pacients']} {f['stacionets']} P. Stradiņa KUS 4. uroloģijas nodaļā {admission} operatīvai ārstēšanai "
        f"sakarā ar {reason}. {f['Pacients']} tiek {f['gatavots']} operatīvai terapijai.",
        f"{day} vispārējā anestēzijā veikta {side} puses {adj} pieloplastika ({technique})"
        + (" ar " + lv_list(extras) if extras else "") + f". Operācija noritēja {intraop}.",
        " ".join(course),
        f"{histology} Uzlabojoties vispārējam stāvoklim, {f['pacients']} {f['izrakstits']} turpmākai aprūpei "
        "ģimenes ārsta uzraudzībā. Sniegtas rekomendācijas.",
    ])

    # Recommendations
    recs = []
    if v["Discharge paper"] == "Temporary":
        recs.append(TEMPORARY_PAPER_TEXT)
    recs += [
        "Atrasties ģimenes ārsta un ārstējošā urologa uzraudzībā.",
        "Diēta: vispusīga. Lietot šķidrumu 1,5–2,0 L diennaktī, ja nav medicīnisku kontrindikāciju.",
        f"Režīms: saudzējošs {WEEKS_TEXT[v['Activity restriction']]} no operācijas dienas – necelt smagumus, "
        "izvairīties no pārmērīgas fiziskas slodzes, pirts un baseina apmeklējuma.",
    ]
    if stent:
        when = (fmt_date(v["Stent removal date (N13)"]) if v.get("Stent removal plan") == "On a specific date"
                else "4–6 nedēļas pēc operācijas")
        recs += [
            f"DJ stenta izņemšana (cistoskopiski) plānota {when}; neatlikt – ilgstoši atstāts stents var inkrustēties "
            "un izraisīt infekciju.",
            "Ar DJ stentu iespējama bieža un neatliekama urinācija, diskomforts vai sāpes sānos urinējot un neliels "
            "asiņu piejaukums urīnā – šīs sūdzības nav bīstamas. Regulāri iztukšot urīnpūsli, nenoturēt urīnu ilgstoši.",
        ]
    if v["Catheter removal"] == "Discharged with catheter":
        if v.get("Planned catheter removal (optional)"):
            recs.append("Kopt urīnpūšļa ilgkatetru; " + CATHETER_CARE_LINES[1][0].lower() + CATHETER_CARE_LINES[1][1:])
            recs.append(f"Katetra evakuācija plānota {fmt_date(v['Planned catheter removal (optional)'])}.")
        else:
            recs += CATHETER_CARE_LINES
    if drain == "Discharged with drain":
        recs.append("Drenas evakuācija pēc ārstējošā urologa norādījumiem; sekot izdalījumu daudzumam.")
    recs.append("Ambulatori turpināt pēcoperācijas brūču apdari, šuvju vietas apstrādājot ar Betadine 100 mg/ml vai "
                "brilijantzaļo šķīdumu, pārsēja maiņa pēc nepieciešamības.")
    if v.get("Skin sutures") in WOUND_SUTURE_TEXT:
        recs.append(WOUND_SUTURE_TEXT[v["Skin sutures"]])
    if blocks:
        recs.append("Terapijā lietot:")
        recs += ["• " + b for b in blocks]
    visit = f"Vizīte pie urologa ambulatori: {WEEKS_TEXT[v['Urology visit']]} pēc operācijas"
    if v["Histology"] == "Pending":
        visit += ", lai izvērtētu histoloģisko atbildi"
    recs.append(visit + ".")
    recs += tests_lines(v, "plain", after_op=True)  # e.g. ultrasound after stent removal, diuretic renography
    recs.append("Ja paaugstinās ķermeņa temperatūra, parādās drebuļi, stipras sāpes sānos, nespēja urinēt, izteikta "
                "hematūrija vai šķidruma tecēšana no brūces, nekavējoties vērsties NMC.")
    return {"clinical_diagnosis": diagnosis, "epicrisis": epicrisis, "recommendations": "\n".join(recs)}


def n13_checks(v):
    blockers, cautions = [], []
    if v.get("Voiding after catheter removal") == "Urinary retention":
        blockers.append("Urinary retention after catheter removal.")
    if v.get("Histology") == "Known" and blank(v.get("PUJ histology result")):
        blockers.append("Type the histology result.")
    if v.get("Pyeloplasty stent") == "DJ stent" and v.get("Stent removal plan") != "On a specific date":
        cautions.append("A DJ stent stays in: if possible give a concrete removal date, so the stent is not forgotten.")
    if v.get("Pyeloplasty drain") == "Discharged with drain":
        cautions.append("Discharged with a drain: high output can mean an anastomotic urine leak – check before discharge.")
    if v.get("Histology") == "Pending" and v.get("Discharge paper") == "Final":
        cautions.append("Histology is still pending, but the discharge paper is marked as final.")
    if v.get("Histology") == "Known" and v.get("Discharge paper") == "Temporary":
        cautions.append("Histology is already known, but the discharge paper is marked as temporary.")
    split = v.get("Split renal function (%)")
    if split not in (None, "") and split < 15:
        cautions.append("Split renal function below 15 %: recovery after pyeloplasty is less likely; make sure "
                        "nephrectomy was considered and the follow-up renogram is planned.")
    if "With complications" in (v.get("Intraoperative course"), v.get("Postoperative course")):
        cautions.append("Complications are recorded: check that the recommendations cover them.")
    return blockers, cautions


BUILDERS["n13"] = n13_parts
CHECKS["N13.0"] = n13_checks


# ---------------------------------------------------------------
# 2p. SCROTAL MASS (D29.4): excision in 4. Urology
# ---------------------------------------------------------------
# Scrotal wall lesions (epidermoid cysts, lipomas, calcinosis) are nearly always benign; every scrotal mass needs an
# ultrasound first (EAU testicular cancer guideline): a solid INTRAtesticular mass means inguinal orchidectomy, and a
# solid paratesticular / spermatic cord mass can be a sarcoma (MRI, inguinal approach, wide excision).

SCROTAL_PLACE = {  # (diagnosis, after 'sakarā ar ...' / in the operation name, genitive)
    "Scrotal wall (subcutaneous)": ("Scrotum zemādas", "sēklinieka maisiņa zemādas", "scrotum zemādas"),
    "Scrotum and perineum": ("Scrotum un starpenes zemādas", "sēklinieka maisiņa un starpenes zemādas",
                             "scrotum un starpenes zemādas"),
    "Scrotal skin": ("Scrotum ādas", "sēklinieka maisiņa ādas", "scrotum ādas"),
    "Spermatic cord / paratesticular": ("Paratestikulārs (sēklas virvītes)", "paratestikulāru (sēklas virvītes)",
                                        "paratestikulāra (sēklas virvītes)"),
}
SCROTAL_SIDE = {"Right": "labajā pusē", "Left": "kreisajā pusē", "Midline (raphe)": "viduslīnijā",
                "Both sides": "abās pusēs"}
SCROTAL_TYPE = {"Epidermoid (sebaceous) cyst": "epidermoidāla cista/ateroma", "Lipoma": "lipoma",
                "Fibroma": "fibroma", "Scrotal calcinosis": "scrotum kalcinoze"}
SCROTAL_ANESTHESIA = {"Spinal": "spinālā anestēzijā", "General": "vispārējā anestēzijā",
                      "Local": "vietējā anestēzijā"}
MALIGNANT_WORDS = ("karcinom", "sarkom", "melanom", "paget", "ļaundabīg", "malign")


def scrotal_words(v):
    """(diagnosis words, 'sakarā ar ...' words, operation name) for the mass(es)."""
    dx_place, acc_place, gen_place = SCROTAL_PLACE[v["Scrotal mass location"]]
    many = v.get("Scrotal mass number") == "Multiple"
    if v["Scrotal mass location"] == "Spermatic cord / paratesticular":  # the adjective agrees with the noun
        dx = dx_place + (" veidojumi" if many else " veidojums")
        acc = acc_place.replace("paratestikulāru", "paratestikulārus" if many else "paratestikulāru") \
            + (" veidojumus" if many else " veidojumu")
        name = gen_place.replace("paratestikulāra", "paratestikulāru" if many else "paratestikulāra") \
            + (" veidojumu ekscīzija" if many else " veidojuma ekscīzija")
    else:
        dx = dx_place + (" veidojumi" if many else " veidojums")
        acc = acc_place + (" veidojumus" if many else " veidojumu")
        name = gen_place + (" veidojumu ekscīzija" if many else " veidojuma ekscīzija")
    return dx, acc, name


def scrotal_parts(sex, v, blocks=()):
    f = SEX_FORMS["Man"]  # scrotum: always male wording
    dx_words, acc_words, op_name = scrotal_words(v)
    day = fmt_date(v["Procedure date"])
    admission = "akūtā kārtā" if v["Admission"] == "Acute" else "plānveida kārtā"
    intraop, postop = c61_courses(v)

    details = [SCROTAL_SIDE[v["Scrotal mass side"]]]
    size = v.get("Scrotal mass size (cm)")
    if size not in (None, ""):
        details.append(f"Ø {size:g} cm".replace(".", ","))
    kind = v.get("Scrotal mass type")
    clinical = ""
    if kind in SCROTAL_TYPE:
        clinical = "; klīniski – " + SCROTAL_TYPE[kind]
    elif kind == "Other (type below)" and not blank(v.get("Scrotal mass type (other)")):
        clinical = "; klīniski – " + v["Scrotal mass type (other)"].strip().rstrip(".")

    # Clinical diagnosis
    diagnosis = f"{dx_words} ({', '.join(details)}{clinical}), st. pēc {op_name.replace('ekscīzija', 'ekscīzijas')} ({day})."
    if v["Histology"] == "Known":
        diagnosis += f" Histoloģiski: {v['Scrotal histology result'].strip().rstrip('.')}."
    if not blank(v.get("Other diagnoses (optional)")):
        diagnosis += "\n" + v["Other diagnoses (optional)"].strip().rstrip(".") + "."

    # Epicrisis
    reason = acc_words + f" ({SCROTAL_SIDE[v['Scrotal mass side']]})"
    if v.get("Scrotal ultrasound before surgery"):
        reason += "; USS: ekstratestikulārs veidojums, sēklinieki bez patoloģiskas atrades"
    sentences = [
        f"{f['Pacients']} {f['stacionets']} P. Stradiņa KUS 4. uroloģijas nodaļā {admission} operatīvai terapijai "
        f"sakarā ar {reason}.",
        f"{f['Pacients']} tiek {f['gatavots']} operācijai, {day} {SCROTAL_ANESTHESIA[v['Scrotal anesthesia']]} tiek "
        f"veikta {op_name}. Operācija noritēja {intraop}.",
        f"Pēcoperācijas periods {postop}."
        + (" Gumijas strēmele evakuēta 1. pēcoperācijas dienā." if v.get("Scrotal drain") else "")
        + (" Novērojama neliela sēklinieku maisiņa hematoma, dinamikā bez palielināšanās." if v.get("Scrotal hematoma")
           else ""),
    ]
    if v["Histology"] == "Pending":
        expected = v.get("Histology expected (date)")
        sentences.append("Patohistoloģiskās izmeklēšanas slēdziens būs pieejams"
                         + (f" no {fmt_date(expected)}." if expected else " pēc 2–3 nedēļām."))
    else:
        sentences.append(f"Patohistoloģiskās izmeklēšanas slēdziens: {v['Scrotal histology result'].strip().rstrip('.')}.")
    sentences.append(f"Uzlabojoties vispārējam stāvoklim, {f['pacients']} {f['izrakstits']} turpmākai aprūpei ģimenes "
                     "ārsta uzraudzībā, sniegtas rekomendācijas.")
    epicrisis = (" ".join(sentences) + "\n\nIzmeklējumi: veikti ambulatori, pieļauj operāciju.\n\n"
                 f"Operācijas: {op_name} ({day}).")

    # Recommendations
    recs = []
    if v["Discharge paper"] == "Temporary":
        recs.append("• " + TEMPORARY_PAPER_TEXT)
    recs += [
        "• Atrasties ģimenes ārsta uzraudzībā.",
        f"• Režīms – saudzējošs {WEEKS_TEXT[v['Activity restriction']]} no operācijas dienas, atturēties no "
        "pārmērīgas fiziskas slodzes, smagumu celšanas un riteņbraukšanas; pirts vai baseina apmeklējuma – 1 mēnesi "
        "no operācijas dienas.",
        "• Diēta – vispusīga, lietot šķidrumu ap 2,0 L/dnn.",
        "• Pirmās 1–2 nedēļas valkāt cieši pieguļošu apakšveļu vai suspensoriju; pirmajās 48 stundās sēklinieku "
        "maisiņam vairākas reizes dienā uz 15–20 minūtēm likt aukstu kompresi (caur drānu).",
        "• Atturēties no dzimumdzīves līdz brūces sadzīšanai (ap 2 nedēļām).",
        "• Ambulatori turpināt pēcoperācijas brūces apdari, šuves vietu apstrādājot ar Betadine 100 mg/ml vai "
        "brilijantzaļo šķīdumu, pārsēja maiņa pēc nepieciešamības; dušā drīkst mazgāties 48 stundas pēc operācijas.",
    ]
    if v.get("Skin sutures") == "Absorbable":
        recs.append("• Diegu evakuācija nav nepieciešama (uzsūcošās šuves).")
    else:
        recs.append("• Diegu evakuācija 10.–14. pēcoperācijas dienā ģimenes ārsta praksē.")
    if blocks:
        recs.append("• Nozīmē lietot:")
        recs += ["    o " + b for b in blocks]
    if v["Histology"] == "Pending":
        recs.append("• Patohistoloģiskās izmeklēšanas rezultātu izvērtēs ārstējošais urologs.")
    recs.append(f"• Kontrole pie urologa ambulatori {WEEKS_TEXT[v['Urology visit']]} pēc operācijas.")
    recs += tests_lines(v, "bullet", after_op=True)
    recs.append("• Ja sēklinieku maisiņš strauji palielinās vai kļūst sasprindzināts, pastiprinās sāpes, brūce asiņo vai "
                "izdalās strutas, paaugstinās ķermeņa temperatūra, nekavējoties vērsties NMC.")
    return {"clinical_diagnosis": diagnosis, "epicrisis": epicrisis, "recommendations": "\n".join(recs)}


def scrotal_checks(v):
    blockers, cautions = [], []
    if v.get("Histology") == "Known" and blank(v.get("Scrotal histology result")):
        blockers.append("Type the histology result.")
    if not v.get("Scrotal ultrasound before surgery"):
        cautions.append("No scrotal ultrasound before surgery: EAU – image every scrotal mass. A solid intratesticular "
                        "mass needs an inguinal orchidectomy, not a scrotal excision.")
    if v.get("Scrotal mass location") == "Spermatic cord / paratesticular" and v.get("Scrotal mass type") in (
            "Not specified", "Fibroma", "Other (type below)"):
        cautions.append("Solid paratesticular / spermatic cord mass: most are benign, but sarcoma (liposarcoma) is "
                        "possible – MRI and an inguinal approach with wide excision are preferred when malignancy is "
                        "suspected.")
    result = (v.get("Scrotal histology result") or "").lower()
    if v.get("Histology") == "Known" and any(w in result for w in MALIGNANT_WORDS):
        cautions.append("Malignant histology: check the margins, plan staging and the oncology council; a re-excision "
                        "with wider margins may be needed.")
    if "amoxiclav-10" in (v.get("_blocks") or []) and not v.get("Inflamed / infected lesion"):
        cautions.append("Amoksiklav after clean scrotal surgery: antibiotics are rarely needed without infection.")
    if v.get("Histology") == "Pending" and v.get("Discharge paper") == "Final":
        cautions.append("Histology is still pending, but the discharge paper is marked as final.")
    if v.get("Histology") == "Known" and v.get("Discharge paper") == "Temporary":
        cautions.append("Histology is already known, but the discharge paper is marked as temporary.")
    if v.get("Scrotal hematoma"):
        cautions.append("Scrotal hematoma: make sure it is not growing before discharge.")
    if "With complications" in (v.get("Intraoperative course"), v.get("Postoperative course")):
        cautions.append("Complications are recorded: check that the recommendations cover them.")
    return blockers, cautions


BUILDERS["scrotal"] = scrotal_parts
CHECKS["D29.4"] = scrotal_checks


# ---------------------------------------------------------------
# 2q. BLADDER STONE (N21.0): cystolithotripsy in 4. Urology
# ---------------------------------------------------------------
# EAU urolithiasis – bladder stones: transurethral cystolithotripsy where possible (strong), percutaneous where
# transurethral is not possible (strong), open for very large stones (weak); laser, pneumatic and mechanical are
# equivalent; look for the cause in every patient (flow rate, residual urine, dipstick/culture, stone analysis).

CYSTOLITHO = {  # (as done, after "St. pēc ...")
    "Transurethral – laser (Ho:YAG)": ("cistoskopija, kontakta lāzera cistolitotripsija",
                                       "cistoskopijas, kontakta lāzera cistolitotripsijas"),
    "Transurethral – pneumatic": ("cistoskopija, pneimatiskā cistolitotripsija",
                                  "cistoskopijas, pneimatiskās cistolitotripsijas"),
    "Transurethral – mechanical lithotriptor": ("cistoskopija, mehāniska cistolitotripsija",
                                                "cistoskopijas, mehāniskas cistolitotripsijas"),
    "Percutaneous cystolithotripsy": ("perkutāna cistolitotripsija", "perkutānas cistolitotripsijas"),
    "Open cystolithotomy": ("atvērta cistolitotomija", "atvērtas cistolitotomijas"),
}
N21_ALSO = {"TURP": "TURP", "Optical urethrotomy": "optiskā uretrotomija", "Foreign body removal": "svešķermeņa evakuācija"}
N21_CAUSE = {
    "BPH / bladder outlet obstruction": "LPH ar infravezikālu obstrukciju",
    "Urethral stricture": "uretras striktūra",
    "Neurogenic bladder": "neirogēns urīnpūslis",
    "Long-term catheter": "ilgstošs urīnpūšļa katetrs",
    "Foreign body / clips or sutures after surgery": "svešķermenis (klipši/šuves pēc operācijas)",
    "Augmented bladder / reservoir": "augmentēts urīnpūslis / rezervuārs",
}
N21_FOUND = {"Ultrasound": "ultrasonogrāfiski", "CT": "datortomogrāfiski", "KUB X-ray": "rentgenoloģiski",
             "Cystoscopy": "cistoskopiski"}
N21_ANESTHESIA = {"Spinal": "spinālā anestēzijā", "General": "vispārējā anestēzijā"}


def n21_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    many = v["Bladder stone number"] == "Multiple"
    size = v.get("Bladder stone size (cm)")
    size_text = ""
    if size not in (None, ""):
        size_text = (f" (lielākais Ø {size:g} cm)" if many else f" (Ø {size:g} cm)").replace(".", ",")
    op_nom, op_gen = CYSTOLITHO[v["Cystolithotripsy method"]]
    also = [N21_ALSO[x] for x in (v.get("Done at the same time (N21)") or []) if x in N21_ALSO]
    day = fmt_date(v["Procedure date"])
    admission = "akūtā kārtā" if v["Admission"] == "Acute" else "plānveida kārtā"
    intraop, postop = c61_courses(v)
    causes = [N21_CAUSE[c] for c in (v.get("Bladder stone cause") or []) if c in N21_CAUSE]
    culture = v.get("Urine culture (N21)")

    # Clinical diagnosis
    diagnosis = ("Urīnpūšļa konkrementi" if many else "Urīnpūšļa konkrements") + size_text + "."
    if culture == "Positive":
        diagnosis += " Urīnceļu infekcija."
    diagnosis += f"\nSt. pēc {op_gen}" + (" un " + lv_list(also) if also else "") + f" ({day})."
    if causes:
        diagnosis += "\nKonkrementu veidošanās iemesls – " + lv_list(causes) + "."
    if not blank(v.get("Other diagnoses (optional)")):
        diagnosis += "\n" + v["Other diagnoses (optional)"].strip().rstrip(".") + "."

    # Epicrisis
    found = [N21_FOUND[x] for x in (v.get("Bladder stone found by") or []) if x in N21_FOUND]
    stone_acc = "urīnpūšļa konkrementiem" if many else "urīnpūšļa konkrementu"
    reason = (f"{lv_list(found)} verificētu{'iem' if many else ''} {stone_acc}" if found else stone_acc) + size_text
    if many and found:
        reason = f"{lv_list(found)} verificētiem {stone_acc}{size_text}"
    course = []
    if v.get("Bladder irrigation after surgery"):
        course.append("Pēcoperācijas periodā veikta urīnpūšļa skalošana.")
    course.append(f"Pēcoperācijas periods {postop}.")
    course.append(c61_catheter_sentence(f, v))
    extra = []
    if v.get("Stone sent for analysis"):
        extra.append("Konkrementa fragmenti nosūtīti ķīmiskā sastāva analīzei.")
    if culture == "Positive":
        organism = (v.get("Culture organism (N21)") or "").strip()
        extra.append("Urīna uzsējums – pozitīvs" + (f" ({organism})" if organism else "") + ".")
    elif culture == "Negative":
        extra.append("Urīna uzsējums – negatīvs.")
    elif culture == "Result pending":
        extra.append("Urīna uzsējuma rezultāts izrakstīšanas brīdī vēl nav zināms.")
    epicrisis = "\n".join(p for p in [
        f"{f['Pacients']} {f['stacionets']} P. Stradiņa KUS 4. uroloģijas nodaļā {admission} operatīvai ārstēšanai "
        f"sakarā ar {reason}. {f['Pacients']} tiek {f['gatavots']} operācijai.",
        f"{day} {N21_ANESTHESIA[v['Anesthesia (N21)']]} veikta {op_nom}" + (" un " + lv_list(also) if also else "")
        + f". Operācija noritēja {intraop}.",
        " ".join(course),
        " ".join(extra),
        f"Uzlabojoties vispārējam stāvoklim, {f['pacients']} {f['izrakstits']} turpmākai aprūpei ģimenes ārsta "
        "uzraudzībā. Sniegtas rekomendācijas.",
    ] if p)

    # Recommendations
    recs = [
        "Atrasties ģimenes ārsta un ārstējošā urologa uzraudzībā.",
        "Diēta: vispusīga. Lietot šķidrumu 2,0–2,5 L diennaktī, ja nav medicīnisku kontrindikāciju.",
        f"Režīms: saudzējošs {WEEKS_TEXT[v['Activity restriction']]} no operācijas dienas, izvairīties no pārmērīgas "
        "fiziskas slodzes un smagumu celšanas.",
        "Pirmajās dienās iespējams neliels asiņu piejaukums urīnā un dedzināšana urinējot; šādā gadījumā palielināt "
        "šķidruma uzņemšanu.",
    ]
    if v["Catheter removal"] == "Discharged with catheter":
        if v.get("Planned catheter removal (optional)"):
            recs.append("Kopt urīnpūšļa ilgkatetru; " + CATHETER_CARE_LINES[1][0].lower() + CATHETER_CARE_LINES[1][1:])
            recs.append(f"Katetra evakuācija plānota {fmt_date(v['Planned catheter removal (optional)'])}.")
        else:
            recs += CATHETER_CARE_LINES
    else:
        recs.append("Regulāri un pilnībā iztukšot urīnpūsli, nenoturēt urīnu ilgstoši.")
    cause_set = set(v.get("Bladder stone cause") or [])
    if "BPH / bladder outlet obstruction" in cause_set and "TURP" not in (v.get("Done at the same time (N21)") or []):
        recs.append("Turpināt infravezikālās obstrukcijas (LPH) ārstēšanu pie urologa – bez tās konkrementi var "
                    "veidoties atkārtoti.")
    if cause_set & {"Neurogenic bladder", "Augmented bladder / reservoir"}:
        recs.append("Regulāra urīnpūšļa skalošana ar fizioloģisko šķīdumu pēc ārstējošā urologa norādījumiem.")
    if v.get("Stone sent for analysis"):
        recs.append("Konkrementa ķīmiskās analīzes rezultātu izvērtēs ārstējošais urologs; turpmākā profilakse "
                    "atbilstoši konkrementa sastāvam.")
    if culture == "Result pending":
        recs.append("Urīna uzsējuma rezultātu pārbauda ģimenes ārsts, nepieciešamības gadījumā koriģējot terapiju.")
    if blocks:
        recs.append("Terapijā lietot:")
        recs += ["• " + b for b in blocks]
    recs.append(f"Vizīte pie urologa ambulatori: {WEEKS_TEXT[v['Urology visit']]} pēc operācijas.")
    recs += tests_lines(v, "plain", after_op=True)
    recs.append("Ja parādās izteikta hematūrija ar asins recekļiem, nav iespējams urinēt, katetrs nefunkcionē, rodas "
                "drudzis vai drebuļi, nekavējoties vērsties NMC.")
    return {"clinical_diagnosis": diagnosis, "epicrisis": epicrisis, "recommendations": "\n".join(recs)}


def n21_checks(v):
    blockers, cautions = [], []
    if v.get("Voiding after catheter removal") == "Urinary retention":
        blockers.append("Urinary retention after catheter removal.")
    causes = v.get("Bladder stone cause") or []
    if not causes:
        cautions.append("EAU (weak): look for the cause of every bladder stone – flow rate and residual urine, urine "
                        "dipstick/culture, stone analysis; choose 'Not found' if the work-up was negative.")
    if "BPH / bladder outlet obstruction" in causes and "TURP" not in (v.get("Done at the same time (N21)") or []):
        cautions.append("Bladder stone with BPO and no TURP: plan treatment of the obstruction (treating both at once "
                        "is safe; EAU notes slightly more short-term incontinence and UTI).")
    if v.get("Urine culture (N21)") == "Positive" and not any(b in (v.get("_blocks") or []) for b in
                                                              ("amoxiclav-10",)):
        cautions.append("Positive urine culture: check whether an antibiotic is needed (by susceptibility).")
    if v.get("Cystolithotripsy method") == "Open cystolithotomy":
        cautions.append("EAU: open cystolithotomy is meant for very large stones; transurethral or percutaneous "
                        "cystolithotripsy is preferred where possible.")
    if "With complications" in (v.get("Intraoperative course"), v.get("Postoperative course")):
        cautions.append("Complications are recorded: check that the recommendations cover them.")
    return blockers, cautions


BUILDERS["n21"] = n21_parts
CHECKS["N21.0"] = n21_checks


# ---------------------------------------------------------------
# 2r. ACUTE URINARY RETENTION (R33) at NMC, sent home with a catheter
# ---------------------------------------------------------------
# Drain the bladder (urethral catheter, suprapubic if that fails), record the volume, check renal function,
# start an alpha-blocker before a trial without catheter (Cochrane: success ~60% vs ~38%), look for and treat the
# cause (BPH, drugs, constipation, UTI, stricture, neurogenic). Admit with renal impairment, upper tract dilatation,
# sepsis, suspected cauda equina or a large volume with risk of post-obstructive diuresis. EAU: 5-ARIs/combination
# reduce long-term AUR risk; recurrent or refractory retention is an indication for surgery.

RETENTION_CAUSE_DX = {
    "BPH": "Labdabīga prostatas hiperplāzija.", "Prostate cancer": "Prostatas CA.",
    "Urethral stricture": "Uretras striktūra.", "Blocked indwelling catheter": "Urīnpūšļa ilgkatetra disfunkcija.",
    "Clot retention (hematuria)": "Makrohematūrija, urīnpūšļa tamponāde ar asins recekļiem.",
    "Drug-induced": "Medikamentu izraisīta urīna retence.",
    "After surgery / anaesthesia": "Pēcoperācijas urīna retence.",
    "Constipation": "Aizcietējums.", "Urinary tract infection / prostatitis": "Urīnceļu infekcija.",
    "Neurogenic bladder": "Neirogēns urīnpūslis.", "Bladder hypotonia suspected": "Urīnpūšļa atonija susp.",
}
RETENTION_HISTORY = {
    "BPH": "LPH", "Prostate cancer": "prostatas CA", "Urethral stricture": "uretras striktūra",
    "Neurogenic bladder": "neirogēns urīnpūslis", "Blocked indwelling catheter": "urīnpūšļa ilgkatetrs",
}
RETENTION_TYPE_DX = {"Acute": "Akūta urīna retence.",
                     "Acute-on-chronic (large volume)": "Hroniska urīna retence ar akūtu paasinājumu.",
                     "Chronic (overflow)": "Hroniska urīna retence ar pārplūdes nesaturēšanu."}
RETENTION_DURATION = {"Hours": "dažas stundas", "1 day": "1 dienu", "2 days": "2 dienas", "3 days": "3 dienas",
                      "More than 3 days": "ilgāk par 3 dienām"}
URINE_LOOK_R33 = {"Clear": " dzidra, dzeltena", "Cloudy": " duļķaina", "Blood-tinged": " ar asiņu piejaukumu",
                  "With clots": " ar asins recekļiem", "Not documented": ""}
RETENTION_TWOC = {
    "Trial without catheter in 2–3 days": "Pēc 2–3 dienām, turpinot alfa blokatoru, ambulatori lemt par urīnpūšļa "
                                          "katetra evakuāciju (urinācijas mēģinājums bez katetra).",
    "Trial without catheter in 1 week": "Pēc 1 nedēļas, turpinot alfa blokatoru, ambulatori lemt par urīnpūšļa "
                                        "katetra evakuāciju (urinācijas mēģinājums bez katetra).",
    "Trial without catheter in 2 weeks": "Pēc 2 nedēļām, turpinot alfa blokatoru, ambulatori lemt par urīnpūšļa "
                                         "katetra evakuāciju (urinācijas mēģinājums bez katetra).",
    "Urologist decides": "Par urīnpūšļa katetra evakuāciju lems ārstējošais urologs.",
    "Long-term catheter": "Ilgstošs urīnpūšļa katetrs; tā maiņa ik pēc 4–6 nedēļām ģimenes ārsta praksē.",
}


def r33_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    causes = v.get("Retention cause") or []
    blocked = "Blocked indwelling catheter" in causes
    drainage = v["Retention drainage"]
    fr, volume = v["Retention catheter Fr"], v.get("Drained volume (ml)")
    day = fmt_date(v["Procedure date"])

    # Clinical diagnosis
    lines = [RETENTION_CAUSE_DX[c] for c in causes if c in RETENTION_CAUSE_DX]
    lines.append(RETENTION_TYPE_DX[v["Retention type"]])
    act = {"Urethral Foley catheter": "Urīnpūšļa katetrizācija", "Urethral Tiemann (coudé) catheter":
           "Urīnpūšļa katetrizācija ar Tīmana katetru", "Catheter changed (blocked catheter)": "Urīnpūšļa katetra maiņa",
           "Suprapubic catheter (epicystostomy)": "Epicistostomija",
           "In-out catheterisation (no catheter left)": "Vienreizēja urīnpūšļa katetrizācija"}[drainage]
    lines.append(f"{act} ({day}).")
    labs = v.get("Labs at NMC")
    if labs in ("Urinary tract infection", "UTI and raised creatinine") and \
            "Urinary tract infection / prostatitis" not in causes:
        lines.append("Urīnceļu infekcija.")
    if not blank(v.get("Other diagnoses (optional)")):
        lines.append(v["Other diagnoses (optional)"].strip().rstrip(".") + ".")
    diagnosis = "\n".join(lines)

    # Epicrisis
    duration = RETENTION_DURATION[v["Retention duration"]]
    if blocked:
        complaint = f"urīnpūšļa ilgkatetra disfunkciju, urīna neizdalīšanos pa katetru {duration}"
    else:
        complaint = f"stiprām sāpēm virs simfīzes, urīna neizdalīšanos {duration}"
    sentences = [f"{f['Pacients']} {f['iestajies']} akūtā kārtā PSKUS NMC ar sūdzībām par {complaint}."]
    history = [RETENTION_HISTORY[c] for c in causes if c in RETENTION_HISTORY and c != "Blocked indwelling catheter"]
    if history:
        sentences.append(f"Anamnēzē – {lv_list(history)}.")
    if v.get("Previous retention episodes"):
        sentences.append("Anamnēzē atkārtotas urīna retences epizodes.")
    if v.get("Ultrasound at NMC"):
        us = f"NMC etapā vispusīgi {f['izmekleta']}, veikta USS vēdera dobumam"
        found = []
        if v.get("Enlarged prostate on US"):
            found.append("palielinātu prostatu")
        found.append("pildītu urīnpūsli")
        us += ", kur vizualizē " + lv_list(found)
        us += "; augšējo urīnceļu dilatācija" if v.get("Upper tract dilatation") == "Present" else \
              "; augšējo urīnceļu dilatāciju nekonstatē"
        sentences.append(us + ".")
    else:
        sentences.append(f"NMC etapā vispusīgi {f['izmekleta']}.")
    look = URINE_LOOK_R33.get(v.get("Urine appearance"), "")
    drained = f"izdalās ap {volume:g} ml{look} urīna" if volume not in (None, "") else "izdalās urīns"
    if drainage == "Urethral Foley catheter":
        sentences.append(f"Sterilos apstākļos veic urīnpūšļa intubāciju ar Foley {fr} Fr katetru, {drained}.")
    elif drainage == "Urethral Tiemann (coudé) catheter":
        sentences.append(f"Foley katetru ievadīt neizdodas; sterilos apstākļos ievada Tīmana {fr} Fr katetru, {drained}.")
    elif drainage == "Catheter changed (blocked catheter)":
        sentences.append(f"Sterilos apstākļos veic urīnpūšļa katetra maiņu, ievada Foley {fr} Fr katetru, {drained}.")
    elif drainage == "Suprapubic catheter (epicystostomy)":
        sentences.append("Uretrālu katetru ievadīt neizdodas; vietējā anestēzijā USS kontrolē veic perkutānu "
                         f"epicistostomiju, ievada Ch {fr} epicistostomas katetru, {drained}.")
    else:
        sentences.append(f"Sterilos apstākļos veic vienreizēju urīnpūšļa katetrizāciju ar {fr} Fr katetru, {drained}; "
                         "katetru evakuē.")
    sentences.append("Sūdzības par spiedošo sajūtu pāriet.")
    lab_text = {"No inflammatory changes": "Laboratoriski bez iekaisuma izmaiņām.",
                "Urinary tract infection": "Laboratoriski – urīnceļu infekcija, paņemts urīna uzsējums.",
                "Raised creatinine": "Laboratoriski – paaugstināts kreatinīna līmenis.",
                "UTI and raised creatinine": "Laboratoriski – urīnceļu infekcija, paaugstināts kreatinīna līmenis; "
                                             "paņemts urīna uzsējums."}.get(labs)
    if lab_text:
        sentences.append(lab_text)
    if not blank(v.get("Other NMC findings (optional)")):
        extra = v["Other NMC findings (optional)"].strip()
        sentences.append(extra if extra.endswith(".") else extra + ".")
    sentences.append("Dotajā brīdī nav indikāciju neatliekamai stacionēšanai. Vispārēji apmierinošā stāvoklī un ar "
                     "rekomendācijām izraksta tālākai ambulatorai ārstēšanai.")

    # Recommendations
    numbered = make_counter()
    recs = [numbered("Atrasties ģimenes ārsta uzraudzībā."),
            numbered("Lietot šķidrumu 1,5–2 l/DN, ja nav medicīnisku kontrindikāciju; izvairīties no alkohola un liela "
                     "šķidruma daudzuma īsā laikā.")]
    if drainage == "Suprapubic catheter (epicystostomy)":
        recs.append(numbered("Kopt epicistostomas katetru: stomas vietu ikdienā apstrādāt un pārsiet; katetra maiņa ik "
                             "pēc 4–6 nedēļām ģimenes ārsta praksē, urīna savācējmaisa maiņa ik pēc 3–7 dienām."))
    elif drainage == "In-out catheterisation (no catheter left)":
        recs.append(numbered("Sekot urinācijai; ja atkārtoti nav iespējams urinēt vai urinē mazām porcijām ar "
                             "pilnuma sajūtu, nekavējoties vērsties NMC."))
    else:
        recs.append(numbered("Kopt urīnpūšļa katetru – uzturēt noslēgtu drenāžas sistēmu, maisu novietot zem "
                             "urīnpūšļa līmeņa. Urīna savācējmaisa maiņa ik pēc 3–7 dienām vai biežāk, ja tas ir bojāts."))
    if blocks:
        recs.append(numbered("Terapijā lietot:"))
        recs += ["    " + b for b in blocks]
    plan = v.get("Catheter plan (R33)")
    if plan in RETENTION_TWOC and drainage != "In-out catheterisation (no catheter left)":
        text = RETENTION_TWOC[plan]
        if drainage == "Suprapubic catheter (epicystostomy)":
            text = text.replace("urīnpūšļa katetra evakuāciju", "epicistostomas katetra noslēgšanu/evakuāciju") \
                       .replace("urīnpūšļa katetrs", "epicistostomas katetrs")
        recs.append(numbered(text))
    recs.append(numbered(f"Ambulatori urologa konsultācija {CONSULT_TEXT[v['Urology consultation']]}, lai izvērtētu "
                         "urīna retences cēloni un tālāko ārstēšanu."))
    recs += tests_lines(v, "numbered", numbered)
    recs.append(numbered("Izvairīties no medikamentiem, kas var veicināt urīna retenci (antiholīnerģiskie līdzekļi, "
                         "pirmās paaudzes antihistamīni, dekongestanti, opioīdi), un no aizcietējumiem."))
    recs.append(numbered("Ja katetrs nefunkcionē vai nav urīna atteces, parādās drudzis, drebuļi, izteikta hematūrija ar "
                         "recekļiem vai jauni neiroloģiski simptomi (kāju vājums, jušanas traucējumi starpenē), "
                         "nekavējoties vērsties NMC."))
    return {"clinical_diagnosis": diagnosis, "epicrisis": " ".join(sentences), "recommendations": "\n".join(recs)}


def r33_checks(v):
    blockers, cautions = [], []
    causes = v.get("Retention cause") or []
    chosen = v.get("_blocks") or []
    if v.get("_sex") == "Woman" and any(c in causes for c in ("BPH", "Prostate cancer")):
        blockers.append("BPH / prostate cancer is chosen as the cause for a woman.")
    if v.get("Fever at NMC"):
        blockers.append("Fever with urinary retention: assess for urosepsis (admission and IV antibiotics) before "
                        "sending the patient home.")
    volume = v.get("Drained volume (ml)") or 0
    if (volume > 1000 or v.get("Labs at NMC") in ("Raised creatinine", "UTI and raised creatinine")
            or v.get("Upper tract dilatation") == "Present" or v.get("Retention type") != "Acute"):
        cautions.append("Large volume, raised creatinine, upper tract dilatation or chronic retention: risk of "
                        "post-obstructive diuresis and renal impairment – consider admission; a trial without catheter "
                        "is likely to fail (chronic retention).")
    if v.get("Back pain or neurological symptoms"):
        cautions.append("Retention with back pain or neurological symptoms: exclude cauda equina syndrome (urgent MRI) "
                        "before discharge.")
    if "Clot retention (hematuria)" in causes:
        cautions.append("Clot retention: 3-way catheter and bladder irrigation; the hematuria needs a work-up "
                        "(cystoscopy, CT urography).")
    if "Drug-induced" in causes:
        cautions.append("Drug-induced retention: name and stop / replace the drug in the text (anticholinergics, "
                        "opioids, antihistamines, decongestants, tricyclics).")
    alpha = any(b in chosen for b in ("tamsulosin-twoc", "alfuzosin", "urorec", "dutasteride-tamsulosin"))
    if ("BPH" in causes and str(v.get("Catheter plan (R33)", "")).startswith("Trial without catheter")
            and not alpha):
        cautions.append("BPH with a planned trial without catheter: start an alpha-blocker at least 2–3 days before "
                        "the trial (improves success).")
    if v.get("Previous retention episodes") and "BPH" in causes:
        cautions.append("Recurrent retention with BPH: EAU – an absolute indication for surgery (e.g. TURP); refer.")
    uti = v.get("Labs at NMC") in ("Urinary tract infection", "UTI and raised creatinine") or \
        "Urinary tract infection / prostatitis" in causes
    if uti and not any(b in chosen for b in ("amoxiclav-7", "cefuroxime-7", "cipro")):
        cautions.append("UTI with retention: choose an antibiotic (and adjust it to the urine culture).")
    if v.get("Retention drainage") == "In-out catheterisation (no catheter left)" and volume > 800:
        cautions.append("Large volume with in-out catheterisation: an indwelling catheter is usually safer.")
    return blockers, cautions


BUILDERS["r33"] = r33_parts
CHECKS["R33"] = r33_checks


# ---------------------------------------------------------------
# 2s. HEMATURIA (R31): NMC consultation (sent home) or admission to 4. Urology
# ---------------------------------------------------------------
# EAU NMIBC (strong): history, urological examination, kidney and bladder ultrasound and/or CT urography, and
# cystoscopy (flexible, outpatient) in every patient with symptoms suggesting bladder cancer; cytology as an
# adjunct. Haematuria on anticoagulants or antiplatelets gets the same work-up (AUA, strong). BAUS: clot retention
# needs admission – large 3-way catheter (20–24 Fr), manual washout first, then continuous irrigation; theatre
# (cystoscopy, clot evacuation, coagulation) if the catheter keeps blocking or Hb keeps falling. Refractory
# (radiation) hemorrhagic cystitis: intravesical alum / formalin, hyperbaric oxygen, selective embolisation.
# Latvia (NVD): unexplained visible haematuria at >= 45 years, or haematuria that persists after UTI treatment, is a
# green corridor criterion – the GP refers with Z03.167 (bladder) for the urologist and cystoscopy.

HEM_CAUSE_DX = {
    "Bladder cancer (known)": "Ca vesicae urinariae.", "Prostate cancer": "Ca prostatae.",
    "Radiation cystitis": "Pēc staru terapijas hemorāģisks cistīts.",
    "BPH (prostatic bleeding)": "Labdabīga prostatas hiperplāzija.", "Urinary stone": "Nierakmeņu slimība.",
    "Urinary tract infection (hemorrhagic cystitis)": "Akūts hemorāģisks cistīts.",
}
HEM_HISTORY = {"Bladder cancer (known)": "urīnpūšļa CA", "Prostate cancer": "prostatas CA",
               "Radiation cystitis": "staru terapija mazajam iegurnim", "BPH (prostatic bleeding)": "LPH",
               "Urinary stone": "nierakmeņu slimība",
               "After a urological procedure": "nesen veikta uroloģiska manipulācija"}
HEM_RISK = {"Smoker or ex-smoker": "smēķēšana",
            "Occupational exposure (aromatic amines, dyes)": "profesionāla saskare ar aromātiskajiem amīniem un "
                                                             "krāsvielām",
            "Pelvic radiotherapy in the past": "staru terapija mazajam iegurnim",
            "Cyclophosphamide in the past": "ciklofosfamīda terapija"}
HEM_DURATION = {"Since today": "kopš šodienas", "1–2 days": "1–2 dienas", "3–7 days": "3–7 dienas",
                "More than a week": "ilgāk par nedēļu"}
HEM_SYMPTOMS = {"Clots in urine": "asins recekļiem urīnā", "Suprapubic pain": "sāpēm virs simfīzes",
                "Flank pain": "sāpēm jostas rajonā", "Dysuria / frequency": "sāpīgu un biežu urinēšanu"}
HEM_US = {"Clots in the bladder": "asins recekļus urīnpūslī", "Bladder wall mass": "urīnpūšļa sienas veidojumu",
          "Enlarged prostate": "palielinātu prostatu", "Kidney mass": "nieres veidojumu",
          "Kidney or ureteral stone": "konkrementu", "Hydronephrosis": "hidronefrozi", "Residual urine": "atlieku urīnu"}
HEM_CT = {"CT urography – no source found": "Veikta CT urogrāfija – hematūrijas avotu nekonstatē.",
          "CT – bladder mass": "Veikta CT ar i/v kontrastvielu – urīnpūslī vizualizē sienas veidojumu.",
          "CT – kidney / upper tract mass": "Veikta CT ar i/v kontrastvielu – vizualizē nieres / augšējo urīnceļu "
                                            "veidojumu.",
          "CT – stone": "Veikta CT – vizualizē konkrementu."}
HEM_COLOUR = {"Clear": "dzidrs", "Light pink": "gaiši rozā", "Pink, no clots": "rozā, bez recekļiem",
              "Red / with clots": "sarkans, ar recekļiem"}
ANTITHROMBOTIC_LV = {  # (accusative, after "lieto"; genitive)
    "Aspirin": ("acetilsalicilskābi", "acetilsalicilskābes"), "Clopidogrel": ("klopidogrelu", "klopidogrela"),
    "Ticagrelor": ("tikagreloru", "tikagrelora"), "Warfarin": ("varfarīnu", "varfarīna"),
    "Apixaban": ("apiksabānu", "apiksabāna"), "Rivaroxaban": ("rivaroksabānu", "rivaroksabāna"),
    "Dabigatran": ("dabigatrānu", "dabigatrāna"), "Edoxaban": ("edoksabānu", "edoksabāna"),
    "LMWH (Fraxiparine / Clexane)": ("mazmolekulāro heparīnu", "mazmolekulārā heparīna"),
}
GREEN_CORRIDOR = {
    "Bladder – Z03.167": "Vērsties pie ģimenes ārsta pēc nosūtījuma uz urologa konsultāciju un cistoskopiju "
                         "\"zaļā koridora\" ietvaros (diagnozes kods Z03.167 – izmeklēšana iespējama urīnpūšļa "
                         "ļaundabīga audzēja dēļ).",
    "Kidney – Z03.164": "Vērsties pie ģimenes ārsta pēc nosūtījuma uz CT ar i/v kontrastvielu un urologa konsultāciju "
                        "\"zaļā koridora\" ietvaros (diagnozes kods Z03.164 – izmeklēšana iespējama nieres (atskaitot "
                        "nieres bļodiņas) ļaundabīga audzēja dēļ).",
}
HEM_PROC_CYSTO = ("Cystoscopy, clot evacuation, coagulation", "TURU (resection / biopsy)", "TURP (prostatic bleeding)")
HEM_FINDING = {
    "No tumour, no bleeding source": "Datu par urīnpūšļa audzēju neiegūst.",
    "Diffuse bleeding (radiation cystitis)": "Urīnpūšļa gļotādā staru cistīta izmaiņas ar difūzu asiņošanu.",
    "Bleeding from the prostate": "Asiņošanas avots – prostata.",
    "Bladder tumour": "Urīnpūslī vizualizē jaunveidojumu.",
    "Blood from a ureteric orifice": "No urīnvada atveres izdalās asiņains urīns – asiņošanas avots augšējos "
                                     "urīnceļos.",
}
HEM_WARNING = ("Ja hematūrija pastiprinās, parādās asins recekļi, nav iespējams urinēt vai katetrs nefunkcionē, rodas "
               "drudzis, reibonis vai izteikts vājums, nekavējoties vērsties NMC.")
HEM_SMOKING = "Ieteicams atmest smēķēšanu – tas ir galvenais urīnpūšļa vēža riska faktors."
UPPER_TRACT_SIGNS = ("Flank pain", "Kidney mass", "Kidney or ureteral stone", "Hydronephrosis",
                     "CT – kidney / upper tract mass", "CT – stone", "Urinary stone", "Blood from a ureteric orifice")


def _num(v, key):
    value = v.get(key)
    return None if value in (None, "") else float(value)


def _comma(x):
    return f"{x:g}".replace(".", ",")


def _hem_unknown(v):
    causes = v.get("Hematuria cause") or []
    return not causes or "Not clear yet" in causes


def _hem_procs(v):
    """(done as, after 'St. pēc ...') for the cystoscopic procedures, or None."""
    procs = v.get("Hematuria procedure") or []
    main = [p.split(" ")[0] for p in procs if p in ("TURU (resection / biopsy)", "TURP (prostatic bleeding)")]
    if main:
        text = lv_list(main) + (" ar recekļu evakuāciju" if "Cystoscopy, clot evacuation, coagulation" in procs else "")
        return text, text
    if "Cystoscopy, clot evacuation, coagulation" in procs:
        return ("cistoskopija, recekļu evakuācija un asiņojošo vietu koagulācija",
                "cistoskopijas, recekļu evakuācijas un koagulācijas")
    return None


def _hem_upper_tract(v):
    found = set(v.get("Hematuria symptoms") or []) | set(v.get("Hematuria cause") or [])
    if v.get("Ultrasound at NMC"):
        found |= set(v.get("Ultrasound findings (R31)") or [])
    found |= {v.get("CT at NMC (R31)"), v.get("Cystoscopy finding (R31)")}
    return any(x in found for x in UPPER_TRACT_SIGNS)


def _hem_dx_lines(v, ward):
    causes = v.get("Hematuria cause") or []
    lines = [HEM_CAUSE_DX[c] for c in causes if c in HEM_CAUSE_DX]
    us = (v.get("Ultrasound findings (R31)") or []) if v.get("Ultrasound at NMC") else []
    ct = v.get("CT at NMC (R31)")
    bladder_mass = ("Bladder wall mass" in us or ct == "CT – bladder mass"
                    or (ward and v.get("Cystoscopy finding (R31)") == "Bladder tumour"))
    kidney_mass = "Kidney mass" in us or ct == "CT – kidney / upper tract mass"
    main = "Makrohematūrija" if v["Hematuria type"].startswith("Visible") else "Mikrohematūrija"
    if _hem_unknown(v) and not (bladder_mass or kidney_mass):
        main += " neskaidras etioloģijas"
    elif "After a urological procedure" in causes:
        main += " pēc uroloģiskas manipulācijas"
    elif "Catheter trauma" in causes:
        main += " urīnpūšļa katetra traumas dēļ"
    lines.append(main + ".")
    if v.get("Clot retention (bladder tamponade)"):
        lines.append("Urīnpūšļa tamponāde ar asins recekļiem.")
    if bladder_mass and "Bladder cancer (known)" not in causes:
        lines.append("Tu vesicae urinariae suspecta.")
    if kidney_mass:
        lines.append("Tu renis suspecta.")
    if ward:
        if v.get("Posthemorrhagic anemia") or (_num(v, "Transfusion (RBC units)") or 0) > 0:
            lines.append("Posthemorāģiska anēmija.")
        procs = v.get("Hematuria procedure") or []
        names = _hem_procs(v)
        if names:
            lines.append(f"St. pēc {names[1]} ({fmt_date(v['Cystoscopy date (R31)'])}).")
        if "Selective embolisation" in procs:
            lines.append(f"St. pēc selektīvas urīnpūšļa artēriju embolizācijas ({fmt_date(v['Embolisation date'])}).")
    else:
        how = v.get("NMC catheter (R31)")
        if how and how != "No catheter – voids":
            act = "Urīnpūšļa katetra maiņa" if how == "Catheter changed (blocked)" else "Urīnpūšļa katetrizācija"
            if v.get("Bladder washout (R31)"):
                act += ", skalošana ar recekļu evakuāciju"
            lines.append(act + ".")
    if not blank(v.get("Other diagnoses (optional)")):
        lines.append(v["Other diagnoses (optional)"].strip().rstrip(".") + ".")
    return lines


def _hem_complaint(v):
    symptoms = [HEM_SYMPTOMS[x] for x in (v.get("Hematuria symptoms") or []) if x in HEM_SYMPTOMS]
    if v.get("Clot retention (bladder tamponade)"):
        symptoms.append("nespēju urinēt")
    if v["Hematuria type"].startswith("Visible"):
        duration = v["Hematuria duration"]
        first = ("atkārtotu asiņu piejaukumu urīnā" if duration == "Recurrent episodes"
                 else f"asiņu piejaukumu urīnā {HEM_DURATION[duration]}")
        return "ar sūdzībām par " + lv_list([first] + symptoms)
    text = "sakarā ar urīna analīzē konstatētu mikrohematūriju"
    return text + (", sūdzas par " + lv_list(symptoms) if symptoms else "")


def _hem_history(v):
    items = []
    for key in (v.get("Hematuria cause") or []):
        if key in HEM_HISTORY and HEM_HISTORY[key] not in items:
            items.append(HEM_HISTORY[key])
    for key in (v.get("Risk factors (R31)") or []):
        if key in HEM_RISK and HEM_RISK[key] not in items:
            items.append(HEM_RISK[key])
    out = [f"Anamnēzē – {lv_list(items)}."] if items else []
    drugs = [ANTITHROMBOTIC_LV[d][0] for d in (v.get("Antithrombotic therapy") or []) if d in ANTITHROMBOTIC_LV]
    if drugs:
        inr = _num(v, "INR (optional)")
        out.append(f"Pastāvīgi lieto {lv_list(drugs)}" + (f", INR {_comma(inr)}" if inr is not None else "") + ".")
    return out


def _hem_findings(v):
    out = []
    if v.get("Ultrasound at NMC"):
        found = [HEM_US[x] for x in (v.get("Ultrasound findings (R31)") or []) if x in HEM_US]
        out.append("Veikta USS vēdera dobumam, kur vizualizē " + lv_list(found) + "." if found
                   else "Veikta USS vēdera dobumam – hematūrijas avotu nevizualizē.")
    if v.get("CT at NMC (R31)") in HEM_CT:
        out.append(HEM_CT[v["CT at NMC (R31)"]])
    labs = []
    hb = _num(v, "Hemoglobin (g/L)")
    if hb is not None:
        labs.append(f"Hb {_comma(hb)} g/L")
    extra = {"No inflammatory changes": "bez iekaisuma izmaiņām",
             "Urinary tract infection": "urīna analīzē iekaisuma izmaiņas, paņemts urīna uzsējums",
             "Raised creatinine": "paaugstināts kreatinīna līmenis",
             "UTI and raised creatinine": "urīna analīzē iekaisuma izmaiņas, paaugstināts kreatinīna līmenis; paņemts "
                                          "urīna uzsējums"}.get(v.get("Labs at NMC"))
    if extra:
        labs.append(extra)
    if labs:
        out.append("Laboratoriski – " + ", ".join(labs) + ".")
    if not blank(v.get("Other NMC findings (optional)")):
        text = v["Other NMC findings (optional)"].strip()
        out.append(text if text.endswith(".") else text + ".")
    return out


def _hem_split_blocks(blocks):
    """(medicines, other advice lines) – advice blocks are not listed under 'Terapijā lietot'."""
    advice = [b for b in blocks if b.startswith("Atkārtotas hematūrijas gadījumā apsvērt")]
    return [b for b in blocks if b not in advice], advice


def _hem_antithrombotic_line(v):
    drugs = [d for d in (v.get("Antithrombotic therapy") or []) if d in ANTITHROMBOTIC_LV]
    if not drugs:
        return ""
    acc = lv_list([ANTITHROMBOTIC_LV[d][0] for d in drugs])
    gen = lv_list([ANTITHROMBOTIC_LV[d][1] for d in drugs])
    plan = v.get("Antithrombotic plan")
    if plan == "Continue":
        return f"Turpināt lietot {acc}."
    if plan == "Pause until the urine is clear":
        return (f"{gen[0].upper() + gen[1:]} lietošanu uz laiku pārtraukt; atsākt, kad urīns kļuvis dzidrs, "
                "saskaņojot ar ģimenes ārstu vai kardiologu.")
    return f"Par {gen} terapijas turpināšanu lemt ģimenes ārstam vai kardiologam."


def r31_nmc_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    diagnosis = "\n".join(_hem_dx_lines(v, ward=False))

    # Epicrisis
    s = [f"{f['Pacients']} {f['iestajies']} akūtā kārtā PSKUS NMC {_hem_complaint(v)}."]
    s += _hem_history(v)
    s.append(f"NMC etapā vispusīgi {f['izmekleta']}.")
    s += _hem_findings(v)
    how = v["NMC catheter (R31)"]
    colour = HEM_COLOUR[v["Urine colour at discharge"]]
    if how == "No catheter – voids":
        s.append(f"{'Novērota' if sex == 'Woman' else 'Novērots'} dinamikā, urinē patstāvīgi, urīns {colour}.")
    else:
        kind = "trīsceļu Foley" if v.get("Catheter type (R31)") == "3-way catheter" else "Foley"
        wash = ", veic manuālu urīnpūšļa skalošanu, evakuē asins recekļus" if v.get("Bladder washout (R31)") else ""
        if how == "Catheter changed (blocked)":
            text = (f"Urīnpūšļa ilgkatetrs nefunkcionē (obturēts ar asins recekļiem); sterilos apstākļos veic katetra "
                    f"maiņu, ievada {kind} {v['Catheter Fr (R31)']} Fr katetru{wash}")
        else:
            text = f"Sterilos apstākļos ievada {kind} {v['Catheter Fr (R31)']} Fr katetru{wash}"
        text += f"; {'pēc skalošanas ' if wash else ''}urīns {colour}."
        if how == "Catheter, washout, removed – voids":
            text += " Katetru evakuē, urinē patstāvīgi."
        elif how == "Discharged with a catheter":
            text += " Urīnpūšļa katetru atstāj."
        s.append(text)
    s.append("Hemodinamika stabila. Dotajā brīdī nav indikāciju neatliekamai stacionēšanai. Vispārēji apmierinošā "
             "stāvoklī un ar rekomendācijām izraksta tālākai ambulatorai ārstēšanai.")

    # Recommendations
    numbered = make_counter()
    recs = [numbered("Atrasties ģimenes ārsta uzraudzībā."),
            numbered("Lietot šķidrumu 2–2,5 l/DN, ja nav medicīnisku kontrindikāciju; līdz urīns kļūst dzidrs, "
                     "izvairīties no smagas fiziskas slodzes, smagumu celšanas un aizcietējumiem.")]
    if how == "Discharged with a catheter":
        recs.append(numbered("Kopt urīnpūšļa katetru – uzturēt noslēgtu drenāžas sistēmu, maisu novietot zem "
                             "urīnpūšļa līmeņa. Par katetra evakuāciju lemt ģimenes ārstam vai urologam, kad urīns "
                             "kļuvis dzidrs."))
    elif how == "Catheter changed (blocked)":
        recs.append(numbered("Kopt urīnpūšļa ilgkatetru; katetra maiņa ik pēc 4–6 nedēļām ģimenes ārsta praksē, urīna "
                             "savācējmaisa maiņa ik pēc 3–7 dienām."))
    line = _hem_antithrombotic_line(v)
    if line:
        recs.append(numbered(line))
    meds, advice = _hem_split_blocks(blocks)
    if meds:
        recs.append(numbered("Terapijā lietot:"))
        recs += ["    " + b for b in meds]
    recs += [numbered(b) for b in advice]
    corridor = v.get("Green corridor (GP referral)")
    if corridor in GREEN_CORRIDOR:
        recs.append(numbered(GREEN_CORRIDOR[corridor]))
    else:
        why = "lai noskaidrotu hematūrijas cēloni" if _hem_unknown(v) else "tālākās ārstēšanas taktikas izvērtēšanai"
        recs.append(numbered(f"Ambulatori urologa konsultācija {CONSULT_TEXT[v['Urology consultation (R31)']]} ar "
                             f"izmeklējumu rezultātiem, {why}."))
    recs += tests_lines(v, "numbered", numbered)
    if "Smoker or ex-smoker" in (v.get("Risk factors (R31)") or []):
        recs.append(numbered(HEM_SMOKING))
    recs.append(numbered(HEM_WARNING))
    return {"clinical_diagnosis": diagnosis, "epicrisis": " ".join(s), "recommendations": "\n".join(recs)}


def r31_ward_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    diagnosis = "\n".join(_hem_dx_lines(v, ward=True))
    procs = v.get("Hematuria procedure") or []
    names = _hem_procs(v)
    treat = v.get("Ward treatment (R31)") or []
    tamponade = v.get("Clot retention (bladder tamponade)")

    # Epicrisis
    p1 = [f"{f['Pacients']} {f['stacionets']} P. Stradiņa KUS 4. uroloģijas nodaļā akūtā kārtā {_hem_complaint(v)}."]
    p1 += _hem_history(v)
    nmc = "NMC konstatēta " + ("makrohematūrija" if v["Hematuria type"].startswith("Visible") else "mikrohematūrija")
    if tamponade:
        nmc += ", urīnpūšļa tamponāde; veikta urīnpūšļa katetrizācija, uzsākta urīnpūšļa skalošana"
    p2 = [nmc + "."] + _hem_findings(v)
    acts = []
    if "Bladder irrigation (3-way catheter)" in treat:
        acts.append("turpināta nepārtraukta urīnpūšļa skalošana pa trīsceļu katetru" if tamponade
                    else "veikta nepārtraukta urīnpūšļa skalošana pa trīsceļu katetru")
    if "Manual bladder washout" in treat:
        acts.append("veikta manuāla urīnpūšļa skalošana ar recekļu evakuāciju")
    units = int(_num(v, "Transfusion (RBC units)") or 0)
    if units:
        acts.append(f"veikta {units} ER masas transfūzija" if units % 10 == 1 and units % 100 != 11
                    else f"veiktas {units} ER masas transfūzijas")
    if "IV tranexamic acid" in treat:
        acts.append("i/v ievadīta traneksāmskābe")
    if "IV antibiotics" in treat:
        acts.append("nozīmēta i/v antibakteriālā terapija")
    if v.get("Antithrombotic therapy") and v.get("Antithrombotic plan") == "Pause until the urine is clear":
        acts.append("uz laiku pārtraukta antitrombotiskā terapija")
    p3 = ["Nodaļā " + lv_list(acts) + "."] if acts else []
    if names:
        p3.append(f"{fmt_date(v['Cystoscopy date (R31)'])} {ANESTHESIA_TEXT[v['Anesthesia (R31)']]} veikta {names[0]}.")
        if v.get("Cystoscopy finding (R31)") in HEM_FINDING:
            p3.append(HEM_FINDING[v["Cystoscopy finding (R31)"]])
    if "Alum bladder irrigation (1%)" in procs:
        p3.append("Veikta urīnpūšļa skalošana ar 1% alauna šķīdumu.")
    if "Selective embolisation" in procs:
        p3.append(f"{fmt_date(v['Embolisation date'])} veikta selektīva urīnpūšļa artēriju embolizācija.")
    histology = v.get("Histology (R31)") if names else None
    if histology == "Pending":
        p3.append("Histoloģiskās izmeklēšanas rezultāts izrakstīšanas brīdī vēl nav zināms.")
    elif histology == "Known":
        p3.append("Histoloģiskās izmeklēšanas rezultāts: " + v["Histology result (R31)"].strip().rstrip(".") + ".")
    colour = v["Urine colour at discharge"]
    p4 = [("Pēc ārstēšanas hematūrija izzūd, urīns dzidrs." if colour == "Clear"
           else f"Pēc ārstēšanas hematūrija mazinās, urīns {HEM_COLOUR[colour]}.")]
    hb = _num(v, "Hemoglobin at discharge (g/L)")
    if hb is not None:
        p4.append(f"Hb izrakstīšanas brīdī {_comma(hb)} g/L.")
    had_catheter = tamponade or names or any(t in treat for t in ("Bladder irrigation (3-way catheter)",
                                                                  "Manual bladder washout"))
    if v.get("Discharged with catheter"):
        text = f"{f['Pacients']} {f['izrakstits']} ar urīnpūšļa ilgkatetru"
        if v.get("Catheter removal date"):
            text += f", kura evakuācija plānota {fmt_date(v['Catheter removal date'])}"
        p4.append(text + ".")
    elif had_catheter:
        p4.append("Urīnpūšļa katetrs evakuēts, urinē patstāvīgi.")
    else:
        p4.append("Urinē patstāvīgi.")
    epicrisis = "\n".join(" ".join(p) for p in (p1, p2, p3, p4) if p)
    epicrisis += f"\nVispārējā apmierinošā stāvoklī {f['izrakstits']} tālākai ārstēšanai ambulatori."

    # Recommendations
    recs = [TEMPORARY_PAPER_TEXT] if (names and v.get("Discharge paper (R31)") == "Temporary") else []
    recs += [
        "Atrasties ģimenes ārsta un ārstējošā urologa uzraudzībā.",
        "Diēta: vispusīga. Lietot šķidrumu 2,0–2,5 L diennaktī, ja nav medicīnisku kontrindikāciju.",
        f"Režīms: saudzējošs {WEEKS_TEXT[v['Activity restriction']]} no {'operācijas' if names else 'izrakstīšanas'} "
        "dienas, izvairīties no smagas fiziskas slodzes, smagumu celšanas un aizcietējumiem.",
    ]
    if v.get("Discharged with catheter"):
        if v.get("Catheter removal date"):
            recs.append("Kopt urīnpūšļa ilgkatetru; " + CATHETER_CARE_LINES[1][0].lower() + CATHETER_CARE_LINES[1][1:])
            recs.append(f"Katetra evakuācija plānota {fmt_date(v['Catheter removal date'])}.")
        else:
            recs += CATHETER_CARE_LINES
    line = _hem_antithrombotic_line(v)
    if line:
        recs.append(line)
    meds, advice = _hem_split_blocks(blocks)
    if meds:
        recs.append("Terapijā lietot:")
        recs += ["• " + b for b in meds]
    recs += advice
    if histology == "Pending":
        recs.append("Patohistoloģiskās izmeklēšanas rezultātu izvērtēs ārstējošais urologs.")
    if "Selective embolisation" in procs:
        recs.append("Pēc embolizācijas pirmajās dienās iespējamas sāpes vēdera lejasdaļā vai sēžamvietā un "
                    "subfebrila temperatūra; stipru sāpju, drudža vai kāju vājuma gadījumā vērsties NMC.")
    if v.get("Green corridor (GP referral)") in GREEN_CORRIDOR:
        recs.append(GREEN_CORRIDOR[v["Green corridor (GP referral)"]])
    recs.append(f"Vizīte pie urologa ambulatori: {WEEKS_TEXT[v['Urology visit']]} pēc izrakstīšanas.")
    recs += tests_lines(v, "plain", after_op=bool(names))
    if "Smoker or ex-smoker" in (v.get("Risk factors (R31)") or []):
        recs.append(HEM_SMOKING)
    recs.append(HEM_WARNING)
    return {"clinical_diagnosis": diagnosis, "epicrisis": epicrisis, "recommendations": "\n".join(recs)}


def r31_checks(v):
    blockers, cautions = [], []
    ward = v.get("_department") == "4"
    causes = v.get("Hematuria cause") or []
    chosen = v.get("_blocks") or []
    tests = v.get("_test_ids") or []
    procs = (v.get("Hematuria procedure") or []) if ward else []
    colour = v.get("Urine colour at discharge")
    drugs = v.get("Antithrombotic therapy") or []
    hb = _num(v, "Hemoglobin (g/L)")
    names = _hem_procs(v) if ward else None
    if v.get("_sex") == "Woman" and any(c in causes for c in ("BPH (prostatic bleeding)", "Prostate cancer")):
        blockers.append("BPH / prostate cancer is chosen as the cause for a woman.")
    if "Not clear yet" in causes and len(causes) > 1:
        cautions.append("'Not clear yet' is chosen together with a cause: remove one of them.")
    if colour == "Red / with clots":
        blockers.append("The urine is still red with clots: ongoing bleeding – BAUS: admit (3-way catheter, manual "
                        "washout, then irrigation); do not discharge.")
    if not ward:
        if v.get("Hemodynamics (R31)") == "Unstable":
            blockers.append("Hemodynamically unstable: resuscitate and admit.")
        if hb is not None and hb < 80:
            blockers.append("Hb below 80 g/L: transfusion range – admit.")
        elif hb is not None and hb < 100:
            cautions.append(f"Hb {_comma(hb)} g/L: recheck the blood count within 48–72 hours (tests box) and admit if "
                            "it falls.")
        if v.get("Clot retention (bladder tamponade)"):
            cautions.append("Clot retention: BAUS – it usually needs admission (3-way catheter 20–24 Fr, manual washout "
                            "first, then continuous irrigation). Send home only if after the washout the urine stays "
                            "clear or light pink without clots, Hb is stable, the catheter drains well and the patient "
                            "can come back quickly.")
        if colour == "Pink, no clots":
            cautions.append("The urine is still pink: make sure it keeps clearing and the patient voids (or the catheter "
                            "drains) before discharge.")
        if v.get("Fever at NMC"):
            cautions.append("Fever with hematuria: exclude pyelonephritis / prostatitis and obstruction; admit if septic, "
                            "obstructed or unable to take oral antibiotics.")
    else:
        if v.get("Histology (R31)") == "Known" and names and blank(v.get("Histology result (R31)")):
            blockers.append("Type the histology result.")
        hb2 = _num(v, "Hemoglobin at discharge (g/L)")
        if hb2 is not None and hb2 < 80:
            cautions.append("Hb at discharge below 80 g/L: check that a transfusion is not needed before discharge.")
        if "Selective embolisation" in procs:
            cautions.append("After embolisation: watch for post-embolisation syndrome (pain, fever), buttock pain or "
                            "ischaemia; check creatinine after the contrast.")
        if v.get("Cystoscopy finding (R31)") == "Bladder tumour" and "TURU (resection / biopsy)" not in procs:
            cautions.append("Bladder tumour seen at cystoscopy: plan TURU (EAU) and the oncology pathway.")
        if names and v.get("Histology (R31)") == "Pending" and v.get("Discharge paper (R31)") == "Final":
            cautions.append("Histology is still pending, but the discharge paper is marked as final.")
        if names and v.get("Histology (R31)") == "Known" and v.get("Discharge paper (R31)") == "Temporary":
            cautions.append("Histology is already known, but the discharge paper is marked as temporary.")
        if v.get("Cystoscopy finding (R31)") == "Blood from a ureteric orifice" and "ctu" not in tests:
            cautions.append("Bleeding from a ureteric orifice: the upper tract needs CT urography (and ureteroscopy if "
                            "CT is not conclusive).")
    # work-up (EAU NMIBC, strong): cystoscopy + upper tract imaging for visible haematuria
    if _hem_unknown(v) and not names:
        if v.get("Green corridor (GP referral)") != "Bladder – Z03.167" and "cysto" not in tests:
            cautions.append("Cause not clear: EAU (strong) – every patient with visible hematuria needs a cystoscopy. "
                            "Choose the green corridor (Z03.167) or tick cystoscopy in the tests.")
    if _hem_unknown(v) and v.get("CT at NMC (R31)") in (None, "Not done") and not any(t in tests for t in
                                                                                     ("ctu", "us_kb", "ct_abd")):
        cautions.append("Cause not clear: EAU (strong) – upper tract imaging (kidney ultrasound and/or CT urography). "
                        "Tick it in the tests.")
    if v.get("Hematuria type", "").startswith("Non-visible"):
        cautions.append("Microhematuria: confirm by microscopy (> 3 RBC/HPF), exclude UTI, then risk-stratify (AUA: "
                        "low risk – repeat urinalysis within 6 months; intermediate – cystoscopy and kidney ultrasound; "
                        "high – cystoscopy and CT urography). Green corridor (Z03.167): age >= 60 with dysuria or "
                        "leukocytosis.")
    if drugs and _hem_unknown(v):
        cautions.append("Hematuria on anticoagulants / antiplatelets needs the same full work-up (AUA, strong) – the "
                        "drug does not explain it.")
    if drugs and v.get("Antithrombotic plan") == "Pause until the urine is clear":
        cautions.append("Before pausing: do not stop dual antiplatelet therapy after a recent coronary stent, or "
                        "anticoagulation with a mechanical valve or recent thrombosis, without the cardiologist; the "
                        "text says who restarts it.")
    if "Warfarin" in drugs:
        inr = _num(v, "INR (optional)")
        if inr is None:
            cautions.append("Warfarin: record the INR.")
        elif inr > 3:
            cautions.append(f"INR {_comma(inr)}: above the usual range – correct it with the prescriber.")
    if "tranexamic" in chosen:
        if _hem_upper_tract(v):
            cautions.append("Tranexamic acid with a possible upper tract source: SmPC – clots may obstruct the ureter. "
                            "Avoid it.")
        if drugs:
            cautions.append("Tranexamic acid together with an antithrombotic drug: thrombosis risk – only under "
                            "supervision (SmPC).")
        if v.get("Labs at NMC") in ("Raised creatinine", "UTI and raised creatinine"):
            cautions.append("Raised creatinine: reduce the tranexamic acid dose (SmPC, by creatinine).")
    if "fraxiparine" in chosen and colour != "Clear":
        cautions.append("Fraxiparine while the urine is not yet clear: weigh the bleeding risk.")
    if "Radiation cystitis" in causes:
        cautions.append("Radiation cystitis – stepwise: washout and irrigation → cystoscopy with clot evacuation and "
                        "coagulation (biopsy to exclude a recurrence or second tumour) → intravesical alum 1% (or "
                        "formalin 1–2% under anaesthesia after a cystogram) → hyperbaric oxygen → selective "
                        "embolisation → urinary diversion as the last resort.")
    uti = "Urinary tract infection (hemorrhagic cystitis)" in causes or \
        v.get("Labs at NMC") in ("Urinary tract infection", "UTI and raised creatinine")
    if uti and not any(b in chosen for b in ("fosfomycin-3g", "amoxiclav-7", "cefuroxime-7", "cipro")) and \
            "IV antibiotics" not in (v.get("Ward treatment (R31)") or [] if ward else []):
        cautions.append("UTI: choose an antibiotic (and adjust it to the urine culture).")
    if "Urinary tract infection (hemorrhagic cystitis)" in causes:
        cautions.append("Hemorrhagic cystitis: repeat urinalysis after treatment – hematuria that persists after the UTI "
                        "is treated needs the full work-up (green corridor criterion at >= 45 years).")
    if "fosfomycin-3g" in chosen and v.get("_sex") == "Man":
        cautions.append("Fosfomycin single dose is for uncomplicated cystitis in women; a UTI in a man is complicated – "
                        "7 days, adjusted to the culture.")
    if v.get("Ultrasound at NMC") and "Hydronephrosis" in (v.get("Ultrasound findings (R31)") or []):
        cautions.append("Hydronephrosis with hematuria: exclude obstruction by clots, stone or tumour; drain if infected "
                        "or renal function falls.")
    return blockers, cautions


BUILDERS["r31_nmc"] = r31_nmc_parts
BUILDERS["r31_ward"] = r31_ward_parts
CHECKS["R31"] = r31_checks


# ---------------------------------------------------------------
# 2t. ACUTE CYSTITIS (N30.0) at NMC, sent home
# ---------------------------------------------------------------
# EAU Urological Infections 2025: diagnose cystitis from a focused history (dysuria, frequency, urgency, no vaginal
# discharge); dipstick is optional (weak); urine culture with suspected systemic UTI, symptoms that do not resolve or
# recur within 4 weeks, atypical symptoms, risk of resistant pathogens and in pregnancy (strong). First line in
# women: fosfomycin 3 g once, nitrofurantoin 5 days, pivmecillinam, nitroxoline; no aminopenicillins or
# fluoroquinolones (strong). Non-antibiotic treatment (e.g. ibuprofen) is an alternative in non-geriatric patients
# after shared decision-making. No routine urinalysis or culture after treatment. Recurrent (>= 2 in 6 months or
# >= 3 a year): culture; no routine work-up in women < 40 without risk factors; vaginal oestrogen after the
# menopause, cranberry, methenamine hippurate, D-mannose (weak), antimicrobial prophylaxis or self-start therapy.
# Campbell-Walsh-Wein: isolated vs unresolved vs recurrent infection; recurrence is reinfection (new organism) or
# bacterial persistence (the same organism from a focus – stone, obstruction, residual urine, diverticulum, chronic
# prostatitis), and persistence needs imaging and a urological work-up. Men: culture, 7 days, think of prostatitis
# and incomplete emptying.

N30_SYMPTOMS = {"Dysuria": "dedzināšanu un sāpēm urinējot", "Frequency": "biežu urinēšanu",
                "Urgency": "neatliekamu vēlmi urinēt", "Suprapubic pain": "sāpēm virs simfīzes",
                "Visible hematuria": "asiņu piejaukumu urīnā", "New urge incontinence": "urīna nesaturēšanu",
                "Cloudy / foul-smelling urine": "duļķainu urīnu ar nepatīkamu smaku"}
N30_DURATION = {"Since today": "kopš šodienas", "1–2 days": "1–2 dienas", "3–7 days": "3–7 dienas",
                "More than a week": "ilgāk par nedēļu"}
N30_RISK = {"Diabetes": "cukura diabēts", "Immunosuppression": "imūnsupresija",
            "Urinary tract abnormality / stone / obstruction": "urīnceļu anomālija / konkrementi / obstrukcija",
            "Recent urological procedure": "nesen veikta uroloģiska manipulācija",
            "Indwelling catheter or self-catheterisation": "urīnpūšļa katetrs / paškatetrizācija",
            "Renal impairment": "nieru funkcijas traucējumi", "Neurogenic bladder": "neirogēns urīnpūslis"}
N30_UA = {"Leukocytes": "leikocitūrija", "Nitrite": "nitrīti pozitīvi", "Erythrocytes": "eritrocitūrija",
          "Bacteria": "bakteriūrija"}
N30_US = {"Upper tract dilatation": "augšējo urīnceļu dilatācija", "Bladder wall thickening":
          "urīnpūšļa sienas sabiezējums", "Bladder stone": "urīnpūšļa konkrements", "Enlarged prostate":
          "palielināta prostata"}
N30_BLOOD = {"No inflammatory changes": "Asins analīzēs bez iekaisuma izmaiņām.",
             "Raised CRP / leukocytosis": "Asins analīzēs – paaugstināts CRO / leikocitoze.",
             "Raised creatinine": "Asins analīzēs – paaugstināts kreatinīna līmenis."}
N30_ANTIBIOTICS = ("fosfomycin-3g", "furamag", "nitrofurantoin", "cefuroxime-5", "amoxiclav-5", "cotrimoxazole-3",
                   "furamag-men", "cotrimoxazole-7", "cefuroxime-7", "amoxiclav-7", "cipro")
N30_WARNING = ("Ja paaugstinās ķermeņa temperatūra virs 38 °C, parādās drebuļi, sāpes jostas rajonā, slikta dūša vai "
               "vemšana, pastiprinās hematūrija vai nav iespējams urinēt, nekavējoties vērsties NMC.")


def _n30_flags(v):
    woman = v.get("_sex") == "Woman"
    us = v.get("Ultrasound (N30)") or []
    abnormal_us = any(x in us for x in ("Post-void residual urine", "Upper tract dilatation", "Bladder stone"))
    uncomplicated = woman and not v.get("Pregnant (N30)") and not (v.get("Risk factors (N30)") or []) \
        and not abnormal_us and not v.get("Same organism soon after treatment")
    return woman, uncomplicated


def n30_nmc_parts(sex, v, blocks=()):
    f = SEX_FORMS[sex]
    woman, uncomplicated = _n30_flags(v)
    symptoms = v.get("Cystitis symptoms") or []
    hemorrhagic = "Visible hematuria" in symptoms
    pregnant = woman and v.get("Pregnant (N30)")
    recurrent = v.get("Recurrent cystitis")
    week = _num(v, "Pregnancy week")

    # Clinical diagnosis
    lines = ["Akūts " + ("nekomplicēts " if uncomplicated else "") + ("hemorāģisks " if hemorrhagic else "") +
             "cistīts."]
    if recurrent:
        lines.append("Recidivējošas apakšējo urīnceļu infekcijas.")
    if pregnant:
        lines.append("Grūtniecība" + (f" ({int(week)}. nedēļa)" if week else "") + ".")
    if not blank(v.get("Other diagnoses (optional)")):
        lines.append(v["Other diagnoses (optional)"].strip().rstrip(".") + ".")

    # Epicrisis
    complaint = lv_list([N30_SYMPTOMS[x] for x in symptoms if x in N30_SYMPTOMS]) or "diskomfortu urinējot"
    s = [f"{f['Pacients']} {f['iestajies']} PSKUS NMC ar sūdzībām par {complaint} "
         f"{N30_DURATION[v['Symptom duration (N30)']]}."]
    negatives = "Drudzi, drebuļus un sāpes jostas rajonā noliedz."
    if woman:
        negatives += " Atzīmē izdalījumus no maksts." if v.get("Vaginal discharge") else \
            " Izdalījumus no maksts noliedz."
    s.append(negatives)
    history = []
    if recurrent:
        history.append("recidivējošas urīnceļu infekcijas" +
                       (" (atkārtoti izdalīts tas pats mikroorganisms)" if v.get("Same organism soon after treatment")
                        else ""))
    history += [N30_RISK[x] for x in (v.get("Risk factors (N30)") or []) if x in N30_RISK]
    if woman and v.get("Postmenopausal"):
        history.append("postmenopauze")
    if history:
        s.append(f"Anamnēzē – {lv_list(history)}.")
    if pregnant:
        s.append("Grūtniecība" + (f" ({int(week)}. nedēļa)" if week else "") + ".")
    if v.get("Antibiotic in the last 3 months"):
        s.append(f"Pēdējo 3 mēnešu laikā {'lietojusi' if woman else 'lietojis'} antibakteriālu terapiju.")
    if v.get("Penicillin allergy"):
        s.append("Penicilīna grupas preparātu nepanesība.")
    s.append(f"NMC etapā vispusīgi {f['izmekleta']}. Objektīvi: vēders mīksts, "
             f"{'virs simfīzes jutīgs' if v.get('Suprapubic tenderness') else 'nesāpīgs'}; Pasternacka simptoms "
             "negatīvs abās pusēs. Ķermeņa temperatūra normāla.")
    ua = [N30_UA[x] for x in (v.get("Urinalysis (N30)") or []) if x in N30_UA]
    culture = v.get("Urine culture taken")
    if ua:
        s.append(f"Urīna analīzē – {lv_list(ua)}" + ("; paņemts urīna uzsējums ar jutības noteikšanu" if culture
                                                     else "") + ".")
    elif culture:
        s.append("Paņemts urīna uzsējums ar jutības noteikšanu.")
    if v.get("Blood tests (N30)") in N30_BLOOD:
        s.append(N30_BLOOD[v["Blood tests (N30)"]])
    us = v.get("Ultrasound (N30)") or []
    if us:
        found = []
        if "Post-void residual urine" in us:
            pvr = _num(v, "Residual urine (ml)")
            found.append("atlieku urīns" + (f" ~{_comma(pvr)} ml" if pvr is not None else ""))
        found += [N30_US[x] for x in us if x in N30_US]
        s.append("USS – " + (", ".join(found) if found else "nieres un urīnpūslis bez patoloģiskas atrades") + ".")
    if not blank(v.get("Other NMC findings (optional)")):
        text = v["Other NMC findings (optional)"].strip()
        s.append(text if text.endswith(".") else text + ".")
    s.append("Dotajā brīdī nav indikāciju neatliekamai stacionēšanai. Vispārēji apmierinošā stāvoklī un ar "
             "rekomendācijām izraksta tālākai ambulatorai ārstēšanai.")

    # Recommendations
    numbered = make_counter()
    recs = [numbered("Atrasties ģimenes ārsta uzraudzībā."),
            numbered("Lietot šķidrumu 1,5–2 l/DN, ja nav medicīnisku kontrindikāciju; regulāri un pilnībā iztukšot "
                     "urīnpūsli, nenoturēt urīnu.")]
    meds, advice = split_advice(blocks)
    delayed = []
    if v.get("Treatment approach (N30)") == "Symptomatic first (delayed antibiotic)":
        abx_texts = {b["text"] for b in BLOCKS["N30.0"] if b["id"] in N30_ANTIBIOTICS}
        delayed = [b for b in meds if b in abx_texts]
        meds = [b for b in meds if b not in abx_texts]
        recs.append(numbered("Sākotnēji simptomātiska terapija un pietiekams šķidruma daudzums."))
    if meds:
        recs.append(numbered("Terapijā lietot:"))
        recs += ["    " + b for b in meds]
    if delayed:
        recs.append(numbered("Ja 48 stundu laikā sūdzības nemazinās vai pastiprinās, sākt antibakteriālo terapiju:"))
        recs += ["    " + b for b in delayed]
    if culture:
        recs.append(numbered("NMC paņemtā urīna uzsējuma rezultātu izvērtēt ģimenes ārstam (Datamed.lv); ja izdalītais "
                             "mikroorganisms nav jutīgs pret nozīmēto antibiotiku, koriģēt terapiju atbilstoši "
                             "jutībai."))
    recs.append(numbered(("Ja arī antibakteriālās terapijas laikā " if delayed else "Ja ") +
                         "48–72 stundu laikā sūdzības nemazinās, saglabājas ārstēšanas beigās vai 4 nedēļu laikā "
                         "atkārtojas, vērsties pie ģimenes ārsta – nepieciešams urīna uzsējums ar jutības noteikšanu un "
                         "citas antibiotikas izvēle."))
    if not pregnant and not hemorrhagic:
        recs.append(numbered("Kontroles urīna analīze pēc ārstēšanas nav nepieciešama, ja sūdzības pilnībā izzudušas."))
    if pregnant:
        recs.append(numbered("Informēt ginekologu, kas vada grūtniecību."))
    if hemorrhagic:
        recs.append(numbered("Ja pēc infekcijas izārstēšanas hematūrija saglabājas, nepieciešama izmeklēšana – vērsties "
                             "pie ģimenes ārsta (nosūtījums uz cistoskopiju; no 45 gadu vecuma \"zaļā koridora\" "
                             "ietvaros, diagnozes kods Z03.167)."))
    if recurrent:
        recs.append(numbered("Recidīvu profilaksei: pietiekams šķidruma daudzums, regulāra urīnpūšļa iztukšošana" +
                             (" (arī pēc dzimumakta)" if woman else "") + ", aizcietējumu profilakse" +
                             ("; nelietot spermicīdus" if woman else "") + "."))
    recs += [numbered(b) for b in advice]
    referral = v.get("Referral (N30)")
    if referral == "Urologist":
        why = ("izvērtēt urīnpūšļa iztukšošanos un prostatu" if not woman
               else "izvērtēt atkārtoto urīnceļu infekciju cēloni un profilakses iespējas" if recurrent
               else "tālākās taktikas izvērtēšanai")
        recs.append(numbered(f"Ambulatori urologa konsultācija – {why}."))
    elif referral == "Gynaecologist":
        why = ("izdalījumu no maksts izvērtēšanai" if v.get("Vaginal discharge")
               else "vaginālās estrogēnu terapijas izvērtēšanai" if v.get("Postmenopausal")
               else "tālākās taktikas izvērtēšanai")
        recs.append(numbered(f"Ginekologa konsultācija – {why}."))
    recs += tests_lines(v, "numbered", numbered)
    recs.append(numbered(N30_WARNING))
    return {"clinical_diagnosis": "\n".join(lines), "epicrisis": " ".join(s), "recommendations": "\n".join(recs)}


def n30_checks(v):
    blockers, cautions = [], []
    woman, uncomplicated = _n30_flags(v)
    chosen = v.get("_blocks") or []
    pregnant = woman and v.get("Pregnant (N30)")
    risks = v.get("Risk factors (N30)") or []
    culture = v.get("Urine culture taken")
    abx = [b for b in chosen if b in N30_ANTIBIOTICS]
    us = v.get("Ultrasound (N30)") or []
    elderly = "Age 65 or older / frail" in risks
    if v.get("Pyelonephritis signs (N30)"):
        blockers.append("Fever, chills, flank pain or vomiting: this is pyelonephritis / a systemic UTI, not cystitis – "
                        "use N10.")
    if pregnant and any(b in chosen for b in ("furamag", "furamag-men")):
        blockers.append("Furamag is contraindicated in pregnancy and breastfeeding (label). Fosfomycin 3 g or a "
                        "beta-lactam by culture instead.")
    if not v.get("Cystitis symptoms"):
        cautions.append("No symptoms are chosen: cystitis is a clinical diagnosis (dysuria, frequency, urgency). "
                        "Bacteriuria without symptoms is not treated (EAU, strong), except in pregnancy.")
    if not abx and v.get("Treatment approach (N30)") == "Antibiotic":
        cautions.append("Choose an antibiotic in the blocks (or 'Symptomatic first').")
    if v.get("Treatment approach (N30)") == "Symptomatic first (delayed antibiotic)":
        if not woman or pregnant or elderly or risks or v.get("Recurrent cystitis"):
            cautions.append("Symptomatic-first treatment: EAU – only for non-geriatric women with mild symptoms after "
                            "shared decision-making; not for men, pregnancy, the frail elderly or complicating factors "
                            "(more pyelonephritis in the trials).")
        if not abx:
            cautions.append("Symptomatic first: also choose the antibiotic to start if there is no improvement in 48 h.")
    if "cipro" in chosen:
        cautions.append("Ciprofloxacin: EAU (strong) – do not use fluoroquinolones for cystitis.")
    if "amoxiclav-5" in chosen and not pregnant and not culture:
        cautions.append("Amoksiklav: EAU (strong) – no aminopenicillins empirically; use it by culture (or in "
                        "pregnancy).")
    if not woman:
        if "fosfomycin-3g" in chosen:
            cautions.append("Fosfomycin single dose is for women; cystitis in a man needs 7 days.")
        if any(b in chosen for b in ("furamag", "nitrofurantoin", "cefuroxime-5", "amoxiclav-5", "cotrimoxazole-3")):
            cautions.append("A short women's regimen is chosen for a man: use a 7-day regimen (blocks 'men').")
        if not culture:
            cautions.append("Cystitis in a man: take a urine culture and think of prostatitis and incomplete emptying "
                            "(residual urine).")
    if pregnant:
        if not culture:
            cautions.append("Pregnancy: EAU (strong) – take a urine culture.")
        if "cotrimoxazole-3" in chosen or "cotrimoxazole-7" in chosen:
            cautions.append("Trimethoprim/sulfamethoxazole in pregnancy: not in the first trimester (trimethoprim) or "
                            "near term (sulfamethoxazole).")
        if "cipro" in chosen:
            blockers.append("Fluoroquinolones are not used in pregnancy.")
        if "nitrofurantoin" in chosen and (_num(v, "Pregnancy week") or 0) >= 36:
            cautions.append("Nitrofurantoin near term: risk of neonatal haemolysis – choose another drug.")
        if "ibuprofen" in chosen:
            cautions.append("NSAIDs in pregnancy: avoid (contraindicated from 20 weeks); use paracetamol.")
    if v.get("Recurrent cystitis") and not culture:
        cautions.append("Recurrent cystitis: EAU (strong) – diagnose it by urine culture.")
    if v.get("Antibiotic in the last 3 months") and not culture:
        cautions.append("Antibiotic in the last 3 months: higher risk of resistant bacteria – take a urine culture "
                        "(EAU, strong).")
    if risks and not culture and not v.get("Recurrent cystitis"):
        cautions.append("Complicating factors are present: take a urine culture and treat for 7 days.")
    if "Indwelling catheter or self-catheterisation" in risks:
        cautions.append("Catheter-associated UTI: not simple cystitis – change the catheter and culture from the new "
                        "one (see T83 / N39.0).")
    if woman and v.get("Vaginal discharge"):
        cautions.append("Vaginal discharge: EAU – the diagnosis of cystitis relies on its absence; consider vaginitis or "
                        "an STI (gynaecologist).")
    if v.get("Blood tests (N30)") == "Raised CRP / leukocytosis":
        cautions.append("Raised CRP / leukocytosis: look again for pyelonephritis or prostatitis.")
    renal = v.get("Blood tests (N30)") == "Raised creatinine" or "Renal impairment" in risks
    if renal and any(b in chosen for b in ("furamag", "furamag-men", "nitrofurantoin")):
        cautions.append("Renal impairment: nitrofurans work poorly and accumulate – nitrofurantoin not below eGFR 45 "
                        "(short courses at 30–44); Furamag is contraindicated in severe renal failure.")
    if "ibuprofen" in chosen and (renal or elderly):
        cautions.append("Ibuprofen with renal impairment or in the elderly: EAU – NSAIDs are potentially "
                        "inappropriate; use paracetamol.")
    if v.get("Penicillin allergy") and any(b in chosen for b in ("amoxiclav-5", "amoxiclav-7")):
        blockers.append("Penicillin allergy: Amoksiklav is chosen.")
    if v.get("Penicillin allergy") and any(b in chosen for b in ("cefuroxime-5", "cefuroxime-7")):
        cautions.append("Penicillin allergy: cefuroxime only if the allergy was not severe (anaphylaxis).")
    if woman and v.get("Postmenopausal") and v.get("Recurrent cystitis") and "estriol" not in chosen:
        cautions.append("Postmenopausal woman with recurrent cystitis: EAU (strong) – vaginal oestrogen (block "
                        "'Vaginal estriol').")
    if v.get("Recurrent cystitis") and v.get("Same organism soon after treatment"):
        cautions.append("Same organism again soon after treatment – bacterial persistence (Campbell-Walsh-Wein): look "
                        "for a focus (stone, obstruction, residual urine, diverticulum, chronic prostatitis) – imaging "
                        "and urology referral.")
    if "Upper tract dilatation" in us:
        cautions.append("Upper tract dilatation with infection: obstruction plus infection needs drainage if any "
                        "systemic sign appears – consider admission and imaging.")
    pvr = _num(v, "Residual urine (ml)")
    if pvr is not None and pvr > 100:
        cautions.append(f"Residual urine {_comma(pvr)} ml: incomplete emptying is a complicating factor – treat for "
                        "7 days and refer to a urologist.")
    if "methenamine" in chosen and (risks or not woman):
        cautions.append("Methenamine hippurate: EAU recommends it for women without urinary tract abnormalities.")
    return blockers, cautions


BUILDERS["n30_nmc"] = n30_nmc_parts
CHECKS["N30.0"] = n30_checks


# ---------------------------------------------------------------
# 2m. FOLLOW-UP TESTS: catalogue, guideline-based suggestions, text
# ---------------------------------------------------------------
# Nothing here goes into a text by default. The doctor ticks tests in the
# "Follow-up tests" box of a diagnosis; the builders write them with tests_lines().

FOLLOWUP_TESTS = {  # id -> group, English name in the app, Latvian text in the discharge
    # urine
    "ua": {"group": "Urine", "title": "Urinalysis", "lv": "urīna analīze"},
    "uc": {"group": "Urine", "title": "Urine culture with susceptibility",
           "lv": "urīna uzsējums ar antibakteriālās jutības noteikšanu"},
    "cyto": {"group": "Urine", "title": "Urine cytology", "lv": "urīna citoloģiskā izmeklēšana"},
    "urine24": {"group": "Urine", "title": "24-hour urine metabolic evaluation (2 collections)",
                "lv": "2 secīgas 24 stundu urīna analīzes metaboliskai izmeklēšanai"},
    "stone": {"group": "Urine", "title": "Stone analysis", "lv": "konkrementa ķīmiskā sastāva analīze"},
    # blood
    "cbc": {"group": "Blood", "title": "Complete blood count", "lv": "pilna asinsaina"},
    "crp": {"group": "Blood", "title": "CRP", "lv": "CRO"},
    "creat": {"group": "Blood", "title": "Creatinine and eGFR", "lv": "kreatinīns, GFĀ"},
    "nak": {"group": "Blood", "title": "Sodium and potassium", "lv": "Na, K"},
    "stone_blood": {"group": "Blood", "title": "Stone blood tests (creatinine, calcium, uric acid)",
                    "lv": "kreatinīns, kalcijs (jonizētais vai kopējais ar albumīnu), urīnskābe"},
    "psa": {"group": "Blood", "title": "PSA", "lv": "PSA"},
    "glucose": {"group": "Blood", "title": "Blood glucose and HbA1c", "lv": "glikoze, HbA1c"},
    "inr": {"group": "Blood", "title": "INR (warfarin)", "lv": "INR"},
    "panel": {"group": "Blood", "title": "Extended blood panel (department list)",
              "lv": "CRO, LDH, sārmainā fosfatāze, kreatinīns, urīnskābe, Ca, P, Na, K, glikoze, ASAT, ALAT, kopējais "
                    "bilirubīns, 25-OH-vitamīns D (D3+D2)"},
    # imaging
    "us_kb": {"group": "Imaging", "title": "Kidney and bladder ultrasound", "lv": "nieru un urīnpūšļa USS"},
    "us_k": {"group": "Imaging", "title": "Kidney ultrasound", "lv": "nieru USS"},
    "us_scrotum": {"group": "Imaging", "title": "Scrotal ultrasound", "lv": "sēklinieku maisiņa USS"},
    "mri_spine": {"group": "Imaging", "title": "MRI of the lumbar spine (cauda equina)",
                  "lv": "MR mugurkaula jostas daļai (cauda equina sindroma izslēgšanai)"},
    "renogram": {"group": "Imaging", "title": "Diuretic renography (MAG3)", "lv": "diurētiskā renoscintigrāfija (MAG3)"},
    "kub": {"group": "Imaging", "title": "Plain X-ray KUB", "lv": "vēdera dobuma pārskata rentgenogramma (KUB)"},
    "ncct": {"group": "Imaging", "title": "Non-contrast low-dose CT (stones)",
             "lv": "nekontrasta zemas devas CT vēdera dobumam un mazajam iegurnim"},
    "ct_abd": {"group": "Imaging", "title": "Contrast CT of the abdomen", "lv": "CT vēdera dobumam ar i/v kontrastvielu"},
    "ctu": {"group": "Imaging", "title": "CT urography", "lv": "CT urogrāfija"},
    "ct_ca": {"group": "Imaging", "title": "CT chest and abdomen", "lv": "CT krūškurvim un vēdera dobumam"},
    "ct_stage": {"group": "Imaging", "title": "Staging CT chest, abdomen and pelvis with CT urography",
                 "lv": "CT krūškurvim, vēdera dobumam un mazajam iegurnim ar CT urogrāfiju"},
    "mri_prostate": {"group": "Imaging", "title": "Prostate MRI", "lv": "prostatas MR"},
    "cxr": {"group": "Imaging", "title": "Chest X-ray", "lv": "krūškurvja rentgenogramma"},
    "bone_pet": {"group": "Imaging", "title": "Bone scan or PET-CT", "lv": "skeleta scintigrāfija vai PET-CT"},
    "rug": {"group": "Imaging", "title": "Retrograde urethrography", "lv": "retrogrāda uretrogrāfija"},
    "rug_vcug": {"group": "Imaging", "title": "Retrograde urethrography + voiding cystourethrography",
                 "lv": "retrogrāda uretrogrāfija un mikcijas cistouretrogrāfija"},
    "rug_pericath": {"group": "Imaging", "title": "Pericatheter urethrography before catheter removal",
                     "lv": "pericatetrāla retrogrāda uretrogrāfija (vai mikcijas cistouretrogrāfija), lai izslēgtu "
                           "kontrastvielas ekstravazāciju"},
    # functional
    "uroflow": {"group": "Functional", "title": "Uroflowmetry", "lv": "urīna plūsmas mērījums (uroflowmetrija)"},
    "pvr": {"group": "Functional", "title": "Post-void residual by ultrasound", "lv": "atlieku urīna noteikšana ar USS"},
    "ipss": {"group": "Functional", "title": "IPSS questionnaire", "lv": "IPSS anketa"},
    "iief": {"group": "Functional", "title": "IIEF-5 questionnaire (erectile function)",
             "lv": "IIEF-5 anketa (erektilā funkcija)"},
    "continence": {"group": "Functional", "title": "Continence assessment (ICIQ-UI SF, pad count)",
                   "lv": "urīna nesaturēšanas novērtējums (ICIQ-UI SF anketa, ieliktņu skaits diennaktī)"},
    "proms": {"group": "Functional", "title": "Stricture questionnaires (USS-PROM, sexual function)",
              "lv": "simptomu un seksuālās funkcijas anketas (USS-PROM, IIEF-5)"},
    # endoscopy
    "cysto": {"group": "Endoscopy", "title": "Cystoscopy", "lv": "cistoskopija"},
    "urethrocysto": {"group": "Endoscopy", "title": "Urethrocystoscopy", "lv": "uretrocistoskopija"},
    "returu": {"group": "Endoscopy", "title": "Second resection (re-TURU)", "lv": "atkārtota TURU (re-TURU)"},
    # later follow-up schedules (written after "Turpmāk:")
    "psa_later": {"group": "Schedule", "title": "Later PSA schedule after prostatectomy",
                  "lv": "PSA ik pēc 6 mēnešiem 3 gadus, pēc tam reizi gadā"},
    "ct_later": {"group": "Schedule", "title": "Later CT schedule after kidney cancer surgery",
                 "lv": "CT krūškurvim un vēdera dobumam atbilstoši recidīva riska grupai"},
    "cysto_later": {"group": "Schedule", "title": "Later cystoscopy and cytology schedule (high risk)",
                    "lv": "cistoskopija un urīna citoloģija ik pēc 3 mēnešiem 2 gadus, ik pēc 6 mēnešiem līdz "
                          "5 gadiem, pēc tam reizi gadā"},
    "ctu_yearly": {"group": "Schedule", "title": "Yearly CT urography (high risk)", "lv": "CT urogrāfija reizi gadā"},
    "stricture_later": {"group": "Schedule", "title": "Later stricture follow-up",
                        "lv": "12 un 24 mēnešus pēc operācijas – uroflowmetrija un simptomu anketas"},
    "cysto_low": {"group": "Schedule", "title": "Later cystoscopy schedule (low risk)",
                  "lv": "cistoskopija 12 mēnešus pēc operācijas, pēc tam reizi gadā 5 gadus"},
    "cysto_int": {"group": "Schedule", "title": "Later cystoscopy schedule (intermediate risk, Ta LG)",
                  "lv": "cistoskopija ik pēc 6 mēnešiem 2 gadus, pēc tam reizi gadā 10 gadus"},
    "ct_cap": {"group": "Imaging", "title": "CT chest, abdomen and pelvis with contrast",
               "lv": "CT krūškurvim, vēdera dobumam un mazajam iegurnim ar i/v kontrastvielu"},
    "ct_rc_later": {"group": "Schedule", "title": "Later CT schedule after radical cystectomy",
                    "lv": "CT krūškurvim, vēdera dobumam un mazajam iegurnim ik pēc 6 mēnešiem līdz 3 gadiem, pēc tam "
                          "reizi gadā"},
    "b12": {"group": "Blood", "title": "Vitamin B12 (yearly after bowel diversion)", "lv": "B12 vitamīns"},
    "naat": {"group": "Urine", "title": "First-void urine NAAT (chlamydia, gonorrhoea, M. genitalium)",
             "lv": "pirmās porcijas urīna NAAT (C. trachomatis, N. gonorrhoeae, M. genitalium)"},
    "trus": {"group": "Imaging", "title": "Transrectal ultrasound of the prostate (TRUS)",
             "lv": "transrektāla prostatas USS (TRUS)"},
    "meares": {"group": "Urine", "title": "Meares–Stamey 2- or 4-glass test",
               "lv": "Meares–Stamey 2 vai 4 trauku tests (hroniska bakteriāla prostatīta izslēgšanai)"},
    "markers": {"group": "Blood", "title": "Testicular tumour markers (AFP, β-hCG, LDH)",
                "lv": "audzēja marķieri (AFP, β-hCG, LDH)"},
    "bgas": {"group": "Blood", "title": "Venous blood gas (bicarbonate)",
             "lv": "venozo asiņu gāzu analīze (bikarbonāts, skābju-sārmu līdzsvars)"},
}

# When a test is done: key -> (English, Latvian, Latvian form after an operation if it differs)
TIMINGS = {
    "asap": ("As soon as possible", "pēc iespējas ātrāk", None),
    "48-72h": ("In 48–72 hours", "pēc 48–72 stundām", None),
    "no_improve": ("If no improvement within 72 hours", "ja 72 stundu laikā stāvoklis neuzlabojas", None),
    "1w": ("In 1 week", "pēc 1 nedēļas", "1 nedēļu pēc operācijas"),
    "2w": ("In 2 weeks", "pēc 2 nedēļām", "2 nedēļas pēc operācijas"),
    "2-6w": ("In 2–6 weeks", "pēc 2–6 nedēļām", "2–6 nedēļas pēc operācijas"),
    "4-6w": ("In 4–6 weeks", "pēc 4–6 nedēļām", "4–6 nedēļas pēc operācijas"),
    "6w": ("In 6 weeks", "pēc 6 nedēļām", "6 nedēļas pēc operācijas"),
    "catheter": ("Before catheter removal", "pirms katetra evakuācijas", None),
    "stent": ("Before DJ stent removal", "pirms DJ stenta izņemšanas", None),
    "visit": ("Before the urology visit", "pirms vizītes pie urologa", None),
    "at_visit": ("At the urology visit", "urologa vizītes laikā", None),
    "after_stent": ("4–6 weeks after stent removal", "4–6 nedēļas pēc DJ stenta izņemšanas", None),
    "3m": ("In 3 months", "pēc 3 mēnešiem", "3 mēnešus pēc operācijas"),
    "6m": ("In 6 months", "pēc 6 mēnešiem", "6 mēnešus pēc operācijas"),
    "12m": ("In 12 months", "pēc 12 mēnešiem", "12 mēnešus pēc operācijas"),
    "histology": ("After the histology result",
                  "pēc histoloģiskās atbildes saņemšanas (laiku noteiks ārstējošais urologs)", None),
    "definitive": ("Before the definitive treatment", "pirms definitīvās ārstēšanas", None),
    "after_definitive": ("After the definitive treatment", "pēc definitīvās ārstēšanas", None),
    "next_change": ("Before the next nephrostomy change", "pirms nākamās nefrostomas maiņas", None),
    "symptoms": ("Only if symptoms appear", "simptomu gadījumā", None),
    "later": ("Later follow-up", "turpmāk", None),
    "none": ("No time given", "", None),
}
TESTS_BY = {  # who arranges the tests (added after "Kontroles izmeklējumi")
    "Not stated": "",
    "Family doctor's referral": " (ar ģimenes ārsta nosūtījumu)",
    "Treating urologist's referral": " (ar ārstējošā urologa nosūtījumu)",
    "At the family doctor's practice": " (ģimenes ārsta praksē)",
}


def tests_lines(v, style="plain", numbered=None, after_op=False):
    """The ticked follow-up tests as recommendation lines, grouped by when they are done.
    style: "plain" (heading + "• " lines), "bullet" ("• " heading + "    o " lines)
    or "numbered" (numbered heading + "    o " lines). Empty list if nothing is ticked."""
    chosen = v.get("_tests") or []
    if not chosen:
        return []
    groups = {}
    for item in chosen:
        groups.setdefault(item["timing"], []).append(item["text"])
    order = [key for key in TIMINGS if key in groups]
    by = TESTS_BY.get(v.get("_tests_by"), "")

    def when(key):
        _, plain, after = TIMINGS[key]
        return after if (after_op and after) else plain

    def lead(text):
        if style == "numbered" and numbered:
            return numbered(text)
        return ("• " if style in ("bullet", "numbered") else "") + text

    sub = "• " if style == "plain" else "    o "
    if len(order) == 1:
        key = order[0]
        head = "Kontroles izmeklējumi" + (f" {when(key)}" if when(key) else "") + by
        return [lead(f"{head}: {', '.join(groups[key])}.")]
    lines = [lead(f"Kontroles izmeklējumi{by}:")]
    for number, key in enumerate(order, start=1):
        text = ", ".join(groups[key])
        if when(key):
            text = f"{when(key)}: {text}"
        lines.append(sub + text[0].upper() + text[1:] + ("." if number == len(order) else ";"))
    return lines


# ----- suggestions: which tests fit this patient (EAU guidelines 2026 or best available data) -----
# Each suggestion: test id, level ("rec" = recommended, "consider", "avoid" = not routinely),
# default timing, the reason shown in the app, and optionally its own Latvian text.

def _s(test, level, timing, why, text=None):
    return {"id": test, "level": level, "timing": timing, "why": why, "text": text}


NO_ROUTINE_UA = "EAU: post-treatment urinalysis in patients without symptoms is not indicated."
NO_ROUTINE_UC = "EAU: post-treatment urine culture in patients without symptoms is not indicated."
NO_CATHETER_UC = ("EAU (strong): no routine urine culture in catheterised patients without symptoms; "
                  "culture only with infection symptoms, from a newly placed catheter.")
PRACTICE_LABS = "Only if inflammatory markers were still raised at discharge (common practice, not an EAU requirement)."


def n10_tests(v):
    out = []
    if v.get("_department") == "NMC":
        if v.get("NMC imaging") == "None":
            out.append(_s("us_kb", "rec", "asap", "No imaging was done at NMC. EAU (strong): image the urinary tract in "
                                                  "pyelonephritis to exclude obstruction or stones; ultrasound first."))
        out.append(_s("ct_abd", "consider", "no_improve", "EAU: contrast CT (or excretory urography) if the patient is still "
                                                          "febrile after 72 hours of treatment."))
        out.append(_s("cbc", "consider", "48-72h", "EAU: outpatients with pyelonephritis are reviewed within 48–72 hours; "
                                                  "blood tests at that review only if not improving."))
        out.append(_s("crp", "consider", "48-72h", "As above: CRP helps if the patient is not improving."))
    else:
        out.append(_s("creat", "consider", "1w", "If renal function was impaired during the stay."))
        out.append(_s("cbc", "consider", "1w", PRACTICE_LABS))
        out.append(_s("crp", "consider", "1w", PRACTICE_LABS))
        out.append(_s("us_kb", "consider", "visit", "EAU (strong) asks for imaging during the episode. Repeat ultrasound only "
                                                   "to recheck a finding (stone, dilatation, residual urine) or if none was done."))
        if v.get("Infection complexity") == "Complicated":
            out.append(_s("ctu", "consider", "visit", "Complicated pyelonephritis: look for the cause. EAU (strong): manage any "
                                                     "urological abnormality or complicating factor."))
    if v.get("_sex") == "Man":
        out.append(_s("uroflow", "consider", "at_visit", "Men with febrile UTI: look for bladder outlet obstruction."))
        out.append(_s("pvr", "consider", "at_visit", "Men with febrile UTI: check that the bladder empties."))
    out.append(_s("ua", "avoid", "visit", NO_ROUTINE_UA))
    out.append(_s("uc", "avoid", "visit", NO_ROUTINE_UC))
    return out


def n39_tests(v):
    out = []
    upper_or_complicated = (v.get("Infection site") in ("Upper urinary tract", "Pyelonephritis")
                            or v.get("Infection complexity") == "Complicated")
    if "Acute kidney injury" in (v.get("Complications") or []) or v.get("Renal function trend") == "Deteriorated":
        out.append(_s("creat", "rec", "1w", "Renal function got worse during this episode: recheck creatinine and eGFR."))
    else:
        out.append(_s("creat", "consider", "1w", "Only if renal function was impaired or the patient takes nephrotoxic drugs."))
    if v.get("Imaging") == "Not done" and upper_or_complicated:
        out.append(_s("us_kb", "rec", "asap", "Upper-tract or complicated UTI without imaging. EAU (strong): image the "
                                              "urinary tract to exclude obstruction or stones."))
    elif v.get("Hydronephrosis") in ("Right", "Left", "Both"):
        out.append(_s("us_k", "consider", "visit", "Hydronephrosis was found: check that it has resolved."))
    if v.get("Catheter at discharge") == "Removed" and v.get("Voiding after removal") in (
            "Feeling of incomplete emptying", "Increased residual urine"):
        out.append(_s("pvr", "rec", "1w", "Incomplete emptying after catheter removal: recheck the residual urine."))
    if v.get("_sex") == "Man":
        out.append(_s("uroflow", "consider", "at_visit", "Men with UTI: look for bladder outlet obstruction "
                                                        "(EAU: manage complicating factors)."))
        out.append(_s("pvr", "consider", "at_visit", "Men with UTI: check that the bladder empties."))
    if "Hematuria" in (v.get("Symptoms") or []):
        out.append(_s("ua", "consider", "4-6w", "EAU: with haematuria, repeat urinalysis a few weeks after treatment; "
                                               "haematuria that persists needs a work-up."))
    if "Recurrent UTIs" in (v.get("Relevant history") or []):
        why = "Recurrent UTIs: look for a urological cause."
        if v.get("_sex") == "Woman":
            why += (" EAU (weak): no extensive routine work-up (cystoscopy, full abdominal ultrasound) in women "
                    "under 40 without risk factors.")
        out.append(_s("us_kb", "consider", "visit", why))
    out.append(_s("cbc", "consider", "1w", PRACTICE_LABS))
    out.append(_s("crp", "consider", "1w", PRACTICE_LABS))
    out.append(_s("ua", "avoid", "visit", NO_ROUTINE_UA))
    out.append(_s("uc", "avoid", "visit", NO_CATHETER_UC if v.get("Has catheter") else NO_ROUTINE_UC))
    return out


STONE_ANALYSIS_WHY = ("EAU (strong): analyse the stone in every first-time stone former (infrared spectroscopy or "
                      "X-ray diffraction), and again after a recurrence.")
STONE_BASIC_WHY = ("EAU: basic metabolic evaluation for every stone former – blood creatinine, calcium and uric acid; "
                   "urine dipstick, pH and sediment.")
STONE_HIGH_RISK_WHY = ("Only for high-risk stone formers (recurrent or early-onset stones, solitary kidney, bilateral "
                       "or large stones, uric acid, infection or cystine stones, bowel disease). EAU: two 24-hour urines "
                       "at least 20 days, ideally about 3 months, after the episode, on the usual diet.")


def n20_tests(v):
    out = []
    if v.get("_department") == "NMC":
        out.append(_s("us_k", "rec", "visit", "Check that the stone has passed and the kidney is not dilated. "
                                              "EAU follow-up imaging: ultrasound and/or KUB."))
        out.append(_s("kub", "consider", "visit", "For a radiopaque stone, KUB with ultrasound shows where the stone is."))
        out.append(_s("stone", "rec", "none", STONE_ANALYSIS_WHY + " Ask the patient to strain the urine and keep the stone.",
                      text="izdalītā konkrementa ķīmiskā sastāva analīze (urīnu filtrēt, konkrementu saglabāt)"))
        out.append(_s("stone_blood", "rec", "visit", STONE_BASIC_WHY))
        out.append(_s("ua", "rec", "visit", STONE_BASIC_WHY))
        out.append(_s("ncct", "consider", "symptoms", "EAU: keep CT for persistent symptoms or for planning an intervention."))
        out.append(_s("cbc", "consider", "visit", "Not needed for an uncomplicated stone (EAU); only if infection is suspected."))
    elif v.get("Admission") == "Acute":  # drained infected, obstructed kidney
        out.append(_s("uc", "rec", "definitive", "EAU (strong): delay stone treatment until the infection has resolved and "
                                                 "repeat the urine culture; culture before any intervention."))
        out.append(_s("creat", "rec" if v.get("Urosepsis") else "consider", "1w",
                      "Obstruction and sepsis can impair renal function: recheck after the drainage."))
        out.append(_s("cbc", "consider", "1w", PRACTICE_LABS))
        out.append(_s("crp", "consider", "1w", PRACTICE_LABS))
        out.append(_s("ncct", "consider", "definitive", "Only if new imaging is needed to plan the definitive treatment "
                                                        "(EAU: CT as preoperative imaging)."))
        out.append(_s("stone", "rec", "after_definitive", STONE_ANALYSIS_WHY))
        out.append(_s("stone_blood", "consider", "after_definitive", STONE_BASIC_WHY))
    else:  # planned endoscopic treatment
        out.append(_s("stone", "rec", "none", STONE_ANALYSIS_WHY + " Check the result of the fragments sent at surgery."))
        out.append(_s("stone_blood", "rec", "visit", STONE_BASIC_WHY))
        out.append(_s("ua", "rec", "visit", STONE_BASIC_WHY))
        out.append(_s("us_k", "rec", "visit", "Look for residual fragments and dilatation. EAU follow-up imaging: ultrasound "
                                              "and/or KUB; CT only if symptomatic or before a new intervention."))
        out.append(_s("kub", "consider", "visit", "For radiopaque stones, KUB together with ultrasound."))
        if v.get("DJ stent placed"):
            if v.get("Urinary tract infection") == "Yes":
                out.append(_s("uc", "rec", "stent", "Infection in this episode: urine culture before the stent is removed "
                                                    "(EAU: culture before urological procedures)."))
            else:
                out.append(_s("uc", "consider", "stent", "Before cystoscopic stent removal, if there are symptoms or a recent infection."))
        out.append(_s("ncct", "consider", "symptoms", "EAU: CT for symptomatic disease or before a new intervention."))
        out.append(_s("creat", "consider", "visit", "Both sides treated, solitary kidney or known renal impairment."))
    out.append(_s("urine24", "consider", "3m", STONE_HIGH_RISK_WHY))
    return out


def c67_tests(v):
    if v.get("C67 operation") == "Radical cystectomy":
        return c67_rc_tests(v)
    out = []
    known = v.get("Histology") == "Known"
    stage = v.get("Tumour stage") if known else None
    high_grade = v.get("Tumour grade") == "High grade (HG)"
    cis = stage == "pTis (CIS)" or bool(v.get("Concomitant CIS"))
    if stage == "pT2":
        out.append(_s("ct_stage", "rec", "asap", "Muscle-invasive tumour. EAU MIBC (strong): stage with CT of the chest, "
                                                 "abdomen and pelvis, including CT urography."))
        out.append(_s("creat", "rec", "asap", "Renal function before contrast CT and for treatment planning "
                                              "(e.g. cisplatin eligibility)."))
        out.append(_s("cysto", "avoid", "3m", "MIBC goes to the oncology council for radical treatment, not to the NMIBC "
                                              "cystoscopy schedule."))
        return out
    if stage == "Other / not malignant":
        out.append(_s("cysto", "consider", "3m", "Not a malignant tumour: further cystoscopy as the urologist decides."))
        return out
    high = known and (high_grade or stage == "pT1" or cis)
    out.append(_s("cysto", "rec", "3m", "EAU NMIBC (strong): the first cystoscopy 3 months after TURU, in all Ta/T1 "
                                        "tumours and CIS."))
    risk = nmibc_risk(v)
    if risk == "low":
        out.append(_s("cysto_low", "rec", "later", "EAU (weak): low risk – cystoscopy at 3 months, then 9 months later, "
                                                   "then yearly for 5 years."))
    elif risk == "intermediate" and v.get("Tumour grade") == "Low grade (LG)":
        out.append(_s("cysto_int", "rec", "later", "EAU (weak): intermediate-risk Ta LG – every 6 months for 2 years, "
                                                   "then yearly for 10 years."))
    if high:
        out.append(_s("cyto", "rec", "3m", "High grade, T1 or CIS: EAU follows these as high risk – cystoscopy with urine "
                                           "cytology at 3 months."))
        out.append(_s("cysto_later", "rec", "later", "EAU high-risk schedule (weak recommendation)."))
        out.append(_s("ctu_yearly", "rec", "later", "EAU (weak): yearly upper tract imaging (CT urography) in high and "
                                                    "very high risk."))
    elif known:
        out.append(_s("cyto", "avoid", "3m", "EAU: no routine cytology in low- or intermediate-risk (low-grade Ta) follow-up."))
    else:
        out.append(_s("cyto", "consider", "3m", "Depends on the histology: EAU adds cytology to cystoscopy for high-grade tumours."))
    if v.get("TURU type") == "TURU":
        if stage == "pT1":
            out.append(_s("returu", "rec", "2-6w", "EAU (strong): second resection for T1 tumours, within 2–6 weeks."))
        elif high_grade or not known:
            out.append(_s("returu", "consider", "2-6w", "EAU (strong): second resection if the first was incomplete or the "
                                                        "specimen has no detrusor muscle (not for Ta LG or primary CIS); "
                                                        "within 2–6 weeks."))
    if "Trigone" in (v.get("Tumour location") or []) or v.get("Tumour number") not in (None, "Single") or high:
        out.append(_s("ctu", "consider", "visit", "EAU (strong): baseline CT urography in selected cases – trigone, "
                                                  "multiple or high-risk tumours – if not done yet."))
    out.append(_s("creat", "consider", "visit", "Before contrast CT urography, if one is planned."))
    out.append(_s("us_kb", "avoid", "visit", "Not part of EAU follow-up when cystoscopy can be done."))
    return out


def nmc_cancer_tests(v):
    """Prostate or kidney cancer seen at NMC and sent home."""
    out = []
    if v.get("NMC reason") == "Urinary retention":
        out.append(_s("creat", "consider", "1w", "Retention can impair renal function: recheck if it was raised at NMC."))
        out.append(_s("us_kb", "consider", "visit", "If renal function is impaired: look for upper tract dilatation."))
    elif v.get("NMC reason") == "Hematuria":
        out.append(_s("cbc", "consider", "1w", "Check haemoglobin after the bleeding."))
    if v.get("Discharged with catheter"):
        out.append(_s("uc", "avoid", "symptoms", NO_CATHETER_UC))
    return out


def c61_tests(v):
    if v.get("_department") == "NMC":
        return nmc_cancer_tests(v)
    out = []
    if v.get("Operation") == "Radical prostatectomy":
        adverse = (v.get("Pathological T stage") in ("pT3b", "pT4") or v.get("Pathological N stage") == "pN1"
                   or v.get("Surgical margin") == "R1")
        why = "EAU (strong): follow up with PSA. PSA should be undetectable about 2 months after RP"
        why += ("; with adverse pathology an early PSA shows persistence (≥0.1 ng/ml)." if adverse else ".")
        out.append(_s("psa", "rec", "6w" if adverse else "3m", why))
        out.append(_s("psa_later", "rec", "later", "EAU: PSA every 6 months for 3 years, then yearly."))
        out.append(_s("continence", "consider", "3m", "EAU: record functional results with validated questionnaires."))
        out.append(_s("iief", "consider", "3m", "EAU: record erectile function with validated questionnaires."))
        out.append(_s("bone_pet", "avoid", "3m", "EAU: imaging has no place in routine follow-up while PSA is not rising."))
        out.append(_s("ct_abd", "avoid", "3m", "EAU: imaging has no place in routine follow-up while PSA is not rising."))
        out.append(_s("ua", "avoid", "visit", "Not part of EAU follow-up after RP; only with urinary symptoms."))
        out.append(_s("cbc", "avoid", "visit", "Not part of EAU follow-up after RP."))
        out.append(_s("panel", "avoid", "visit", "Your department's previous list. EAU follow-up after RP needs PSA "
                                                 "and a disease-specific history; the rest only if clinically needed."))
    elif v.get("Operation") == "TURP":
        why = "Checks the result of the TURP (EAU male LUTS follow-up: IPSS, uroflowmetry, residual urine)."
        out.append(_s("ipss", "consider", "at_visit", why))
        out.append(_s("uroflow", "consider", "at_visit", why))
        out.append(_s("pvr", "consider", "at_visit", why))
        out.append(_s("psa", "consider", "3m", "Prostate cancer follow-up continues with PSA; the TURP itself lowers PSA."))
    return out


RCC_FIRST_CT = {"Low": "6m", "Intermediate": "6m", "High": "3m"}  # EAU RCC follow-up table
RCC_LATER_CT = {
    "Low": "CT krūškurvim un vēdera dobumam 18 un 30 mēnešus pēc operācijas, pēc 3 gadiem – ik pēc 2 gadiem",
    "Intermediate": "CT krūškurvim un vēdera dobumam 12, 24 un 36 mēnešus pēc operācijas, pēc tam reizi gadā, "
                    "pēc 5 gadiem – ik pēc 2 gadiem",
    "High": "CT krūškurvim un vēdera dobumam 6, 12, 18, 24 un 36 mēnešus pēc operācijas, pēc tam reizi gadā, "
            "pēc 5 gadiem – ik pēc 2 gadiem",
}


def c64_tests(v):
    if v.get("_department") == "NMC":
        return nmc_cancer_tests(v)
    out = []
    egfr = v.get("eGFR at discharge")
    why = "EAU: follow renal function after kidney surgery"
    if egfr not in (None, "") and egfr < 45:
        why += "; eGFR is below 45, so EAU suggests a nephrology referral"
    out.append(_s("creat", "rec", "visit", why + "."))
    if v.get("Histology") == "Pending":
        out.append(_s("ct_ca", "rec", "histology", "The CT schedule depends on the recurrence risk (EAU), which needs "
                                                   "the histology."))
    elif v.get("RCC subtype") == "Benign":
        out.append(_s("ct_ca", "avoid", "6m", "Benign histology: no oncological CT follow-up."))
    elif v.get("Histology") == "Known":
        risk = v.get("Follow-up risk group") or "Intermediate"
        months = TIMINGS[RCC_FIRST_CT[risk]][0].lower()
        why = f"EAU {risk.lower()}-risk schedule: first CT of the chest and abdomen {months} (abdominal MRI is an alternative)."
        if v.get("Kidney operation") == "Partial nephrectomy" and v.get("RCC margin") == "R1":
            why += " Positive margin after partial nephrectomy: EAU (weak) suggests closer follow-up."
        out.append(_s("ct_ca", "rec", RCC_FIRST_CT[risk], why))
        out.append(_s("ct_later", "rec", "later", f"EAU {risk.lower()}-risk schedule.", text=RCC_LATER_CT[risk]))
    out.append(_s("cxr", "avoid", "6m", "EAU: chest X-ray and ultrasound are too insensitive for small RCC metastases."))
    out.append(_s("bone_pet", "avoid", "6m", "EAU: no routine bone scan or PET in follow-up."))
    return out


def n35_tests(v):
    out = []
    if v.get("_department") == "NMC":
        out.append(_s("rug_vcug", "rec", "definitive", "EAU: urethrography (retrograde, plus voiding through the "
                                                       "suprapubic catheter) shows the site and length of the stricture "
                                                       "before definitive treatment."))
        out.append(_s("urethrocysto", "consider", "definitive", "EAU: urethroscopy adds to urethrography in assessing the stricture."))
        out.append(_s("uc", "rec", "definitive", "EAU: screen for and treat bacteriuria before urethral surgery "
                                                 "(urine culture is better than urinalysis)."))
        raised = v.get("Creatinine raised (N35)") or (v.get("Drained volume (ml)") or 0) > 1000
        out.append(_s("creat", "rec" if raised else "consider", "1w",
                      "Large retention or raised creatinine: recheck renal function." if raised
                      else "Only if renal function was raised at NMC."))
        out.append(_s("us_kb", "consider", "visit", "After a large retention or raised creatinine: look for upper tract dilatation."))
        return out
    op = v.get("Stricture operation")
    if op in URETHROPLASTY:
        low_risk = (op == "Urethroplasty – excision and primary anastomosis"
                    and v.get("Stricture location") in ("Bulbar urethra", "Membranous (posterior) urethra")
                    and v.get("Stricture cause") != "Lichen sclerosus")
        if v.get("Catheter at discharge (N35)") == "In place":
            out.append(_s("rug_pericath", "rec", "catheter", "EAU (strong): urethrography before catheter removal to exclude "
                                                             "extravasation; the first one 7–10 days after an uncomplicated "
                                                             "urethroplasty (weak)."))
        out.append(_s("uroflow", "rec", "3m", "EAU: uroflowmetry at 3, 12 and 24 months"
                                              + ("." if low_risk else ", and at 5 years in standard risk.")))
        out.append(_s("proms", "rec", "3m", "EAU (strong): patient-reported outcome measures and a validated sexual "
                                            "function questionnaire."))
        out.append(_s("urethrocysto", "rec", "3m", "EAU (weak): flexible cystoscopy 3 months after urethroplasty "
                                                   "(retrograde urethrography is an alternative)."))
        if low_risk:
            out.append(_s("stricture_later", "rec", "later", "EAU low-risk schedule (anastomotic bulbar or membranous repair "
                                                            "without lichen sclerosus).",
                          text="12 un 24 mēnešus pēc operācijas – uroflowmetrija un simptomu anketas, cistoskopija vai "
                               "uretrogrāfija pēc indikācijām"))
        else:
            out.append(_s("stricture_later", "rec", "later", "EAU standard-risk schedule (penile, graft or other risk factors).",
                          text="12 un 24 mēnešus pēc operācijas – uroflowmetrija, simptomu anketas un cistoskopija vai "
                               "uretrogrāfija; uroflowmetrija un anketas arī pēc 5 gadiem"))
        out.append(_s("pvr", "consider", "3m", "EAU: residual urine helps only together with other tests, not on its own."))
    else:
        out.append(_s("uroflow", "rec", "3m", "EAU gives no schedule after dilation, urethrotomy or meatotomy; uroflowmetry "
                                              "with a symptom score is the usual check for an early recurrence."))
        out.append(_s("pvr", "consider", "3m", "Helpful together with uroflowmetry."))
        out.append(_s("proms", "consider", "3m", "A symptom questionnaire to follow the result."))
        out.append(_s("urethrocysto", "consider", "3m", "No EAU schedule after urethrotomy/dilation; a control "
                                                        "cystoscopy at 3 months is your department's practice – or only "
                                                        "if the stream weakens."))
    return out


def n40_tests(v):
    out = []
    why = "EAU male LUTS: review 4–6 weeks after catheter removal with IPSS, uroflowmetry and residual urine."
    out.append(_s("ipss", "rec", "at_visit", why))
    out.append(_s("uroflow", "rec", "at_visit", why))
    out.append(_s("pvr", "rec", "at_visit", why))
    out.append(_s("iief", "consider", "at_visit", "EAU also lists erectile and ejaculatory function at this review."))
    if "Upper tract dilatation / renal impairment" in (v.get("BPH indication") or []):
        out.append(_s("creat", "rec", "visit", "Renal impairment before surgery: check that it improves."))
        out.append(_s("us_k", "rec", "visit", "Upper tract dilatation before surgery: check that it has resolved."))
    result = (v.get("Histology result (BPH)") or "").lower()
    if v.get("Histology") == "Known" and any(w in result for w in ("karcinom", "adenokarcinom", "ca ", "gleason", "isup")):
        out.append(_s("psa", "rec", "3m", "Incidental prostate cancer: PSA follow-up and further assessment in the "
                                          "oncology council."))
        out.append(_s("mri_prostate", "consider", "3m", "Incidental prostate cancer: MRI if curative treatment is being considered."))
    else:
        out.append(_s("psa", "consider", "3m", "Not part of EAU follow-up after surgery. A new PSA baseline after TURP helps "
                                               "later prostate cancer testing, if that would change management."))
    out.append(_s("ua", "consider", "symptoms", "Only if burning or urgency do not settle (infection)."))
    out.append(_s("uc", "avoid", "at_visit", NO_ROUTINE_UC))
    return out


def n47_tests(v):
    out = []
    if v.get("Balanoposthitis") or (not v.get("Surgery") and v.get("Conservative reason") == "Balanoposthitis"):
        out.append(_s("glucose", "consider", "none", "Balanoposthitis, especially candidal or recurrent, can be the first "
                                                     "sign of diabetes (UK BASHH guidance)."))
    return out


def t83_tests(v):
    out = []
    if v.get("Urinary retention (diagnosis)") or (v.get("Drained volume (ml)") or 0) > 1000:
        out.append(_s("creat", "consider", "1w", "Retention or a large drained volume: check renal function."))
    if v.get("Urine appearance at change") in ("Blood-tinged", "With clots"):
        out.append(_s("us_kb", "consider", "visit", "Blood in the urine: look for clots, stones or a bladder lesion."))
    else:
        out.append(_s("us_kb", "consider", "visit", "If the catheter blocks again and again: look for bladder stones."))
    out.append(_s("cysto", "consider", "visit", "Repeated blockage (encrustation, bladder stones) or haematuria: cystoscopy "
                                                "at the urology review."))
    out.append(_s("uc", "avoid", "symptoms", NO_CATHETER_UC))
    return out


def neph_tests(v):
    if v.get("_department") in ("NMC", "4"):
        return nf_tests(v)
    out = [_s("creat", "consider", "next_change", "Renal function in an obstructed kidney – useful before the next change, "
                                                  "especially with malignant obstruction.")]
    if v.get("Nephrostomy function") == "Reduced":
        out.append(_s("us_k", "rec", "1w", "Reduced drainage: check the tube position and the dilatation."))
    else:
        out.append(_s("us_k", "consider", "next_change", "Only if drainage worsens or the tube position is in doubt."))
    out.append(_s("uc", "avoid", "symptoms", "Nephrostomy urine is usually colonised. EAU: do not treat asymptomatic "
                                             "bacteriuria; culture only with infection symptoms."))
    return out


TEST_ADVISORS = {"N10": n10_tests, "N39.0": n39_tests, "N20": n20_tests, "C67": c67_tests, "C61": c61_tests,
                 "C64": c64_tests, "N35": n35_tests, "N40": n40_tests, "N47": n47_tests, "T83": t83_tests,
                 "Z43.6": neph_tests}


def nf_tests(v):
    out = []
    blood = v.get("Blood tests (NF)")
    creat = blood in ("Raised creatinine", "Raised CRP and creatinine")
    crp = blood in ("Raised CRP / leukocytosis", "Raised CRP and creatinine")
    out.append(_s("creat", "rec" if (creat or v.get("Solitary kidney")) else "consider",
                  "48-72h" if creat else "1w", "Renal function after the obstruction is relieved (sooner if raised)."))
    if crp and v.get("_department") == "NMC":
        out.append(_s("crp", "rec", "48-72h", "Raised CRP at NMC: recheck to exclude a developing infection."))
        out.append(_s("cbc", "rec", "48-72h", "With the CRP recheck."))
    elif v.get("_department") == "4":
        out.append(_s("crp", "consider", "1w", PRACTICE_LABS))
    if v.get("Urine culture taken (NF)"):
        out.append(_s("uc", "rec", "none", "Check the culture and adjust the antibiotic if there are infection signs.",
                      text="paņemtā urīna uzsējuma rezultāts (Datamed.lv)"))
    else:
        out.append(_s("uc", "avoid", "symptoms", "Nephrostomy urine is usually colonised. EAU: do not treat "
                                                 "asymptomatic bacteriuria; culture only with infection symptoms."))
    out.append(_s("us_k", "consider", "1w", "If drainage falls or the tube position is in doubt."))
    return out


def n43_tests(v):
    out = []
    if not v.get("Scrotal ultrasound before surgery"):
        out.append(_s("us_scrotum", "rec", "visit", "No ultrasound before surgery: in adults a hydrocele can be reactive "
                                                    "to a testicular tumour, so the testis should be imaged."))
    elif v.get("Scrotal hematoma"):
        out.append(_s("us_scrotum", "consider", "1w", "Scrotal hematoma: recheck if it grows or does not settle."))
    else:
        out.append(_s("us_scrotum", "consider", "symptoms", "Only if the swelling does not settle or a recurrence is "
                                                            "suspected; routine imaging is not needed."))
    return out


TEST_ADVISORS["N43"] = n43_tests


def n13_tests(v):
    out = []
    stent = v.get("Pyeloplasty stent") == "DJ stent"
    out.append(_s("us_k", "rec", "after_stent" if stent else "4-6w",
                  "Common practice after pyeloplasty: ultrasound a few weeks after the stent is out, to check the "
                  "hydronephrosis is not getting worse (no adult EAU guideline on PUJ obstruction)."))
    out.append(_s("renogram", "rec", "6m", "Diuretic renography 3–6 months after surgery is the usual test of success "
                                           "(better drainage, stable or better split function)."))
    split = v.get("Split renal function (%)")
    out.append(_s("creat", "rec" if (split not in (None, "") and split < 40) else "consider", "visit",
                  "Renal function after surgery on an obstructed kidney; especially with reduced split function."))
    if stent:
        out.append(_s("uc", "consider", "stent", "Before cystoscopic stent removal if there are urinary symptoms or a "
                                                 "recent infection (EAU: treat bacteriuria before procedures that "
                                                 "breach the mucosa)."))
    out.append(_s("us_k", "consider", "later", "Later yearly ultrasound for 1–2 years: late failures do happen.",
                  text="nieru USS reizi gadā 1–2 gadus"))
    out.append(_s("ua", "avoid", "visit", NO_ROUTINE_UA))
    return out


TEST_ADVISORS["N13.0"] = n13_tests


def scrotal_tests(v):
    out = []
    if not v.get("Scrotal ultrasound before surgery"):
        out.append(_s("us_scrotum", "rec", "visit", "No ultrasound before surgery: image the testes and the operated "
                                                    "area (EAU: ultrasound for every scrotal mass)."))
    else:
        out.append(_s("us_scrotum", "consider", "symptoms", "Only if swelling or a lump returns; routine imaging is "
                                                            "not needed after excision of a benign lesion."))
    result = (v.get("Scrotal histology result") or "").lower()
    if v.get("Histology") == "Known" and any(w in result for w in MALIGNANT_WORDS):
        out.append(_s("ct_ca", "rec", "asap", "Malignant histology: staging as decided by the oncology council."))
    out.append(_s("panel", "avoid", "visit", "Blood tests are not needed after excision of a benign scrotal lesion."))
    out.append(_s("ua", "avoid", "visit", "Not needed after scrotal wall surgery."))
    return out


TEST_ADVISORS["D29.4"] = scrotal_tests


def n21_tests(v):
    out = []
    why = "EAU (weak): look for the cause of bladder stones – flow rate and residual urine."
    out.append(_s("uroflow", "rec", "at_visit", why))
    out.append(_s("pvr", "rec", "at_visit", why))
    if not v.get("Stone sent for analysis"):
        out.append(_s("stone", "consider", "none", "EAU: stone analysis for all first-time stone formers – if no "
                                                   "fragments were sent, analyse any that pass.",
                      text="izdalīto konkrementa fragmentu ķīmiskā sastāva analīze"))
    out.append(_s("ua", "rec", "visit", "EAU: urine dipstick and pH as part of the cause work-up."))
    out.append(_s("uc", "consider", "visit", "EAU: urine culture is optional in the work-up; take it if there were "
                                            "infection signs or the stone may be an infection stone."))
    persisting = set(v.get("Bladder stone cause") or []) & {"Neurogenic bladder", "Long-term catheter",
                                                            "Augmented bladder / reservoir"}
    out.append(_s("us_kb", "rec" if persisting else "consider", "6m",
                  "EAU (weak): individualise imaging follow-up; check for recurrence when the cause persists."))
    out.append(_s("stone_blood", "consider", "visit", "EAU: metabolic assessment is part of the work-up, mainly when "
                                                      "there is no clear obstructive cause."))
    if "Urethral stricture" not in (v.get("Bladder stone cause") or []):
        out.append(_s("urethrocysto", "consider", "visit", "EAU (weak): cystourethroscopy or urethrogram in selected "
                                                           "patients (suspected stricture or bladder neck problem)."))
    out.append(_s("us_k", "consider", "visit", "EAU (weak): upper tract imaging in selected patients."))
    return out


TEST_ADVISORS["N21.0"] = n21_tests


def r33_tests(v):
    out = []
    causes = v.get("Retention cause") or []
    raised = v.get("Labs at NMC") in ("Raised creatinine", "UTI and raised creatinine")
    out.append(_s("creat", "rec", "1w" if not raised else "48-72h",
                  "Retention can impair renal function; recheck creatinine and eGFR (sooner if raised at NMC)."))
    if v.get("Back pain or neurological symptoms"):
        out.append(_s("mri_spine", "rec", "asap", "Retention with back pain or neurological signs: urgent MRI to exclude "
                                                 "cauda equina syndrome."))
    if not v.get("Ultrasound at NMC"):
        out.append(_s("us_kb", "rec", "visit", "Upper tracts, bladder and prostate volume were not imaged at NMC."))
    if v.get("Labs at NMC") in ("Urinary tract infection", "UTI and raised creatinine"):
        out.append(_s("uc", "rec", "none", "Check the culture taken at NMC and adjust the antibiotic.",
                      text="NMC paņemtā urīna uzsējuma rezultāts (Datamed.lv)"))
    else:
        out.append(_s("uc", "avoid", "visit", NO_CATHETER_UC))
    if v.get("_sex") != "Woman" and any(c in causes for c in ("BPH", "Prostate cancer", "Not clear yet")):
        out.append(_s("psa", "consider", "6w", "Retention and the catheter raise PSA – measure it no earlier than about "
                                               "6 weeks, if a cancer diagnosis would change management."))
        out.append(_s("ipss", "consider", "at_visit", "EAU male LUTS assessment of the cause."))
    if v.get("Catheter plan (R33)") != "Long-term catheter":
        out.append(_s("uroflow", "rec", "at_visit", "After a successful trial without catheter: flow rate and residual "
                                                    "urine show whether the bladder empties."))
        out.append(_s("pvr", "rec", "at_visit", "After a successful trial without catheter: residual urine."))
    if "Urethral stricture" in causes:
        out.append(_s("urethrocysto", "rec", "visit", "Stricture as the cause: assess it before definitive treatment."))
    if "Clot retention (hematuria)" in causes:
        out.append(_s("ctu", "rec", "visit", "Hematuria work-up: CT urography and cystoscopy."))
        out.append(_s("cysto", "rec", "visit", "Hematuria work-up: cystoscopy."))
    if v.get("Labs at NMC") in ("Urinary tract infection", "UTI and raised creatinine"):
        out.append(_s("cbc", "consider", "1w", PRACTICE_LABS))
        out.append(_s("crp", "consider", "1w", PRACTICE_LABS))
    return out


TEST_ADVISORS["R33"] = r33_tests


def r31_tests(v):
    out, seen = [], set()

    def add(item):
        if item["id"] not in seen:
            seen.add(item["id"])
            out.append(item)

    causes = v.get("Hematuria cause") or []
    ward = v.get("_department") == "4"
    procs = (v.get("Hematuria procedure") or []) if ward else []
    done = ward and _hem_procs(v) is not None
    raised = v.get("Labs at NMC") in ("Raised creatinine", "UTI and raised creatinine")
    uti = "Urinary tract infection (hemorrhagic cystitis)" in causes or \
        v.get("Labs at NMC") in ("Urinary tract infection", "UTI and raised creatinine")
    if _hem_unknown(v) and not done:
        if v.get("Green corridor (GP referral)") == "Bladder – Z03.167":
            add(_s("cysto", "consider", "visit", "Already in the green corridor referral (Z03.167) – tick it only if you "
                                                 "want it written separately."))
        else:
            add(_s("cysto", "rec", "visit", "EAU (strong): cystoscopy for every patient with symptoms suggesting bladder "
                                            "cancer (visible hematuria) – flexible, outpatient; it cannot be replaced by "
                                            "cytology or other tests."))
    if "Radiation cystitis" in causes and not done:
        add(_s("cysto", "consider", "visit", "Radiation cystitis: cystoscopy (with biopsy if needed) excludes a recurrence "
                                             "or a second tumour."))
    if _hem_unknown(v) and v.get("CT at NMC (R31)") in (None, "Not done"):
        if raised:
            add(_s("us_kb", "rec", "visit", "Raised creatinine: ultrasound first; CT urography only after renal function "
                                            "is checked."))
            add(_s("ctu", "consider", "visit", "EAU (strong): ultrasound and/or CT urography for the upper tract; check "
                                               "eGFR before contrast."))
        else:
            add(_s("ctu", "rec", "visit", "EAU (strong): kidney and bladder ultrasound and/or CT urography in the initial "
                                          "work-up; CT urography is the most sensitive test of the upper tract."))
            add(_s("us_kb", "consider", "visit", "Alternative if CT urography is not possible (contrast allergy, "
                                                 "pregnancy, low eGFR)."))
    if _hem_unknown(v):
        add(_s("cyto", "consider", "visit", "EAU (strong): voided cytology as an adjunct to cystoscopy to find high-grade "
                                            "tumour / CIS (at least 25 ml fresh urine); most useful in smokers and "
                                            "high-risk patients."))
    if v.get("Hematuria type", "").startswith("Non-visible"):
        add(_s("ua", "rec", "visit", "Confirm microhematuria by microscopy (> 3 RBC/HPF); a dipstick alone is not enough "
                                     "(AUA, strong).", text="urīna analīze ar mikroskopiju (eritrocīti redzes laukā)"))
    if uti:
        add(_s("uc", "rec", "none", "Check the culture taken at NMC / in the ward and adjust the antibiotic.",
               text="paņemtā urīna uzsējuma rezultāts (Datamed.lv)"))
        add(_s("ua", "rec", "2-6w", "Repeat urinalysis after the infection is treated: hematuria that persists needs "
                                    "the full work-up.", text="atkārtota urīna analīze pēc infekcijas izārstēšanas"))
    hb = [x for x in (_num(v, "Hemoglobin (g/L)"), _num(v, "Hemoglobin at discharge (g/L)")) if x is not None]
    low = any(x < 100 for x in hb) or (_num(v, "Transfusion (RBC units)") or 0) > 0
    drugs = v.get("Antithrombotic therapy") or []
    add(_s("cbc", "rec" if (low or drugs) else "consider", "48-72h" if low else "1w",
           "Hb check after bleeding – sooner if Hb was low, after a transfusion, or on antithrombotic drugs."))
    if "Warfarin" in drugs:
        add(_s("inr", "rec", "48-72h", "Warfarin: check the INR, especially if the dose was paused or changed."))
    contrast = "Selective embolisation" in procs
    add(_s("creat", "rec" if (raised or contrast) else "consider", "48-72h" if (raised or contrast) else "1w",
           "Renal function: raised at NMC, after contrast (embolisation, CT), or before CT urography."))
    return out


TEST_ADVISORS["R31"] = r31_tests


def n30_tests(v):
    out = []
    woman, uncomplicated = _n30_flags(v)
    pregnant = woman and v.get("Pregnant (N30)")
    risks = v.get("Risk factors (N30)") or []
    recurrent = v.get("Recurrent cystitis")
    # (checking the NMC culture result is already written in the text when a culture was taken)
    if pregnant:
        out.append(_s("uc", "rec", "2w", "Pregnancy: follow-up culture 1–2 weeks after treatment (common practice; "
                                         "untreated bacteriuria raises the risk of pyelonephritis). EAU (strong): "
                                         "culture in pregnancy – take one at NMC if not done.",
                      text="kontroles urīna uzsējums 1–2 nedēļas pēc ārstēšanas"))
        out.append(_s("ua", "avoid", "visit", NO_ROUTINE_UA))
    else:
        out.append(_s("uc", "consider", "no_improve", "EAU (strong): culture if symptoms do not resolve or recur within "
                                                      "4 weeks – this is already written as advice in the text."))
        if "Visible hematuria" in (v.get("Cystitis symptoms") or []):
            out.append(_s("ua", "rec", "2-6w", "Hemorrhagic cystitis: check that the hematuria has gone after "
                                               "treatment; if it persists – cystoscopy (green corridor Z03.167 from 45 "
                                               "years).",
                          text="atkārtota urīna analīze pēc infekcijas izārstēšanas (hematūrijas kontrolei)"))
        else:
            out.append(_s("ua", "avoid", "visit", NO_ROUTINE_UA))
    if not woman:
        out.append(_s("pvr", "consider", "at_visit", "Men: check bladder emptying (residual urine)."))
        out.append(_s("uroflow", "consider", "at_visit", "Men: look for bladder outlet obstruction."))
        out.append(_s("psa", "avoid", "visit", "A UTI raises PSA – do not measure it now (wait at least 6 weeks)."))
    if recurrent:
        if woman and not risks and not pregnant:
            out.append(_s("cysto", "avoid", "visit", "EAU (weak): no routine cystoscopy or full abdominal ultrasound in "
                                                     "women under 40 with recurrent cystitis and no risk factors."))
            out.append(_s("us_kb", "consider", "visit", "Only in women over 40, with risk factors, or with residual "
                                                        "urine / hematuria."))
        else:
            out.append(_s("us_kb", "rec", "visit", "Recurrent UTI with risk factors or in a man: look for a cause "
                                                   "(residual urine, stones, obstruction)."))
    if recurrent and v.get("Same organism soon after treatment"):
        out.append(_s("ct_abd", "consider", "visit", "Bacterial persistence (Campbell-Walsh-Wein): CT to find a focus "
                                                     "(stone, obstruction, abscess)."))
        out = [o for o in out if o["id"] != "cysto"]
        out.append(_s("cysto", "consider", "visit", "Bacterial persistence: cystoscopy for a diverticulum, stone or "
                                                    "fistula if imaging does not explain it."))
    if v.get("Blood tests (N30)") == "Raised creatinine" or "Renal impairment" in risks:
        out.append(_s("creat", "rec", "1w", "Recheck renal function (and the antibiotic dose)."))
    out.append(_s("cbc", "avoid", "none", "Blood tests are not needed for uncomplicated cystitis."))
    return out


TEST_ADVISORS["N30.0"] = n30_tests


def c67_rc_tests(v):
    out = []
    diversion = v.get("Urinary diversion")
    bowel = diversion in ("Ileal conduit", "Orthotopic neobladder")
    palliative = "Palliative" in str(v.get("Cystectomy type (man)") or v.get("Cystectomy type (woman)") or "") or \
        str(v.get("Indication (RC)")).startswith("Palliative")
    transfused = (_num(v, "Postoperative transfusion (RC, units)") or 0) > 0 or v.get("Anaemia (RC)")
    out.append(_s("creat", "rec", "4-6w", "Renal function after RC: up to 29% progress to advanced CKD within a year "
                                          "(EAU follow-up)."))
    out.append(_s("nak", "rec", "4-6w", "Electrolytes after bowel diversion."))
    if bowel:
        out.append(_s("bgas", "rec", "4-6w", "Hyperchloraemic metabolic acidosis after bowel diversion (more with "
                                             "neobladders) harms kidneys and bone."))
        out.append(_s("b12", "rec", "12m", "EAU: measure vitamin B12 yearly after bowel diversion (low in ~17%)."))
    out.append(_s("cbc", "rec" if transfused else "consider", "1w" if transfused else "4-6w",
                  "Hb after the operation and transfusions."))
    stents = v.get("Ureteral stents (RC)") == "Discharged with stents" and diversion != "No diversion – nephrostomies"
    out.append(_s("us_k", "rec", "after_stent" if stents else "4-6w",
                  "Ureteroenteric stricture (up to 20%) or obstruction: kidney ultrasound after stent removal."))
    if not palliative:
        out.append(_s("ct_cap", "rec", "6m", "EAU follow-up: CT every 6 months until year 3, then yearly."))
        out.append(_s("ct_rc_later", "rec", "later", "EAU follow-up framework (expert opinion): CT every 6 months to "
                                                     "year 3, then yearly; CT urography for the upper tract in "
                                                     "multifocal disease, CIS or positive ureteral margins."))
    out.append(_s("cyto", "consider", "6m", "Not routine after RC. EAU / Campbell-Walsh-Wein: urethral wash cytology "
                                            "if the urethra is retained and there was CIS, prostatic urethra or "
                                            "multifocal disease."))
    out.append(_s("panel", "consider", "6m", "From your department's cystectomy discharge."))
    out.append(_s("uc", "avoid", "symptoms", "Conduit / neobladder urine is always colonised – culture only with "
                                             "infection symptoms (EAU: do not treat asymptomatic bacteriuria)."))
    return out


TEST_ADVISORS["C67"] = c67_tests


def n45_tests(v):
    out = []
    nmc = v.get("_department") == "NMC"
    abscess = _eo_abscess(v)
    raised = v.get("Labs (EO)") in ("Raised CRP / leukocytosis", "Raised CRP and creatinine")
    out.append(_s("crp", "rec" if raised else "consider", "48-72h" if nmc else "1w",
                  "Review at 3 days if not improving (IUSTI); your discharge example rechecks CBC, CRP, creatinine "
                  "and urea 5 days after discharge."))
    out.append(_s("cbc", "rec" if raised else "consider", "48-72h" if nmc else "1w", "With the CRP."))
    if v.get("Labs (EO)") == "Raised CRP and creatinine":
        out.append(_s("creat", "rec", "1w", "Creatinine was raised."))
    conservative_abscess = abscess and (nmc or v.get("Treatment (EO)") != "Surgery")
    out.append(_s("us_scrotum", "rec" if (conservative_abscess or v.get("Inflammation site (EO)") == "Orchitis")
                  else "consider", "4-6w",
                  "Persistent swelling: ultrasound for abscess, infarction or tumour (IUSTI 2024; Campbell-Walsh-Wein).",
                  text="kontroles sēklinieku maisiņa USS (izslēgt abscesu, infarktu, audzēju)"))
    out.append(_s("markers", "consider", "visit", "If a testicular mass persists after treatment: an 'orchitis' that "
                                                  "does not resolve may be a germ cell tumour – markers with ultrasound."))
    if _eo_sti(v):
        out.append(_s("naat", "rec", "4-6w", "IUSTI 2024: test of cure for chlamydia / M. genitalium about 4 weeks after "
                                             "treatment (gonorrhoea NAAT 2 weeks)."))
    if (v.get("Likely source (EO)") or "").startswith("Enteric"):
        out.append(_s("pvr", "consider", "at_visit", "Older men with enteric epididymitis: look for outlet obstruction "
                                                     "(Campbell-Walsh-Wein)."))
        out.append(_s("uroflow", "consider", "at_visit", "As above."))
    if "Diabetes" in (v.get("Risk factors (EO)") or []):
        out.append(_s("glucose", "consider", "1w", "Diabetes: glycaemic control affects healing and infection."))
    out.append(_s("uc", "avoid", "visit", NO_ROUTINE_UC))
    return out


TEST_ADVISORS["N45"] = n45_tests


def n41_tests(v):
    out = []
    nmc = v.get("_department") == "NMC"
    raised = v.get("Labs (AP)") in ("Raised CRP / leukocytosis", "Raised CRP and creatinine")
    out.append(_s("crp", "rec" if (raised and nmc) else "consider", "48-72h" if (raised and nmc) else "visit",
                  "Your NMC example repeats blood tests (CBC, CRP, creatinine, electrolytes) and urinalysis after "
                  "3 weeks, before the urology visit; sooner if the patient is not improving."))
    out.append(_s("cbc", "rec" if (raised and nmc) else "consider", "48-72h" if (raised and nmc) else "visit",
                  "With the CRP."))
    out.append(_s("creat", "rec" if v.get("Labs (AP)") == "Raised CRP and creatinine" else "consider", "visit",
                  "Renal function (raised at NMC, or before the visit)."))
    out.append(_s("nak", "consider", "visit", "From your NMC example (electrolytes)."))
    out.append(_s("ua", "consider", "visit", "From your NMC example: urinalysis before the urology visit."))
    out.append(_s("trus", "rec" if _ap_abscess(v) and v.get("Abscess treatment (AP)") == "Antibiotics only"
                  else "consider", "no_improve",
                  "EAU (weak): TRUS in selected cases to rule out an abscess – e.g. no improvement after 48–72 hours; "
                  "after conservative abscess treatment to confirm shrinkage."))
    out.append(_s("pvr", "rec", "at_visit", "After recovery: check bladder emptying (outlet obstruction is a "
                                            "complicating factor – EAU strong: manage it)."))
    out.append(_s("uroflow", "consider", "at_visit", "Outlet obstruction work-up after recovery."))
    if "Recurrent prostatitis" in (v.get("Risk factors (AP)") or []):
        out.append(_s("meares", "rec", "at_visit", "EAU (strong): Meares–Stamey test for suspected chronic bacterial "
                                                   "prostatitis – after the acute episode, off antibiotics."))
    if "STI risk (young, sexually active)" in (v.get("Risk factors (AP)") or []):
        out.append(_s("naat", "consider", "4-6w", "Test of cure if an STI was found."))
    if "Diabetes" in (v.get("Risk factors (AP)") or []):
        out.append(_s("glucose", "consider", "1w", "Diabetes: glycaemic control."))
    out.append(_s("psa", "avoid", "6w", "PSA rises with prostatitis – do not measure it now; if needed, at least 6–8 "
                                        "weeks after recovery."))
    out.append(_s("uc", "avoid", "visit", NO_ROUTINE_UC))
    return out


TEST_ADVISORS["N41.0"] = n41_tests


# ----- tables the doctor fills in (st.data_editor) -----

def antibiotics_editor(key):
    empty = pd.DataFrame({
        "Medication": pd.Series(dtype="object"),
        "Dose": pd.Series(dtype="object"),
        "Route": pd.Series(dtype="object"),
        "Frequency": pd.Series(dtype="object"),
        "Start date": pd.Series(dtype="datetime64[ns]"),
        "Continue at discharge": pd.Series(dtype="bool"),
        "End date": pd.Series(dtype="datetime64[ns]"),
    })
    edited = st.data_editor(
        empty, key=key, num_rows="dynamic", width="stretch",
        column_config={
            "Medication": st.column_config.SelectboxColumn(options=list(MEDICATION_CATALOG), required=True),
            "Dose": st.column_config.TextColumn(help="For example 500 mg"),
            "Route": st.column_config.SelectboxColumn(options=list(ROUTE_WORD), required=True),
            "Frequency": st.column_config.SelectboxColumn(options=FREQUENCIES),
            "Start date": st.column_config.DateColumn(format="DD.MM.YYYY", required=True),
            "Continue at discharge": st.column_config.CheckboxColumn(
                help="Tick the antibiotics that go home with the patient (they need dose, frequency and end date)"),
            "End date": st.column_config.DateColumn(format="DD.MM.YYYY"),
        },
    )
    return clean_rows(edited)


def labs_editor(key):
    empty = pd.DataFrame({
        "Test": pd.Series(dtype="object"),
        "Value": pd.Series(dtype="float64"),
        "Date": pd.Series(dtype="datetime64[ns]"),
    })
    edited = st.data_editor(
        empty, key=key, num_rows="dynamic", width="stretch",
        column_config={
            "Test": st.column_config.SelectboxColumn(
                options=list(LAB_TESTS), required=True,
                help="Units: CRO mg/l, leukocytes ×10⁹/l, creatinine µmol/l, eGFR ml/min/1.73 m²"),
            "Value": st.column_config.NumberColumn(required=True),
            "Date": st.column_config.DateColumn(format="DD.MM.YYYY", required=True),
        },
    )
    return clean_rows(edited)


def table_input(spec, key):
    rows = antibiotics_editor(key) if spec["table"] == "antibiotics" else labs_editor(key)
    if rows or spec.get("optional"):
        return rows
    return None


# ---------------------------------------------------------------
# 2b. DRAG AND DROP PANEL
# A small web page that Streamlit shows inside the page. Ctrl+click moves a block
# between "Available" and "Chosen". You do not need to edit this part.
# ---------------------------------------------------------------

BLOCK_PICKER_HTML = r"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  :root {
    --text: #31333f;
    --card: #f0f2f6;
    --accent: #ff4b4b;
    --border: rgba(128, 128, 128, 0.35);
  }
  html, body {
    margin: 0;
    padding: 0;
    background: transparent;
    color: var(--text);
    font-family: "Source Sans Pro", "Segoe UI", sans-serif;
  }
  body { padding: 2px 2px 6px 2px; }
  .box {
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 8px 10px;
    margin-bottom: 10px;
  }
  .box.over { outline: 2px dashed var(--accent); }
  .box-title { font-weight: 600; font-size: 15px; margin: 0 0 2px 2px; }
  .tip { font-size: 12px; opacity: 0.65; margin: 0 2px 4px 2px; }
  .hint { opacity: 0.6; font-size: 13px; padding: 10px 4px; text-align: center; }
  .group {
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    opacity: 0.65;
    margin: 10px 2px 2px 2px;
  }
  .card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 7px 10px;
    margin: 4px 0;
    cursor: grab;
    user-select: none;
    -webkit-user-select: none;
    font-size: 14px;
    line-height: 1.3;
  }
  .card:hover { border-color: var(--accent); }
  .card.avoid { border-left: 4px solid #d93025; }
  .card.dragging { opacity: 0.4; }
  .drop-line { height: 0; position: relative; }
  .drop-line::before {
    content: "";
    position: absolute;
    left: 0;
    right: 0;
    top: -2px;
    height: 3px;
    background: var(--accent);
    border-radius: 2px;
  }
</style>
</head>
<body>
<div class="box" id="chosenBox">
  <div class="box-title">Chosen</div>
  <div class="tip">Goes into the recommendations in this order. Ctrl+click a block to send it back.</div>
  <div id="chosenList"></div>
</div>
<div class="box" id="availBox">
  <div class="box-title">Available</div>
  <div class="tip">Drag a block up, or Ctrl+click it. Hover a block to read its notes.</div>
  <div id="availList"></div>
</div>
<script>
(function () {
  var items = [];
  var byId = {};
  var chosen = [];
  var signature = null;
  var dragId = null;

  var chosenBox = document.getElementById("chosenBox");
  var availBox = document.getElementById("availBox");
  var chosenList = document.getElementById("chosenList");
  var availList = document.getElementById("availList");

  function send(type, data) {
    var message = { isStreamlitMessage: true, type: type };
    for (var key in data) { message[key] = data[key]; }
    window.parent.postMessage(message, "*");
  }

  function setHeight() {
    send("streamlit:setFrameHeight", { height: document.body.scrollHeight + 8 });
  }

  function publish() {
    send("streamlit:setComponentValue", { value: { chosen: chosen.slice() }, dataType: "json" });
  }

  function applyTheme(theme) {
    if (!theme) { return; }
    var root = document.documentElement.style;
    if (theme.textColor) { root.setProperty("--text", theme.textColor); }
    if (theme.secondaryBackgroundColor) { root.setProperty("--card", theme.secondaryBackgroundColor); }
    if (theme.primaryColor) { root.setProperty("--accent", theme.primaryColor); }
    if (theme.font) { document.body.style.fontFamily = theme.font; }
  }

  function hint(text) {
    var el = document.createElement("div");
    el.className = "hint";
    el.textContent = text;
    return el;
  }

  function clearDropMarks() {
    chosenBox.classList.remove("over");
    availBox.classList.remove("over");
    var lines = document.querySelectorAll(".drop-line");
    for (var i = 0; i < lines.length; i++) { lines[i].parentNode.removeChild(lines[i]); }
  }

  function chosenCards() {
    return chosenList.querySelectorAll(".card:not(.dragging)");
  }

  function insertionIndex(clientY) {
    var cards = chosenCards();
    for (var i = 0; i < cards.length; i++) {
      var rect = cards[i].getBoundingClientRect();
      if (clientY < rect.top + rect.height / 2) { return i; }
    }
    return cards.length;
  }

  function showLine(index) {
    var old = document.querySelectorAll(".drop-line");
    for (var i = 0; i < old.length; i++) { old[i].parentNode.removeChild(old[i]); }
    var line = document.createElement("div");
    line.className = "drop-line";
    var cards = chosenCards();
    if (index < cards.length) { chosenList.insertBefore(line, cards[index]); }
    else { chosenList.appendChild(line); }
  }

  function toggle(id) {
    var index = chosen.indexOf(id);
    if (index >= 0) { chosen.splice(index, 1); } else { chosen.push(id); }
    render();
    publish();
  }

  function makeCard(item) {
    var card = document.createElement("div");
    card.className = "card" + (item.kind === "avoid" ? " avoid" : "");
    card.textContent = item.title;
    card.setAttribute("data-id", item.id);
    card.title = (item.note ? item.note + "\n\n" : "") + "Ctrl+click moves this block between the lists.";
    card.draggable = true;
    card.addEventListener("dragstart", function (event) {
      dragId = item.id;
      event.dataTransfer.effectAllowed = "move";
      try { event.dataTransfer.setData("text/plain", item.id); } catch (err) {}
      setTimeout(function () { card.classList.add("dragging"); }, 0);
    });
    card.addEventListener("dragend", function () {
      dragId = null;
      clearDropMarks();
      render();
    });
    card.addEventListener("click", function (event) {
      if (event.ctrlKey || event.metaKey) {
        event.preventDefault();
        toggle(item.id);
      }
    });
    return card;
  }

  function render() {
    chosenList.innerHTML = "";
    availList.innerHTML = "";

    if (chosen.length === 0) {
      chosenList.appendChild(hint("Nothing chosen yet. Drag a block here, or Ctrl+click it."));
    }
    chosen.forEach(function (id) { chosenList.appendChild(makeCard(byId[id])); });

    var lastGroup = null;
    var shown = 0;
    items.forEach(function (item) {
      if (chosen.indexOf(item.id) >= 0) { return; }
      if (item.group !== lastGroup) {
        var heading = document.createElement("div");
        heading.className = "group";
        heading.textContent = item.group;
        availList.appendChild(heading);
        lastGroup = item.group;
      }
      availList.appendChild(makeCard(item));
      shown += 1;
    });
    if (shown === 0) { availList.appendChild(hint("All blocks are chosen.")); }

    setHeight();
  }

  chosenBox.addEventListener("dragover", function (event) {
    if (dragId === null) { return; }
    event.preventDefault();
    if (event.dataTransfer) { event.dataTransfer.dropEffect = "move"; }
    chosenBox.classList.add("over");
    showLine(insertionIndex(event.clientY));
  });
  chosenBox.addEventListener("dragleave", function (event) {
    if (!chosenBox.contains(event.relatedTarget)) { clearDropMarks(); }
  });
  chosenBox.addEventListener("drop", function (event) {
    event.preventDefault();
    if (dragId === null) { return; }
    var id = dragId;
    var index = insertionIndex(event.clientY);
    var before = chosen.join("|");
    clearDropMarks();
    chosen = chosen.filter(function (x) { return x !== id; });
    chosen.splice(index, 0, id);
    dragId = null;
    render();
    if (chosen.join("|") !== before) { publish(); }
  });

  availBox.addEventListener("dragover", function (event) {
    if (dragId === null) { return; }
    event.preventDefault();
    if (event.dataTransfer) { event.dataTransfer.dropEffect = "move"; }
    availBox.classList.add("over");
  });
  availBox.addEventListener("dragleave", function (event) {
    if (!availBox.contains(event.relatedTarget)) { availBox.classList.remove("over"); }
  });
  availBox.addEventListener("drop", function (event) {
    event.preventDefault();
    if (dragId === null) { return; }
    var id = dragId;
    var before = chosen.length;
    clearDropMarks();
    chosen = chosen.filter(function (x) { return x !== id; });
    dragId = null;
    render();
    if (chosen.length !== before) { publish(); }
  });

  window.addEventListener("message", function (event) {
    var data = event.data;
    if (!data || data.type !== "streamlit:render") { return; }
    var args = data.args || {};
    applyTheme(data.theme);
    var newSignature = JSON.stringify(args.items || []);
    if (newSignature !== signature) {
      var firstTime = (signature === null);
      signature = newSignature;
      items = args.items || [];
      byId = {};
      items.forEach(function (item) { byId[item.id] = item; });
      var start = firstTime ? (args.initial || []) : chosen;
      chosen = start.filter(function (id) { return byId[id]; });
    }
    render();
  });

  window.addEventListener("resize", setHeight);
  send("streamlit:componentReady", { apiVersion: 1 });
})();
</script>
</body>
</html>
"""

# Streamlit needs the page as a file, so it is saved in a folder next to app.py.
PICKER_DIR = Path(__file__).parent / "block_picker_component_v2"
PICKER_DIR.mkdir(exist_ok=True)
_picker_file = PICKER_DIR / "index.html"
if not _picker_file.exists() or _picker_file.read_text(encoding="utf-8") != BLOCK_PICKER_HTML:
    _picker_file.write_text(BLOCK_PICKER_HTML, encoding="utf-8")
_block_picker = components.declare_component("block_picker_v2", path=str(PICKER_DIR))


def block_picker(code, blocks):
    """Show the panel for one diagnosis and return the ids of the chosen blocks, in order."""
    saved = st.session_state.get(f"chosen-{code}", [])
    result = _block_picker(
        items=[{"id": b["id"], "title": b["title"], "group": b["group"],
                "note": b["note"], "kind": b.get("kind", "")} for b in blocks],
        initial=saved,
        default={"chosen": saved},
        key=f"picker-{code}",
    )
    ids = result.get("chosen", []) if isinstance(result, dict) else []
    st.session_state[f"chosen-{code}"] = ids
    return ids


# ----- the follow-up tests box (one per diagnosis) -----

TEST_LEVELS = {"rec": ("🟢", "Recommended for this patient"), "consider": ("🟡", "Consider"),
               "avoid": ("⚪", "Not routinely needed")}
TIMING_KEYS = list(TIMINGS)


def _tests_tick(keys, value):
    for key in keys:
        st.session_state[key] = value


def _tests_clear(code, keys):
    _tests_tick(keys, False)
    st.session_state[f"{code}-tests-other"] = []
    st.session_state[f"{code}-tests-free"] = ""


def _timing_box(container, key, default):
    return container.selectbox("When", TIMING_KEYS, index=TIMING_KEYS.index(default),
                               format_func=lambda k: TIMINGS[k][0], key=key, label_visibility="collapsed")


def render_tests(code, v):
    """The follow-up tests box of one diagnosis. Nothing is ticked by default.
    Returns (ticked tests as [{"id", "text", "timing"}], who arranges them)."""
    advisor = TEST_ADVISORS.get(code)
    suggestions, seen = [], set()
    for s in (advisor(v) if advisor else []):
        if s["id"] not in seen:  # the first suggestion of a test wins
            seen.add(s["id"])
            suggestions.append(s)
    rank = {"rec": 0, "consider": 1, "avoid": 2}
    suggestions.sort(key=lambda s: rank[s["level"]])
    keys = {s["id"]: f"{code}-t-{s['id']}" for s in suggestions}
    rec_keys = [keys[s["id"]] for s in suggestions if s["level"] == "rec"]
    n_rec = len(rec_keys)
    n_consider = sum(s["level"] == "consider" for s in suggestions)

    title = "Follow-up tests (laboratory and imaging)"
    if suggestions:
        title += f" – {n_rec} recommended, {n_consider} to consider"
    chosen = []
    with st.expander(title, expanded=True):
        st.caption("Nothing is added by default: tick a test and it goes into the recommendations. The suggestions "
                   "follow the EAU guidelines (or the best available data) for the choices above and change with them.")
        c1, c2, c3 = st.columns([1, 1, 2], vertical_alignment="bottom")
        c1.button("Tick all 🟢", key=f"{code}-tests-all", on_click=_tests_tick, args=(rec_keys, True),
                  disabled=not rec_keys, width="stretch")
        c2.button("Clear", key=f"{code}-tests-none", on_click=_tests_clear, args=(code, list(keys.values())),
                  width="stretch")
        by = c3.selectbox("Who arranges them", list(TESTS_BY), key=f"{code}-tests-by")
        if not suggestions:
            st.caption("No tests are routinely needed in this situation. You can still add any test below.")

        level = None
        for number, s in enumerate(suggestions):
            if s["level"] != level:
                level = s["level"]
                icon, word = TEST_LEVELS[level]
                st.markdown(f"**{icon} {word}**")
            col_tick, col_when = st.columns([3, 2], vertical_alignment="center")
            if col_tick.checkbox(FOLLOWUP_TESTS[s["id"]]["title"], key=keys[s["id"]]):
                when = _timing_box(col_when, keys[s["id"]] + "-when", s["timing"])
                chosen.append({"id": s["id"], "text": s.get("text") or FOLLOWUP_TESTS[s["id"]]["lv"], "timing": when})
            # the reason, shown once under a run of tests that share it
            following = suggestions[number + 1] if number + 1 < len(suggestions) else None
            if not (following and following["why"] == s["why"] and following["level"] == level):
                st.caption(s["why"])

        st.markdown("**Other tests**")
        others = st.multiselect("Add any other test", list(FOLLOWUP_TESTS), key=f"{code}-tests-other",
                                format_func=lambda t: f"{FOLLOWUP_TESTS[t]['title']} ({FOLLOWUP_TESTS[t]['group'].lower()})",
                                placeholder="Search all tests…", label_visibility="collapsed")
        for test in others:
            if any(c["id"] == test for c in chosen):
                continue  # already ticked above
            col_name, col_when = st.columns([3, 2], vertical_alignment="center")
            col_name.markdown(FOLLOWUP_TESTS[test]["title"])
            when = _timing_box(col_when, f"{code}-to-{test}-when", "visit")
            chosen.append({"id": test, "text": FOLLOWUP_TESTS[test]["lv"], "timing": when})
        col_name, col_when = st.columns([3, 2], vertical_alignment="bottom")
        free = col_name.text_input("Another test, typed in Latvian", key=f"{code}-tests-free",
                                   placeholder="e.g. PTH, D vitamīns")
        if free.strip():
            when = _timing_box(col_when, f"{code}-tests-free-when", "visit")
            chosen.append({"id": "free", "text": free.strip().rstrip("."), "timing": when})
    return chosen, by


# ---------------------------------------------------------------
# 2j. OTHER DOCUMENTS: urologist's primary evaluation and decursus
# ---------------------------------------------------------------

# Presets for the reason of admission: operation, diagnosis after "iepriekš diagnosticētu ...", diagnosis.
# Presets for the primary evaluation:
# (admission, operation, reason after "iepriekš diagnosticētu ...", diagnosis, complaints, NMC course)
PE_PRESETS = {
    "TURP – complicated BPH": ("Planned", "TURP", "komplicētu LPH", "Komplicēta LPH", "", ""),
    "TURU – bladder tumour": ("Planned", "TURU", "urīnpūšļa jaunveidojumu", "Urīnpūšļa jaunveidojums", "", ""),
    "Re-TURU – bladder cancer after TURU": ("Planned", "atkārtota TURU (re-TURU)",
                                            "urīnpūšļa uroteliālo karcinomu, st. pēc TURU",
                                            "Urīnpūšļa CA, st. pēc TURU", "", ""),
    "Circumcision – phimosis": ("Planned", "cirkumcīzija", "fimozi", "Fimoze", "", ""),
    "Hydrocele plasty": ("Planned", "{adj} puses hidrocēles plastika", "{adj} puses hidrocēli",
                         "{Adj} puses hidrocēle", "", ""),
    "URS + DJ stent – ureteral stone": ("Planned", "{adj} puses ureterorenoskopija (URS), DJ stenta ievietošana",
                                        "{m} urīnvada konkrementu", "Nierakmeņu slimība. {M} urīnvada konkrements",
                                        "", ""),
    "FURS – kidney stone": ("Planned", "{adj} puses fleksibla ureterorenoskopija (FURS) ar kontakta lāzerlitotripsiju",
                            "{adj} nieres konkrementu", "Nierakmeņu slimība. {Adj} nieres konkrements", "", ""),
    "DJ stent removal – cystoscopy": ("Planned", "cistoskopija, {adj} puses DJ stenta evakuācija", "nierakmeņu slimību",
                                      "Nierakmeņu slimība. St. pēc {adj} puses DJ stenta ievietošanas", "", ""),
    "Laparoscopic radical prostatectomy – prostate cancer": ("Planned", "laparoskopiska radikāla prostatektomija",
                                                             "prostatas CA", "Prostatas CA", "", ""),
    "Nephrectomy – kidney tumour": ("Planned", "laparoskopiska {adj} puses nefrektomija", "{adj} nieres jaunveidojumu",
                                    "{Adj} nieres jaunveidojums", "", ""),
    "Laparoscopic kidney resection – kidney tumour": ("Planned", "laparoskopiska {adj} puses nieres rezekcija",
                                                      "{adj} nieres audzēju", "Ca renis {lat}", "", ""),
    "Scrotal mass – excision": ("Planned", "scrotum zemādas veidojuma ekscīzija",
                                "sēklinieka maisiņa zemādas veidojumu ({loc} pusē)",
                                "Scrotum zemādas veidojums ({loc} pusē)", "", ""),
    "Laparoscopic pyeloplasty – PUJ stricture": ("Planned", "{adj} puses laparoskopiska pieloplastika",
                                                 "{adj} puses pieloureterālā segmenta striktūru",
                                                 "{Adj} puses pieloureterālā segmenta striktūra", "", ""),
    "Cystolithotripsy – bladder stone": ("Planned", "cistolitotripsija", "urīnpūšļa konkrementu",
                                         "Urīnpūšļa konkrements", "", ""),
    "Optical urethrotomy – urethral stricture": ("Planned", "optiskā uretrotomija", "{part_acc}", "{part_dx}", "", ""),
    "Laser urethrotomy – urethral stricture": ("Planned", "lāzeruretrotomija", "{part_acc}", "{part_dx}", "", ""),
    "Urethral dilation – urethral stricture": ("Planned", "uretras dilatācija", "{part_acc}", "{part_dx}", "", ""),
    "Meatotomy – meatal stricture": ("Planned", "meatotomija", "uretras ārējās atveres (meatus) striktūru",
                                     "Meatus striktūra", "", ""),
    "Urethroplasty – urethral stricture": ("Planned", "uretroplastika", "{part_acc}", "{part_dx}", "", ""),
    "Acute bacterial prostatitis – acute admission (no operation)": (
        "Acute", "", "", "Komplicēta UCI. Akūts bakteriāls prostatīts",
        "sāpēm vēdera lejasdaļā, biežu, sāpīgu urināciju, paaugstinātu ķermeņa temperatūru",
        "NMC nodaļā vispusīgi izmeklēts. Konstatēta komplicēta UCI, akūts prostatīts."),
    "Prostatic abscess – drainage": (
        "Acute", "prostatas abscesa drenāža", "akūtu bakteriālu prostatītu ar prostatas abscesu",
        "Akūts bakteriāls prostatīts ar prostatas abscesu", "", ""),
    "Epididymo-orchitis – acute admission (conservative)": (
        "Acute", "", "", "{Adj} puses epididimoorhīts", "sāpēm un tūsku sēklinieku maisiņā, febrilu temperatūru",
        "Izmeklēts NMC. Sēklinieku maisiņa USS – epididimoorhīts. Uzsākta antibakteriāla terapija."),
    "Orchiectomy – purulent orchitis / testicular abscess": (
        "Acute", "{adj} puses orhektomija", "{adj} puses strutainu orhītu ar abscedēšanos",
        "{Adj} puses strutains orhīts ar abscedēšanos", "", ""),
    "Epididymectomy – epididymitis / epididymal abscess": (
        "Acute", "{adj} puses epididimektomija", "{adj} puses epididimītu", "{Adj} puses epididimīts", "", ""),
    "Radical cystoprostatectomy + ileal conduit – bladder cancer": (
        "Planned", "radikāla cistoprostatektomija, ileum conduit izveide, abpusēja iliakāla un obturatoru "
                   "limfadenektomija", "urīnpūšļa vēzi", "Ca vesicae urinariae", "", ""),
    "Radical cystectomy + hysterectomy + ileal conduit – bladder cancer (woman)": (
        "Planned", "radikāla cistektomija ar histerektomiju, piedēkļu un maksts priekšējās sienas rezekciju, ileum "
                   "conduit izveide, abpusēja iliakāla un obturatoru limfadenektomija", "urīnpūšļa vēzi",
        "Ca vesicae urinariae", "", ""),
    "Cystectomy – bladder cancer (nephrostomies, no diversion)": (
        "Planned", "cistektomija", "urīnpūšļa vēzi", "Ca vesicae urinariae. Abpusējas nefrostomas", "", ""),
    "Hematuria – acute admission (no operation)": (
        "Acute", "", "", "Makrohematūrija", "asins piejaukumu urīnā",
        "Izmeklēts NMC. Veic CT ar i/v kontrastvielu."),
    "Hematuria with bladder tamponade – acute admission (no operation)": (
        "Acute", "", "", "Makrohematūrija. Urīnpūšļa tamponāde", "asins piejaukumu urīnā",
        "Izmeklēts NMC – konstatēta makrohematūrija, urīnpūšļa tamponāde; veikta urīnpūšļa katetrizācija, uzsākta "
        "urīnpūšļa skalošana."),
    "Urinary retention – acute admission (no operation)": (
        "Acute", "", "", "Akūta urīna retence", "nespēju urinēt, sāpēm vēdera lejasdaļā",
        "Izmeklēts NMC. Sterilos apstākļos veikta urīnpūšļa katetrizācija."),
    "Acute admission, other (no operation)": ("Acute", "", "", "", "", "Izmeklēts NMC."),
    "Other (type it)": ("Planned", "", "", "", "", ""),
}
PE_PARTS = {  # urethral stricture location -> words for the presets with {part_...}
    "Penile (brīvā daļa)": {"part_acc": "uretras brīvās daļas striktūru", "part_dx": "Uretras brīvās daļas striktūra"},
    "Bulbar": {"part_acc": "uretras bulbārās daļas striktūru", "part_dx": "Uretras bulbārās daļas striktūra"},
    "Membranous": {"part_acc": "uretras membranozās daļas striktūru", "part_dx": "Uretras membranozās daļas striktūra"},
    "Not specified": {"part_acc": "uretras striktūru", "part_dx": "Uretras striktūra"},
}
PE_INDICATIONS = {  # default indication text for these presets (EAU: second resection within 2–6 weeks)
    "Radical cystoprostatectomy + ileal conduit – bladder cancer":
        "operabls augstas malignitātes pakāpes muskuli invadējošs urīnpūšļa vēzis; onkoloģiskā konsīlija lēmums – "
        "radikāla cistektomija",
    "Radical cystectomy + hysterectomy + ileal conduit – bladder cancer (woman)":
        "operabls augstas malignitātes pakāpes muskuli invadējošs urīnpūšļa vēzis; onkoloģiskā konsīlija lēmums – "
        "radikāla cistektomija",
    "Cystectomy – bladder cancer (nephrostomies, no diversion)":
        "urīnpūšļa vēzis; onkoloģiskā konsīlija lēmums – cistektomija",
    "Re-TURU – bladder cancer after TURU": "atkārtota rezekcija pēc TURU (pT1 audzējs / nepilnīga rezekcija / "
                                           "paraugā nav detruzora muskuļa)",
}
PE_DRE = {"Acute bacterial prostatitis – acute admission (no operation)": "prostata palielināta, izteikti sāpīga",
          "Prostatic abscess – drainage": "prostata palielināta, izteikti sāpīga, palpējama fluktuācija"}
PE_GENITALS = {"Hydrocele plasty": "{adj} puses hidrocēle",
               "Epididymo-orchitis – acute admission (conservative)": "{mm} sēklinieks un piedēklis palielināti, "
                                                                       "palpatori sāpīgi, sēklinieku maisiņa āda "
                                                                       "hiperemēta",
               "Orchiectomy – purulent orchitis / testicular abscess": "{mm} sēklinieks palielināts, pietūcis, "
                                                                        "palpatori sāpīgs",
               "Epididymectomy – epididymitis / epididymal abscess": "{mm} sēklinieka piedēklis palielināts, cietāks, "
                                                                      "palpatori sāpīgs",
               "Scrotal mass – excision": "scrotum zemādā {loc} pusē palpējams nesāpīgs, kustīgs veidojums, sēklinieki "
                                          "bez patoloģiskas atrades"}  # 'Ārējie dzimumorgāni' for these presets
PE_SIDE_WORDS = {"Right": {"adj": "labās", "Adj": "Labās", "lat": "dxt.", "m": "labā", "M": "Labā", "loc": "labajā",
                           "mm": "labais"},
                 "Left": {"adj": "kreisās", "Adj": "Kreisās", "lat": "sin.", "m": "kreisā", "M": "Kreisā",
                          "loc": "kreisajā", "mm": "kreisais"}}
ORDINALS = ["Pirmā", "Otrā", "Trešā", "Ceturtā", "Piektā", "Sestā", "Septītā", "Astotā", "Devītā", "Desmitā"]
DAY_TYPES = ["Iestāšanās diena", "Pirmsoperācijas diena", "Operācijas diena"] + \
            [f"{o} pēcoperācijas diena" for o in ORDINALS] + ["Izrakstīšanas diena"]
DAY_EXTRA = {"Operācijas diena": "Operācijas diena. Pēcoperācijas apskate."}


def sexed(text, sex):
    """Change male wording in a free-text sentence to female, if needed."""
    if sex != "Woman":
        return text
    for male, female in (("Pacientam ", "Pacientei "), ("Pacients ", "Paciente "), ("pacientam ", "pacientei "),
                         ("izrakstīts ", "izrakstīta "), ("Izmeklēts ", "Izmeklēta "), ("stacionēts ", "stacionēta ")):
        text = text.replace(male, female)
    return text


# ----- Word (.docx) versions, formatted like the department's documents -----

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
HDR_KEYS = ["hdr-patient", "hdr-history", "hdr-dept", "hdr-surgeons", "hdr-nurse", "hdr-anesth", "hdr-date-op",
            "hdr-date-pe", "hdr-date-dec"]


def _docx_new():
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.left_margin, section.right_margin = Cm(2.5), Cm(1.5)
    section.top_margin, section.bottom_margin = Cm(2), Cm(2)
    normal = doc.styles["Normal"]
    normal.font.name, normal.font.size = "Times New Roman", Pt(12)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    normal.paragraph_format.space_after = Pt(0)
    return doc


def _docx_bytes(doc):
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def _docx_line(doc, segments, indent=Cm(0.6), space_after=Pt(4)):
    """One line made of (label, value) pairs: the label bold and underlined, the value plain."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent, p.paragraph_format.space_after = indent, space_after
    for number, (label, value) in enumerate(segments):
        if number:
            p.add_run("  ")
        run = p.add_run(label)
        run.bold = run.underline = True
        if value:
            p.add_run(" " + value)
    return p


def _docx_page_header(doc, h, skip_first_page=False):
    """'Pacients … · Slimības vēstures nr. …' in small letters at the top right of every page."""
    parts = [f"Pacients: {h['patient']}" if h.get("patient") else "",
             f"Slimības vēstures nr.: {h['history']}" if h.get("history") else ""]
    text = "   ".join(p for p in parts if p)
    if not text:
        return
    section = doc.sections[0]
    section.different_first_page_header_footer = skip_first_page  # first page: it is already in the title
    p = section.header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(text)
    run.font.size = Pt(8)


def _docx_rule(doc):
    """A horizontal line (bottom border of an empty paragraph)."""
    p = doc.add_paragraph()
    border = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for key, value in (("w:val", "single"), ("w:sz", "18"), ("w:space", "1"), ("w:color", "000000")):
        bottom.set(qn(key), value)
    border.append(bottom)
    p._p.get_or_add_pPr().append(border)
    p.paragraph_format.space_after = Pt(10)


def _docx_text_block(doc, text, justify=False, indent=None):
    for line in text.split("\n"):
        p = doc.add_paragraph(line)
        if justify:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if indent is not None:
            p.paragraph_format.left_indent = indent


def operation_docx(h, name, paragraphs):
    """Like the department's 'Operācijas apraksts' file: header table, underlined title, justified text."""
    doc = _docx_new()
    rows = [("Nodaļa:", h.get("dept", "")), ("Pacients:", h.get("patient", "")), ("Vēstures nr:", h.get("history", "")),
            ("Operācija:", name), ("Datums:", h.get("date", "")), ("", ""),
            ("Ķirurgs/ Asistents:", h.get("surgeons", "")), ("Operācijas māsa:", h.get("nurse", ""))]
    if h.get("anesth"):
        rows.append(("Anesteziologs:", h["anesth"]))
    table = doc.add_table(rows=len(rows), cols=2)
    table.autofit = False
    table.columns[0].width, table.columns[1].width = Cm(4.5), Cm(12)
    for row, (label, value) in zip(table.rows, rows):
        row.cells[0].width, row.cells[1].width = Cm(4.5), Cm(12)
        row.cells[0].paragraphs[0].add_run(label)
        row.cells[1].paragraphs[0].add_run(value)
    for _ in range(3):
        doc.add_paragraph()
    title = doc.add_paragraph()
    title.add_run("Operācijas apraksts:").underline = True
    for text in paragraphs:
        p = doc.add_paragraph(text)
        if text.strip():
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(1.25)
            p.paragraph_format.line_spacing = 1.5
    return _docx_bytes(doc)


def decursus_docx(doc, decursus):
    for line in decursus.split("\n"):
        doc.add_paragraph(line)


def evaluation_docx(h, reason, history, objective, diagnosis, conclusion, decursus):
    """Like the department's 'Urologa apskate': title, date and patient, a line, bold underlined labels;
    the decursus follows on a new page."""
    doc = _docx_new()
    _docx_page_header(doc, h, skip_first_page=True)
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Urologa apskate")
    run.bold, run.font.size = True, Pt(14)
    _docx_line(doc, [("Datums", h.get("date", ""))])
    _docx_line(doc, [("Pacients", h.get("patient", ""))], indent=Cm(0))
    _docx_line(doc, [("Slimības vēstures nr.", h.get("history", ""))], indent=Cm(0))
    _docx_rule(doc)
    _docx_text_block(doc, reason, justify=True, indent=Cm(0.6))
    doc.add_paragraph()
    for heading, lines in (("Citas pārciestas slimības:", history), ("Objektīvais stāvoklis:", objective)):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent, p.paragraph_format.space_after = Cm(0.6), Pt(6)
        p.paragraph_format.space_before = Pt(6)
        run = p.add_run(heading)
        run.bold, run.font.size = True, Pt(14)
        for segments in lines:
            _docx_line(doc, segments)
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.paragraph_format.left_indent, p.paragraph_format.space_after = Cm(0.6), Pt(8)
    run = p.add_run("Diagnoze, Slēdziens:")
    run.bold, run.underline, run.font.size = True, True, Pt(13)
    p.add_run(" " + diagnosis).bold = True
    _docx_text_block(doc, conclusion, justify=True, indent=Cm(0.6))
    if decursus:
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        decursus_docx(doc, decursus)
    return _docx_bytes(doc)


def render_doc_header(kind):
    """Header fields that only go into the Word file. kind: "op", "pe" or "dec"."""
    st.session_state.setdefault("hdr-dept", "4.Uroloģijas")
    st.session_state.setdefault(f"hdr-date-{kind}", datetime.date.today())
    st.session_state.setdefault("hdr-surgeons", st.session_state.get("pe-surgeons", ""))
    st.session_state.setdefault("hdr-nurse", st.session_state.get("pe-nurse", ""))
    with st.expander("Word file header (optional)", expanded=False):
        st.caption("Used only in the Word file. On the online version leave the patient's name and history number "
                   "empty and type them in Word.")
        c1, c2 = st.columns(2)
        h = {"patient": c1.text_input("Pacients", key="hdr-patient").strip()}
        day = c2.date_input("Datums", format="DD.MM.YYYY", key=f"hdr-date-{kind}")
        h["date"] = fmt_date(day) + ("." if kind == "op" else "") if day else ""
        h["history"] = c1.text_input("Slimības vēstures nr.", key="hdr-history").strip()
        if kind == "op":
            h["dept"] = c2.text_input("Nodaļa", key="hdr-dept").strip()
            h["surgeons"] = c1.text_input("Ķirurgs/ Asistents", key="hdr-surgeons").strip()
            h["nurse"] = c2.text_input("Operācijas māsa", key="hdr-nurse").strip()
            h["anesth"] = c1.text_input("Anesteziologs (optional)", key="hdr-anesth").strip()
    return h


def word_button(label, data_maker, file_name, key):
    if not HAS_DOCX:
        st.info("To download a Word file, install python-docx once: py -3.13 -m pip install python-docx")
        return
    st.download_button(label, data=data_maker(), file_name=file_name, mime=DOCX_MIME, key=key,
                       icon=":material/download:")


def render_primary_evaluation(sex):
    f = SEX_FORMS[sex]
    st.subheader("Urologist's primary evaluation")
    with st.expander("1. Reason for admission", expanded=True):
        preset = st.selectbox("Preset", list(PE_PRESETS), key="pe-preset")
        adm_d, op_d, acc_d, dx_d, compl_d, course_d = PE_PRESETS[preset]
        side = None
        if "{part" in op_d + acc_d + dx_d:  # urethral stricture: the location words change
            side = st.radio("Stricture location", list(PE_PARTS), horizontal=True, key="pe-part")
            op_d, acc_d, dx_d = (t.format(**PE_PARTS[side]) for t in (op_d, acc_d, dx_d))
        elif "{" in op_d + acc_d + dx_d:  # this preset has a side: the side words change with it
            side = st.radio("Side", ["Left", "Right"], horizontal=True, key="pe-side")
            words = PE_SIDE_WORDS[side]
            op_d, acc_d, dx_d = (t.format(**words) for t in (op_d, acc_d, dx_d))
        tag = f"{preset}-{side}"  # new boxes (with the new side words) when the preset or the side changes
        admission = st.radio("Admission", ["Planned", "Acute"], index=0 if adm_d == "Planned" else 1,
                             horizontal=True, key=f"pe-admission-{preset}")
        operation = st.text_input("Operation (empty if no operation)", op_d, key=f"pe-op-{tag}")
        if operation.strip():
            dx_acc = st.text_input("Reason, as it follows 'sakarā ar iepriekš diagnosticētu ...'", acc_d,
                                   key=f"pe-dxacc-{tag}")
            complaints = st.text_input("Complaints (optional)", key=f"pe-compl-{preset}")
            course = ""
        else:
            dx_acc = ""
            complaints = st.text_input("Complaints, as they follow 'ar sūdzībām par ...'", compl_d,
                                       key=f"pe-compl-{preset}")
            course = st.text_area("Before admission (NMC, investigations, findings)", course_d,
                                  key=f"pe-course-{preset}", height=90)
        diagnosis = st.text_area("Diagnosis (Slēdziens)", dx_d, key=f"pe-dx-{tag}", height=70)
    with st.expander("2. Past diseases", expanded=True):
        other = st.text_area("Citas slimības (empty = Nav datu)", key="pe-other", height=70)
        c1, c2, c3 = st.columns(3)
        tbc = c1.radio("TBC", ["Noliedz", "Apstiprina", "Nav datu"], key="pe-tbc")
        hep = c2.radio("Hepatīti", ["Noliedz", "Apstiprina", "Nav datu"], key="pe-hep")
        allergy = c3.radio("Medikamentu nepanesība", ["Noliedz", "Ir", "Nav datu"], key="pe-allergy")
        allergy_text = st.text_input("Which medications", key="pe-allergy-text") if allergy == "Ir" else ""
        operations = st.text_area("Pārciestas operācijas (empty = Nav datu)", key="pe-ops", height=70)
    with st.expander("3. Objective status", expanded=True):
        c1, c2, c3 = st.columns(3)
        general = c1.selectbox("Vispārējais stāvoklis", ["Apmierinošs", "Vidēji smags", "Smags"], key="pe-general")
        conscious = c2.selectbox("Samaņa", ["Skaidra", "Apmiglota", "Sopors"], key="pe-conscious")
        tongue = c3.selectbox("Mēle", ["Valga, bez aplikuma", "Sausa, bez aplikuma", "Valga, ar aplikumu",
                                       "Sausa, ar aplikumu"], key="pe-tongue")
        edema = c1.selectbox("Tūskas", ["Nav", "Kāju distālajās daļās", "Ģeneralizētas"], key="pe-edema")
        nodes = c2.selectbox("Perifērie limfmezgli", ["Nepalpējas", "Palpējas palielināti"], key="pe-nodes")
        breath = c3.selectbox("Elpošana", ["Vezikulāra", "Novājināta", "Ar sēcieniem"], key="pe-breath")
        rr = c1.number_input("Elpošanas frekvence (x/min)", 5, 60, 16, key="pe-rr")
        hr = c2.number_input("Sirdsdarbība (x/min)", 20, 250, 70, key="pe-hr")
        rhythm = c3.selectbox("Ritms", ["Ritmiska", "Aritmiska"], key="pe-rhythm")
        bp = c1.text_input("TA (mm/Hg)", "124/78", key="pe-bp")
        abdomen = c2.text_input("Vēders", "Mīksts, neuzpūsts; nesāpīgs", key="pe-abdomen")
        peritoneal = c3.selectbox("Peritoneja kairinājuma simptomi", ["NEGATĪVI", "POZITĪVI"], key="pe-perit")
        kidneys = c1.selectbox("Nieres apklauvējot", ["Nesāpīgas", "Sāpīga labā", "Sāpīga kreisā", "Sāpīgas abas"],
                               key="pe-kidneys")
        bladder = c2.selectbox("Urīnpūslis", ["Nepalpējas", "Palpējas, pilns"], key="pe-bladder")
        genitals_d = PE_GENITALS.get(preset, "neizmainīti").format(**PE_SIDE_WORDS.get(side, PE_SIDE_WORDS["Left"]))
        genitals = c3.text_input("Ārējie dzimumorgāni", genitals_d, key=f"pe-genitals-{tag}")
        dre = st.text_input("DRI / per rectum (optional)", PE_DRE.get(preset, ""), key=f"pe-dre-{tag}")
    with st.expander("4. Conclusion and plan", expanded=True):
        if operation.strip():
            indication = st.text_input("Indication (empty = the diagnosis)", PE_INDICATIONS.get(preset, ""),
                                       key=f"pe-indication-{tag}",
                                       placeholder="e.g. ekskrēcijas traucējumi; operabls prostatas CA")
            consent = st.checkbox(f"{f['Pacients']} operācijai piekrīt", True, key="pe-consent")
            no_contra = st.checkbox("Kontrindikāciju nav", True, key="pe-contra")
            prepared = st.checkbox("Operācijai sagatavots", True, key="pe-prepared")
            surgeons = st.text_input("Operē (surgeons)", key="pe-surgeons")
            nurse = st.text_input("Op. māsa", key="pe-nurse")
            plan = ""
        else:
            plan = st.text_area("Plan", "Turpina ārstēšanos stacionārā!\nTerapija ordināciju lapā.", key="pe-plan",
                                height=80)
            actions = st.text_input("Done at admission (optional)",
                                    placeholder="e.g. Ievietots Foley 18 CH urīnpūšļa ilgkatetrs.", key="pe-actions")
            if actions.strip():
                plan += "\n" + actions.strip().rstrip(".") + "."
    header = render_doc_header("pe")
    add_decursus = st.checkbox("Add decursus (daily notes) to this evaluation", key="pe-add-dec")
    decursus = render_decursus_cards(sex) if add_decursus else ""

    if not diagnosis.strip() or (operation.strip() and not dx_acc.strip()) or (not operation.strip() and not complaints.strip()):
        st.info("Fill in the reason (or complaints) and the diagnosis to see the text.")
        return
    kind = "plānveida kārtā" if admission == "Planned" else "akūtā kārtā"
    if operation.strip():
        reason = (f"{f['Pacients']} {f['stacionets']} P. Stradiņa KUS 4. uroloģijas nodaļā {kind} operatīvai "
                  f"ārstēšanai – {operation.strip()} – sakarā ar iepriekš diagnosticētu {dx_acc.strip().rstrip('.')}.")
        if complaints.strip():
            reason += f" Sūdzības par {complaints.strip().rstrip('.')}."
    else:
        reason = f"{f['Pacients']} {f['stacionets']} {kind} ar sūdzībām par {complaints.strip().rstrip('.')}."
        if course.strip():
            reason += " " + sexed(course.strip(), sex)
            reason = reason if reason.endswith(".") else reason + "."
        reason += f" {f['stacionets'].capitalize()} 4. uroloģijas nodaļā."
    # Each line is a list of (label, value) pairs: plain text here, bold underlined labels in the Word file
    history = [
        [("Citas slimības:", other.strip() or "Nav datu")],
        [("TBC", tbc)],
        [("Hepatīti", hep)],
        [("Medikamentu nepanesība", allergy if allergy != "Ir" else (allergy_text.strip() or "Ir"))],
        [("Pārciestas operācijas:", operations.strip() or "Nav datu")],
    ]
    objective = [
        [("Vispārējais stāvoklis", general)],
        [("Samaņa:", conscious)],
        [("Mēle:", tongue)],
        [("Tūskas:", edema)],
        [("Perifērie limfmezgli", nodes)],
        [("Elpošana", f"{rr} x/min. {breath}")],
        [("Sirdsdarbība", f"{hr} x/min. {rhythm}; TA {bp.strip()} mm/Hg.")],
        [("Vēders:", abdomen.strip())],
        [("Peritoneja kairinājuma simptomi", peritoneal)],
        [("Nieres:", "Nepalpējas."), ("Apklauvējot:", kidneys.lower() if kidneys == "Nesāpīgas" else kidneys)],
        [("Urīnpūslis:", f"{bladder}.")],
        [("Ārējie dzimumorgāni", genitals.strip())],
    ] + ([[("DRI", dre.strip())]] if dre.strip() else [])
    as_text = lambda line: "  ".join(f"{label}  {value}" for label, value in line)
    dx_line = diagnosis.strip().rstrip(".") + "."
    lines = ([reason, "", "Citas pārciestas slimības:"] + [as_text(l) for l in history]
             + ["", "Objektīvais stāvoklis:"] + [as_text(l) for l in objective]
             + ["", f"Diagnoze, Slēdziens: {dx_line}"])
    if operation.strip():
        op = operation.strip()
        dx = diagnosis.strip().rstrip(".")
        parts = [f"{'Pacientei' if sex == 'Woman' else 'Pacientam'} indicēta {op}.",
                 f"Indikācijas operācijai – {(indication.strip().rstrip('.') or (dx if dx.startswith(('Ca ', 'CA ')) or dx[1:2].isupper() else dx[:1].lower() + dx[1:]))}."]
        if consent:
            parts.append(f"{f['Pacients']} operācijai piekrīt.")
        if no_contra:
            parts.append("Kontrindikāciju nav.")
        if prepared:
            parts.append("Operācijai " + ("sagatavota." if sex == "Woman" else "sagatavots."))
        parts.append(f"Operācija: {op}.")
        if surgeons.strip():
            parts.append(f"Operē: {surgeons.strip()}.")
        if nurse.strip():
            parts.append(f"Op. māsa: {nurse.strip()}.")
        parts.append("Skat. op. aprakstu.")
        conclusion = " ".join(parts)
    else:
        conclusion = plan.strip()
    lines += ["", conclusion]
    text = "\n".join(lines)
    if decursus:
        text += "\n\n" + decursus
    st.divider()
    st.text_area("Primary evaluation", text, height=650, key=f"pe-out-{hash(text)}")
    word_button("Word file (Urologa apskate" + (" + decursus" if decursus else "") + ")",
                lambda: evaluation_docx(header, reason, history, objective, dx_line, conclusion, decursus),
                f"Urologa_apskate_{header.get('date', '').replace('.', '-') or 'bez_datuma'}.docx", "pe-word")


# ----- decursus: one card per day, everything visible at once -----

DEC_FIELDS = {  # widget name -> default value of a new day
    "kind": "Apskate", "daytype": "(auto)", "status": "apmierinošs",
    "feeling": "Feels well, no complaints", "complaints": "",
    "pulse": 70, "rhythm": "ritmisks", "bp": "120/80", "resp": 16, "breath": "vezikulāra", "temp": 36.6,
    "abdomen": "Vēders palpatori mīksts, nesāpīgs.", "voiding": "", "notes": "",
    "cath_ml": None, "cath_look": "(no catheter)", "cath_out": False,
    "drain_ml": None, "drain_look": "(no drain)", "drain_out": False,
    "dressing": "(no dressing)",
    "ending": "Turpina ārstēšanu stacionārā.",
}
CATH_LOOKS = ["(no catheter)", "dzeltens", "dzidrs", "sarkanīgs", "sarkans", "viegli sarkans", "viegli sārts", "viegli sangvinozs", "sangvinozs", "ar recekļiem", "duļķains"]
DRAIN_LOOKS = ["(no drain)", "serozi", "serozi-sangvinozi", "sangvinozi-serozi", "serosangvinozi", "sangvinozi"]
DRESSINGS = {  # "Pārsiešana" choice -> text
    "(no dressing)": "",
    "Clean, healing": "Pārsiešana. Operācijas brūces tīras, sausas, dzīst primāri. Ādu apstrādā ar “Braunol”. "
                      "Sterili pārsēji.",
    "Clean (short)": "Pārsiešana. Pārsēji tīri, brūce tīra. Ādu apstrādā ar “Braunol”. Sterili pārsēji.",
    "Slight serous discharge": "Pārsiešana. Pārsējos neliels serozs izdalījums, brūces bez iekaisuma pazīmēm. Ādu "
                               "apstrādā ar “Braunol”. Sterili pārsēji.",
    "Slight bloody discharge": "Pārsiešana. Pārsējos neliels sangvinozs izdalījums, brūces bez iekaisuma pazīmēm. Ādu "
                               "apstrādā ar “Braunol”. Sterili pārsēji.",
    "Mild redness of the wound edges": "Pārsiešana. Brūces malās neliela hiperēmija, izdalījumu nav. Ādu apstrādā ar "
                                       "“Braunol”. Sterili pārsēji.",
}
FEELINGS = {  # complaints choice -> text ("Pacients" is changed for a woman)
    "Feels well, no complaints": "Pacients jūtas apmierinoši, sūdzību nav.",
    "Mild pain at the wounds": "Pacients jūtas apmierinoši, sūdzas par nelielu sāpīgumu operācijas brūču rajonā.",
    "Pain at the wounds": "Pacients sūdzas par sāpēm operācijas brūču rajonā, ko mazina pretsāpju terapija.",
    "Type my own": "",
}
DEC_ENDINGS = ["Turpina ārstēšanu stacionārā.", "Izrakstīts tālākai terapijai ambulatori.",
               "Izrakstīts ar urīnpūšļa ilgkatetru.", "(nothing)"]


def _dec_add_day():
    n = st.session_state.get("dd-count", 1)
    last = n - 1
    # copy the findings of the last examination day (not of a duty-staff entry)
    source = next((i for i in range(last, -1, -1) if st.session_state.get(f"dd{i}-kind", "Apskate") == "Apskate"), None)
    for field in DEC_FIELDS:
        value = st.session_state.get(f"dd{source}-{field}", DEC_FIELDS[field]) if source is not None else DEC_FIELDS[field]
        st.session_state[f"dd{n}-{field}"] = value
    if st.session_state.get(f"dd{n}-cath_out"):  # the catheter was removed on the last day
        st.session_state[f"dd{n}-cath_ml"], st.session_state[f"dd{n}-cath_look"] = None, "(no catheter)"
        st.session_state[f"dd{n}-cath_out"] = False
    if st.session_state.get(f"dd{n}-drain_out"):  # the drain was removed on the last day: no drain any more
        st.session_state[f"dd{n}-drain_ml"], st.session_state[f"dd{n}-drain_look"] = None, "(no drain)"
        st.session_state[f"dd{n}-drain_out"] = False
    st.session_state[f"dd{n}-kind"] = "Apskate"
    st.session_state[f"dd{n}-daytype"] = "(auto)"
    last_date = st.session_state.get(f"dd{last}-until") or st.session_state.get(f"dd{last}-date")
    st.session_state[f"dd{n}-date"] = (last_date + datetime.timedelta(days=1)) if last_date else None
    st.session_state[f"dd{n}-until"] = None
    st.session_state[f"dd{n}-ending"] = "Turpina ārstēšanu stacionārā."
    st.session_state["dd-count"] = n + 1


def _dec_remove_day():
    st.session_state["dd-count"] = max(1, st.session_state.get("dd-count", 1) - 1)


DEC_PRESETS = {  # typical findings filled into every day card (you can still change each day)
    "(none)": {},
    "TURU / re-TURU": {"cath_look": "viegli sārts", "drain_look": "(no drain)", "dressing": "(no dressing)",
                       "abdomen": "Vēders palpatori mīksts, nesāpīgs.",
                       "notes": "Urīnpūšļa skalošana nav nepieciešama."},
    "URS / FURS": {"cath_look": "(no catheter)", "drain_look": "(no drain)", "dressing": "(no dressing)",
                   "abdomen": "Vēders palpatori mīksts, nesāpīgs. Nieres apklauvējot nesāpīgas.",
                   "voiding": "Mikcija brīva, spontāna, urīns ar nelielu asiņu piejaukumu.", "notes": ""},
    "TURP": {"cath_look": "viegli sārts", "drain_look": "(no drain)", "dressing": "(no dressing)",
             "abdomen": "Vēders palpatori mīksts, nesāpīgs.", "notes": "Turpina urīnpūšļa skalošanu."},
    "Cystolithotripsy (3-way catheter)": {"cath_look": "sarkanīgs", "drain_look": "(no drain)",
                                          "dressing": "(no dressing)", "abdomen": "Vēders palpatori mīksts, nesāpīgs.",
                                          "voiding": "", "notes": "Urīnpūšļa skalošana nav nepieciešama."},
    "Hematuria (3-way catheter, irrigation)": {"cath_look": "viegli sarkans", "drain_look": "(no drain)",
                                               "dressing": "(no dressing)",
                                               "abdomen": "Vēders palpatori mīksts, nesāpīgs, vēdera izeja normāla.",
                                               "voiding": "", "notes": "Turpina urīnpūšļa skalošanu pa trīsceļu "
                                                                       "katetru."},
    "Urethrotomy / dilation (catheter)": {"cath_look": "dzeltens", "drain_look": "(no drain)",
                                          "dressing": "(no dressing)", "abdomen": "Vēders palpatori mīksts, nesāpīgs.",
                                          "voiding": "", "notes": "Urīnpūšļa ilgkatetrs funkcionē."},
    "Urethroplasty (catheter, perineal wound)": {"cath_look": "dzeltens", "drain_look": "(no drain)",
                                                 "dressing": "Clean (short)",
                                                 "abdomen": "Vēders palpatori mīksts, nesāpīgs.", "voiding": "",
                                                 "notes": "Urīnpūšļa ilgkatetrs funkcionē. Starpenes brūce bez "
                                                          "hematomas."},
    "Radical cystectomy + ileal conduit": {"cath_look": "(no catheter)", "drain_look": "serosangvinozi",
                                           "dressing": "Clean (short)",
                                           "abdomen": "Vēders palpatori mīksts, sāpīgs operācijas brūces rajonā, "
                                                      "peristaltika novājināta.",
                                           "voiding": "Pa urostomu dzeltens urīns adekvātā daudzumā.",
                                           "notes": "Uretera katetri funkcionē. Urostoma sārta, vitāla."},
    "Acute prostatitis (conservative)": {"cath_look": "(no catheter)", "drain_look": "(no drain)",
                                         "dressing": "(no dressing)",
                                         "abdomen": "Vēders palpatori mīksts, nesāpīgs, vēdera izeja normāla.",
                                         "voiding": "Urinācija spontāna, urīna daudzums adekvāts.",
                                         "notes": "Sāpes un vājums mazinās."},
    "Epididymo-orchitis (conservative)": {"cath_look": "(no catheter)", "drain_look": "(no drain)",
                                          "dressing": "(no dressing)",
                                          "abdomen": "Vēders palpatori mīksts, nesāpīgs, vēdera izeja normāla.",
                                          "voiding": "Urinācija spontāna, urīna daudzums adekvāts.",
                                          "notes": "Sēklinieku maisiņa tūska un jutīgums dinamikā mazinās."},
    "Cystectomy with nephrostomies": {"cath_look": "(no catheter)", "drain_look": "serosangvinozi",
                                      "dressing": "Clean (short)",
                                      "abdomen": "Vēders palpatori mīksts, sāpīgs operācijas brūces rajonā, "
                                                 "peristaltika novājināta.",
                                      "voiding": "Pa nefrostomām dzeltens urīns adekvātā daudzumā.", "notes": ""},
    "Scrotal operation": {"cath_look": "(no catheter)", "drain_look": "(no drain)", "dressing": "Clean (short)",
                          "abdomen": "Vēders palpatori mīksts, nesāpīgs.",
                          "voiding": "Mikcija brīva, spontāna.",
                          "notes": "Sēklinieka maisiņš bez izteiktas tūskas, hematomu nekonstatē."},
    "Laparoscopic operation": {"cath_look": "dzeltens", "drain_look": "serosangvinozi", "dressing": "Clean (short)",
                               "abdomen": "Vēders palpatori mīksts, nesāpīgs. Peristaltika izklausāma.", "notes": ""},
}


def _dec_apply_preset():
    values = DEC_PRESETS.get(st.session_state.get("dec-preset"), {})
    for i in range(st.session_state.get("dd-count", 1)):
        for field, value in values.items():
            st.session_state[f"dd{i}-{field}"] = value


def render_decursus_cards(sex):
    """The decursus as cards (one per day). Returns the decursus text."""
    st.session_state.setdefault("dd-count", 1)
    st.selectbox("Typical course (fills the day cards; you can still change each day)", list(DEC_PRESETS),
                 key="dec-preset", on_change=_dec_apply_preset)
    op_date = st.date_input("Operation date (optional, fills the day type)", value=None, format="DD.MM.YYYY",
                            key="dec-opdate")
    days = []
    for i in range(st.session_state["dd-count"]):
        for field, default in DEC_FIELDS.items():
            st.session_state.setdefault(f"dd{i}-{field}", default)
        st.session_state.setdefault(f"dd{i}-date", datetime.date.today() if i == 0 else None)
        with st.container(border=True):
            c = st.columns([2, 2, 3, 2])
            day = c[0].date_input(f"Day {i + 1}", format="DD.MM.YYYY", key=f"dd{i}-date")
            kind = c[1].selectbox("Entry", ["Apskate", "Dežūrpersonāla uzraudzībā"], key=f"dd{i}-kind")
            if kind == "Apskate":
                daytype = c[2].selectbox("Day type", ["(auto)", "(none)"] + DAY_TYPES, key=f"dd{i}-daytype")
                status = c[3].selectbox("Stāvoklis", ["apmierinošs", "vidēji smags", "smags"], key=f"dd{i}-status")
            else:
                until = c[2].date_input("Until", value=None, format="DD.MM.YYYY", key=f"dd{i}-until")
                notes = c[3].text_input("Note", key=f"dd{i}-notes")
                if day:
                    span = fmt_date(day) + (f"-{fmt_date(until)}" if until else "")
                    days.append(f"{span}\nDežūrpersonāla uzraudzībā." + (f" {notes.strip()}" if notes.strip() else ""))
                continue
            q = st.columns([2, 3])
            feeling = q[0].selectbox("Sūdzības", list(FEELINGS), key=f"dd{i}-feeling")
            complaints = q[1].text_input("Own text" if feeling == "Type my own" else "Extra complaints (optional)",
                                         key=f"dd{i}-complaints")
            v = st.columns(6)
            pulse = v[0].number_input("P", 20, 250, key=f"dd{i}-pulse")
            rhythm = v[1].selectbox("Ritms", ["ritmisks", "aritmisks"], key=f"dd{i}-rhythm")
            bp = v[2].text_input("TA", key=f"dd{i}-bp")
            resp = v[3].number_input("E", 5, 60, key=f"dd{i}-resp")
            breath = v[4].selectbox("Elpošana", ["vezikulāra", "novājināta", "ar sēcieniem"], key=f"dd{i}-breath")
            temp = v[5].number_input("T °C", 34.0, 42.0, step=0.1, format="%.1f", key=f"dd{i}-temp")
            a = st.columns(2)
            abdomen = a[0].text_input("Vēders", key=f"dd{i}-abdomen")
            voiding = a[1].text_input("Mikcija / urīns", key=f"dd{i}-voiding",
                                      placeholder="e.g. Spontāna mikcija, izdalās dzidrs urīns.")
            k = st.columns([2, 3, 2], vertical_alignment="bottom")
            cath_ml = k[0].number_input("Katetrs ml/dnn", 0, 10000, step=100, key=f"dd{i}-cath_ml")
            cath_look = k[1].selectbox("Urīns pa katetru", CATH_LOOKS, key=f"dd{i}-cath_look")
            cath_out = k[2].checkbox("Katetru evakuē", key=f"dd{i}-cath_out")
            k = st.columns([2, 3, 2], vertical_alignment="bottom")
            drain_ml = k[0].number_input("Drena ml/dnn", 0, 5000, step=10, key=f"dd{i}-drain_ml")
            drain_look = k[1].selectbox("Izdalījumi pa drenu", DRAIN_LOOKS, key=f"dd{i}-drain_look")
            drain_out = k[2].checkbox("Drenu evakuē", key=f"dd{i}-drain_out")
            dressing = st.selectbox("Pārsiešana", list(DRESSINGS), key=f"dd{i}-dressing")
            b = st.columns([3, 2])
            notes = b[0].text_input("Other", key=f"dd{i}-notes")
            ending = b[1].selectbox("Ending", DEC_ENDINGS, key=f"dd{i}-ending")
        if not day:
            continue
        if daytype == "(auto)" and op_date:
            diff = (day - op_date).days
            daytype = ("Operācijas diena" if diff == 0 else f"{ORDINALS[diff - 1]} pēcoperācijas diena"
                       if 0 < diff <= len(ORDINALS) else "Pirmsoperācijas diena" if diff < 0 else "(none)")
        parts = [fmt_date(day)]
        if daytype not in ("(auto)", "(none)"):
            parts.append(DAY_EXTRA.get(daytype, daytype + "."))
        said = [FEELINGS[feeling], complaints.strip()]
        said = " ".join(t if t.endswith(".") else t + "." for t in said if t)
        if said:
            parts.append(sexed(said, sex))
        parts.append(f"Obj. Vispārējais stāvoklis {status}.")
        parts.append(f"P {pulse:g} x min, {rhythm}, TA {bp.strip()} mmHg, E {resp:g} x min, {breath}, T {temp:.1f}°C.")
        tubes = []
        if cath_look != "(no catheter)":
            tubes.append("Pa urīnpūšļa ilgkatetru " + cath_look + " urīns"
                         + (f" ap {cath_ml:g} ml/dnn" if cath_ml is not None else "")
                         + (", katetru evakuē" if cath_out else ""))
        if drain_look != "(no drain)":
            tubes.append("Pa drenu " + drain_look + " izdalījumi" + (f" ap {drain_ml:g} ml/dnn" if drain_ml is not None else "")
                         + (", drenu evakuē" if drain_out else ""))
        for extra in [abdomen, voiding] + tubes + [DRESSINGS[dressing], notes]:
            if extra.strip():
                e = sexed(extra.strip(), sex)
                parts.append(e if e.endswith(".") else e + ".")
        if ending != "(nothing)":
            catheter = " ar funkcionējošu urīnpūšļa ilgkatetru" if "ilgkatetru" in ending else ""
            parts.append("Vispārējā apmierinošā stāvoklī " + ("izrakstīta" if sex == "Woman" else "izrakstīts")
                         + catheter + " tālākai terapijai ambulatori." if ending.startswith("Izrakstīts") else ending)
        days.append("\n".join(parts))
    b = st.columns(2)
    b[0].button("＋ Add next day (copies the last one)", on_click=_dec_add_day, key="dd-add", width="stretch")
    b[1].button("－ Remove last day", on_click=_dec_remove_day, key="dd-remove", width="stretch",
                disabled=st.session_state["dd-count"] <= 1)
    return "\n\n".join(days)


def render_decursus_table(sex):
    """Kept for compatibility: the decursus is now written as cards."""
    return render_decursus_cards(sex)


def render_decursus(sex):
    st.subheader("Decursus")
    left, right = st.columns([3, 2], gap="large")
    with left:
        text = render_decursus_cards(sex)
    with right:
        header = render_doc_header("dec")
        st.text_area("Decursus text", text, height=900, key=f"dec-out-{hash(text)}")
        if text:
            def make():
                doc = _docx_new()
                _docx_page_header(doc, header)  # name and number at the top right of every page
                decursus_docx(doc, text)
                return _docx_bytes(doc)
            word_button("Word file (decursus)", make, "Decursus.docx", "dec-word")


# ---------------------------------------------------------------
# 2k. OPERATION DESCRIPTIONS (Operācijas apraksts)
# ---------------------------------------------------------------

TURU_LOCATION_LOC = {  # where in the bladder (locative)
    "Left lateral wall": "kreisajā sānu sienā", "Right lateral wall": "labajā sānu sienā",
    "Posterior wall": "mugurējā sienā", "Anterior wall": "priekšējā sienā", "Dome": "jumtā",
    "Trigone": "trīsstūrī", "Bladder neck": "kakliņā",
}


# ----- URS and FURS, assembled from what was actually done -----

URS_LOCATIONS = {  # where in the ureter (locative)
    "Lower third": "apakšējā trešdaļā", "Middle third": "vidējā trešdaļā", "Upper third": "augšējā trešdaļā",
    "Intramural part": "intramurālajā daļā", "Pelviureteric junction": "pieloureterālajā segmentā",
}
URS_REACH = ["To the stone", "Whole ureter, up to the renal pelvis", "Not possible (narrow ureter)",
             "No URS – stent only"]
URS_FINDINGS = ["Stone", "No stone (passed)", "Stricture", "Oedema / inflamed mucosa"]
URS_ACTIONS = ["Laser lithotripsy", "Pneumatic lithotripsy", "Basket extraction", "Loop (cilpa) extraction",
               "Stone migrated to the kidney"]
URS_RESULTS = {"Stone-free": "Kontrolē atlieku konkrementus nekonstatē.",
               "Small fragments (< 2 mm)": "Kontrolē saglabājas sīki konkrementa fragmenti (< 2 mm).",
               "Stone remains": "Konkrementu pilnībā evakuēt neizdodas."}
FURS_PLACES = {"Renal pelvis": "nieres bļodiņā", "Upper calyx group": "augšējā kausiņu grupā",
               "Middle calyx group": "vidējā kausiņu grupā", "Lower calyx group": "apakšējā kausiņu grupā"}
FURS_ACTIONS = ["Laser lithotripsy", "Basket extraction", "Move lower-pole stone up (basket)"]
LASER_MODE = {"Fragmenting": "fragmentācijas režīmā", "Dusting": "pulverizācijas (dusting) režīmā"}


def cystolithotripsy_operation(c1, c2, c3, pt_acc):
    """Transurethral cystolithotripsy (department format). Returns (paragraphs, name)."""
    st.caption("EAU: transurethral cystolithotripsy where possible (strong), preferably with a continuous-flow "
               "instrument (nephroscope or resectoscope, weak); laser, pneumatic and mechanical lithotripsy are "
               "equivalent. Send the stone for analysis.")
    pt_nom = "Paciente" if pt_acc == "Pacienti" else "Pacients"
    anesthesia = c1.selectbox("Anestēzija", ["Intubācijas narkozē", "Spinālā anestēzijā"], key="op-cl-an")
    scope = c2.selectbox("Instruments", ["cistoskopu", "nefroskopu (nepārtrauktas plūsmas)",
                                         "rezektoskopu (nepārtrauktas plūsmas)"], key="op-cl-scope")
    energy = c3.selectbox("Litotripsija", ["lāzers (Ho:YAG)", "pneimatiskā", "mehāniskais litotriptors"],
                          key="op-cl-energy")
    mucosa = c1.selectbox("Gļotāda", ["neizmainīta", "hiperemēta", "trabekulāra", "trabekulāra, ar pseidodivertikuliem"],
                          key="op-cl-mucosa")
    urine = c2.selectbox("Urīns", ["dzidrs", "duļķains"], key="op-cl-urine")
    many = c3.radio("Konkrementi", ["Viens", "Vairāki"], horizontal=True, key="op-cl-many") == "Vairāki"
    size = c1.number_input("Izmērs (cm)", 0.0, 15.0, value=None, step=0.5, key="op-cl-size")
    shape = c2.selectbox("Forma", ["ovālas formas", "apaļas formas", "neregulāras formas"], key="op-cl-shape")
    fr = c3.selectbox("3-kanālu katetrs (Fr)", ["18", "20", "22", "24"], index=2, key="op-cl-fr")
    balloon = c1.number_input("Balons (ml)", 5, 60, 15, key="op-cl-balloon")
    colour = c2.selectbox("Urīns pa katetru", ["viegli rozīgas krāsas", "dzidrs", "sārtas krāsas"], key="op-cl-colour")
    ward = c3.text_input("Pārved uz", "4. nodaļu", key="op-cl-ward")
    prostate = st.text_input("Papildu atrade (optional)", key="op-cl-extra",
                             placeholder="e.g. Prostatas vidējā daiva izvirzās urīnpūslī.")
    codes = st.checkbox("Manipulāciju kodi tekstā", True, key="op-cl-codes")
    k1, k2 = st.columns(2)
    code_scope = k1.text_input("Cistoskopijas kods", "19161", key="op-cl-code1") if codes else ""
    code_laser = k2.text_input("Litotripsijas kods", "19175", key="op-cl-code2") if codes else ""
    tag = lambda code: f" ({code.strip()})" if codes and code.strip() else ""
    laser = energy.startswith("lāzers")
    size_text = (f" ~{size:g} cm lielu" if size else "").replace(".", ",")
    if many:
        stones = f"vairākus{(' līdz' + size_text.replace(' lielu', ' lielus')) if size else ''} konkrementus"
    else:
        stones = f"vienu{size_text} {shape} konkrementu"
    if laser:
        litho = (f"Caur {scope.split(' ')[0]} ievada lāzera stīgu{tag(code_laser)}. Vizuālā kontrolē veic "
                 f"{'konkrementu' if many else 'konkrementa'} fragmentāciju līdz sīkām šķembām. Evakuē lāzera stīgu.")
    elif energy == "pneimatiskā":
        litho = (f"Caur {scope.split(' ')[0]} ievada pneimatiskā litotriptora zondi{tag(code_laser)}. Vizuālā kontrolē "
                 f"veic {'konkrementu' if many else 'konkrementa'} fragmentāciju līdz sīkām šķembām. Evakuē zondi.")
    else:
        litho = (f"Ievada mehānisko litotriptoru{tag(code_laser)} un vizuālā kontrolē fragmentē "
                 f"{'konkrementus' if many else 'konkrementu'} līdz sīkām šķembām.")
    text = [
        f"{pt_nom} atrodas dorsālās litotomijas guļā. {anesthesia}, pēc operācijas lauka apdares ar polispirtu "
        f"šķīdumu urīnpūslī ievada {scope.split(' (')[0]}{tag(code_scope)}. Instrumenta ievadīšana brīva. "
        f"Urīnpūšļa gļotāda {mucosa}, urīns {urine}. Urīnpūšļa lūmenā vizualizē {stones}."
        + (f" {prostate.strip().rstrip('.')}." if prostate.strip() else "")
        + f" {litho} Atskalo šķembas; kontroles cistoskopijā atlieku konkrementus nekonstatē. Hemostāze apmierinoša. "
        "Evakuē instrumentu.",
        f"Operāciju nobeidzot urīnpūslī ievada {fr} Fr izmēra 3-kanālu Foley katetru, balons uzpildīts līdz "
        f"{balloon} ml. Pa katetru {colour if colour == 'dzidrs' else colour + ' urīns'}"
        + (" urīns" if colour == "dzidrs" else "") + ".",
        f"{pt_nom} vispārēji apmierinošā stāvoklī pārved uz {ward.strip()}. Konkrementa fragmentus nosūta ķīmiskā "
        "sastāva analīzei.",
    ]
    return text, "Cistoskopija, cistolitotripsija"


def hematuria_operation(c1, c2, c3, pt_acc):
    """Cystoscopy with clot evacuation and coagulation (hematuria / clot retention). Returns (paragraphs, name)."""
    st.caption("Evacuate all clots first (Ellik evacuator or bladder syringe), find the source, coagulate the bleeding "
               "points, take a biopsy of suspicious areas to exclude a tumour, and finish with a large 3-way catheter "
               "and irrigation. If a tumour is resected, use the TURU description.")
    pt_nom = "Paciente" if pt_acc == "Pacienti" else "Pacients"
    anesthesia = c1.selectbox("Anestēzija", ["Intubācijas narkozē", "Spinālā anestēzijā"], key="op-hc-an")
    scope = c2.selectbox("Instruments", ["rezektoskopu (nepārtrauktas plūsmas)", "cistoskopu"], key="op-hc-scope")
    evac = c3.selectbox("Recekļu evakuācija", ["Ellika evakuatoru", "Žanē šļirci"], key="op-hc-evac")
    clots = c1.number_input("Recekļi (ml, optional)", 0, 3000, value=None, step=50, key="op-hc-ml")
    mucosa = c2.selectbox("Gļotāda pēc evakuācijas", ["hiperemēta, ar staru cistīta izmaiņām (teleangiektāzijas)",
                                                       "hiperemēta", "trabekulāra", "neizmainīta"], key="op-hc-mucosa")
    source = c3.selectbox("Asiņošanas avots", ["Difūza gļotādas asiņošana", "Atsevišķi asiņojoši asinsvadi", "Prostata",
                                               "Urīnpūšļa jaunveidojums", "Labā urīnvada atvere",
                                               "Kreisā urīnvada atvere", "Avots nav atrasts"], key="op-hc-source")
    energy = c1.selectbox("Koagulācija", ["bipolāru", "monopolāru"], key="op-hc-energy")
    biopsy = c2.checkbox("Biopsija no aizdomīgām vietām", key="op-hc-biopsy")
    fr = c3.selectbox("3-kanālu katetrs (Fr)", ["20", "22", "24"], index=1, key="op-hc-fr")
    balloon = c1.number_input("Balons (ml)", 5, 60, 20, key="op-hc-balloon")
    irrigation = c2.checkbox("Nepārtraukta skalošana", True, key="op-hc-irr")
    colour = c3.selectbox("Urīns pa katetru", ["viegli rozīgas krāsas", "dzidrs", "sārtas krāsas"], key="op-hc-colour")
    ward = st.text_input("Pārved uz", "4. nodaļu", key="op-hc-ward")
    codes = st.checkbox("Manipulāciju kodi tekstā", True, key="op-hc-codes")
    k1, k2 = st.columns(2)
    code_scope = k1.text_input("Cistoskopijas kods", "19161", key="op-hc-code1") if codes else ""
    code_coag = k2.text_input("Koagulācijas kods (optional)", key="op-hc-code2") if codes else ""
    tag = lambda code: f" ({code.strip()})" if codes and code.strip() else ""
    coag = f"{energy} koagulāciju{tag(code_coag)}"
    orifice = source.endswith("urīnvada atvere")
    found = {
        "Difūza gļotādas asiņošana": f"Vizualizē difūzu gļotādas asiņošanu; veic asiņojošo vietu {coag}.",
        "Atsevišķi asiņojoši asinsvadi": f"Vizualizē atsevišķus asiņojošus asinsvadus; veic to {coag}.",
        "Prostata": f"Asiņošanas avots – asinsvadi prostatas virsmā; veic to {coag}.",
        "Urīnpūšļa jaunveidojums": f"Asiņošanas avots – urīnpūšļa jaunveidojums; veic tā virsmas {coag}.",
        "Avots nav atrasts": "Aktīvu asiņošanas avotu nekonstatē.",
    }.get(source) or (f"No {'labās' if source.startswith('Labā') else 'kreisās'} urīnvada atveres izdalās asiņains "
                      "urīns – asiņošanas avots augšējos urīnceļos; urīnpūšļa gļotādā aktīvas asiņošanas vietas "
                      "nekonstatē.")
    amount = f" (~{clots:g} ml)" if clots else ""
    end = (" Urīnpūslis atskalots, recekļu atliekas nekonstatē." if (orifice or source == "Avots nav atrasts")
           else " Kontroles cistoskopijā hemostāze apmierinoša, recekļu atliekas nekonstatē.")
    text = [
        f"{pt_nom} atrodas dorsālās litotomijas guļā. {anesthesia}, pēc operācijas lauka apdares ar polispirtu "
        f"šķīdumu urīnpūslī ievada {scope.split(' (')[0]}{tag(code_scope)}. Instrumenta ievadīšana brīva. Urīnpūslī "
        f"liels daudzums asins recekļu{amount}, kurus fragmentē un evakuē ar {evac}. Pēc recekļu evakuācijas "
        f"urīnpūšļa gļotāda {mucosa}. {found}"
        + (" No aizdomīgām gļotādas vietām paņem biopsiju." if biopsy else "") + end + " Evakuē instrumentu.",
        f"Operāciju nobeidzot urīnpūslī ievada {fr} Fr izmēra 3-kanālu Foley katetru, balons uzpildīts līdz "
        f"{balloon} ml." + (" Uzsāk nepārtrauktu urīnpūšļa skalošanu." if irrigation else "")
        + f" Pa katetru {colour if colour == 'dzidrs' else colour + ' urīns'}" + (" urīns" if colour == "dzidrs" else "")
        + ".",
        f"{pt_nom} vispārēji apmierinošā stāvoklī pārved uz {ward.strip()}."
        + (" Biopsijas materiālu nosūta histoloģiskai izmeklēšanai." if biopsy else ""),
    ]
    coagulated = not (orifice or source == "Avots nav atrasts")
    name = "Cistoskopija, recekļu evakuācija" + (", koagulācija" if coagulated else "") + (", biopsija" if biopsy else "")
    return text, name


def cystectomy_operation(woman, c1, c2, c3, pt_dat, pt_acc):
    """Radical cystectomy (department format, from the department's examples). Returns (paragraphs, name)."""
    st.caption("EAU: standard LND (obturator, internal and external iliac); no bowel preparation; ERAS. Consider a "
               "frozen section of the ureteral margins with CIS (Campbell-Walsh-Wein). Neobladder is not described "
               "here – type it in the extra field if done.")
    if woman:
        kind = c1.selectbox("Apjoms", ["ar dzemdi, piedēkļiem un maksts priekšējo sienu",
                                       "ar maksts priekšējo sienu (dzemde iepriekš izņemta)",
                                       "orgānus saudzējoša"], key="op-rc-kind-w")
        urethra = c2.checkbox("Uretrektomija", True, key="op-rc-ure-w")
        nerve = False
    else:
        kind = "cistoprostatektomija"
        nerve = c1.checkbox("Nervus saudzējoša", key="op-rc-nerve")
        urethra = c2.checkbox("Uretrektomija", key="op-rc-ure-m")
    lnd = c3.selectbox("Limfadenektomija", ["standarta abpusēja", "paplašināta abpusēja", "nav"], key="op-rc-lnd")
    diversion = c1.selectbox("Urīna novadīšana", ["ileum conduit", "ureterokutanostomija", "nav (nefrostomas)"],
                             key="op-rc-div")
    frozen = c2.checkbox("Urīnvadu malu steidzamā histoloģija", key="op-rc-frozen")
    anesthesia = c3.selectbox("Anestēzija", ["Epidurālā un intubācijas anestēzijā", "Intubācijas anestēzijā"],
                              key="op-rc-an")
    finding = st.text_input("Atrade (optional)", key="op-rc-finding",
                            placeholder="e.g. Vizualizē plašu audzēju ar ieaugšanu a. iliaca interna sinistra.")
    k1, k2, k3, k4 = st.columns(4)
    if diversion == "ileum conduit":
        dist = k1.number_input("No ileocekālā leņķa (cm)", 5, 60, 25, key="op-rc-dist")
        length = k2.number_input("Segmenta garums (cm)", 8, 30, 15, key="op-rc-len")
        technique = k3.selectbox("Ureteroileālā anastomoze", ["atsevišķi (Bricker)", "kopīgā platformā (Wallace)"],
                                 key="op-rc-tech")
    elif diversion == "ureterokutanostomija":
        ucs_side = k1.selectbox("Ureterokutanostoma", ["abpusēja", "vienpusēja (transureteroureterostomija)"],
                                key="op-rc-ucs")
    skin = k4.selectbox("Āda", ["intrakutānu šuvi", "atsevišķām šuvēm", "skavām"], key="op-rc-skin")
    drains = c1.number_input("Drenas", 0, 4, 1, key="op-rc-drains")
    ward = c2.text_input("Pārved uz", "nodaļu", key="op-rc-ward")

    text = [f"{anesthesia}, {pt_dat} muguras guļā veic vidējo apakšējo laparotomijas griezienu, atverot "
            "preperitoneālo telpu. Mobilizē peritoneju kraniāli." + (f" {finding.strip().rstrip('.')}." if finding.strip()
                                                                      else "")]
    if lnd != "nav":
        text.append(("Veic abpusēju standarta iegurņa limfadenektomiju – izdala obturatoros, iekšējos un ārējos "
                     "iliakālos limfmezglus" if lnd == "standarta abpusēja" else
                     "Veic abpusēju paplašinātu iegurņa limfadenektomiju līdz kopējo iliakālo artēriju līmenim") +
                    "; preparātus nosūta histoloģiskai izmeklēšanai atsevišķi pa pusēm.")
    ureters = ("Izdala abus urīnvadus, kurus pārdala distālajā trešdaļā" +
               (", urīnvadu distālās malas nosūta steidzamai histoloģiskai izmeklēšanai" if frozen else "") + ".")
    if not woman:
        text.append(f"{ureters} Cauršuj un liģē urīnpūšļa asinsvadu kājiņas bilaterāli. " +
                    ("Saudzējot neirovaskulāros kūlīšus, i" if nerve else "I") +
                    "zdala prostatas sānu virsmas, pārdala puboprostatiskās saites, nošuj un pārdala dorsālo vēnu "
                    "kompleksu. Pārdala uretru prostatas apeksa līmenī. Vienā blokā ekstirpē urīnpūsli ar prostatu un "
                    "sēklas pūslīšiem" + (", uretru ekstirpē atsevišķi no perineālas pieejas" if urethra else "") +
                    ". Rūpīga hemostāze.")
    elif kind.startswith("ar dzemdi"):
        text.append(f"Mobilizē urīnpūsli. {ureters} Cauršuj un liģē urīnpūšļa asinsvadu kājiņas bilaterāli. Iešķeļot "
                    "maksti gar laterālajām sienām un mobilizējot urīnpūsli distāli, liģējot audus kraniāli izdala "
                    "dzemdes kaklu. Atverot peritoneju, vienā blokā ekstirpē urīnpūsli" +
                    (", uretru" if urethra else "") + ", maksts priekšējo sienu, dzemdi ar piedēkļiem. Nošuj maksts "
                    "defektu. Rūpīga hemostāze.")
    elif kind.startswith("ar maksts"):
        text.append(f"Mobilizē urīnpūsli. {ureters} Cauršuj un liģē urīnpūšļa asinsvadu kājiņas bilaterāli. Iešķeļot "
                    "maksti gar laterālajām sienām un mobilizējot urīnpūsli distāli, vienā blokā ekstirpē urīnpūsli" +
                    (", uretru" if urethra else "") + " un maksts priekšējo sienu. Nošuj maksts defektu. Rūpīga "
                    "hemostāze.")
    else:
        text.append(f"Mobilizē urīnpūsli. {ureters} Cauršuj un liģē urīnpūšļa asinsvadu kājiņas bilaterāli. Saudzējot "
                    "dzemdi un maksti, izdala urīnpūsli no maksts priekšējās sienas un ekstirpē urīnpūsli" +
                    (" ar uretru" if urethra else "") + ". Rūpīga hemostāze.")
    stoma = True
    if diversion == "ileum conduit":
        text.append(f"Pārdala peritoneju. {dist} cm no ileocekālā leņķa {length} cm garā segmentā izdala ileum. "
                    "Iešķeļot peritoneju un pārdalot mezentēriju ar ENDO GIA palīdzību, izdala zarnas segmentu. Ar "
                    "mehāniskiem šuvējiem veido zarnu anastomozi gals galā. Kreiso urīnvadu izvada retroperitoneāli uz "
                    "labo pusi. Urīnvadus implantē ileum segmentā ne-antirefluksa tehnikā "
                    f"{'atsevišķi (Bricker)' if technique.startswith('atsevišķi') else 'kopīgā platformā (Wallace)'}, "
                    "kanulējot urīnvadus ar UK. UK izvada caur atvērto ileum galu. Rūpīga hemostāze. Labajā pusē parastā "
                    "kārtā veido ileostomu (urostomu), UK izvadot uz āru.")
    elif diversion == "ureterokutanostomija":
        text.append("Urīnvadus izvada caur vēdera priekšējo sienu un veido " +
                    ("abpusēju ureterokutanostomu" if ucs_side == "abpusēja" else
                     "transureteroureterostomiju un vienpusēju ureterokutanostomu") +
                    ", urīnvadus kanulē ar UK. Rūpīga hemostāze.")
    else:
        stoma = False
        text.append("Urīnvadu distālos galus nosien (urīna novadīšana pa nefrostomām).")
    text.append((f"Vēdera dobumā {drains} drena{'s' if drains > 1 else ''}, b" if drains else "B") +
                f"rūci sašuj pa kārtām. Āda tiek sašūta ar {skin}." +
                (" Virs stomas piestiprina urīna uztvērējmaisiņu." if stoma else "") +
                f" {pt_acc} pārved uz {ward.strip()}.")
    if woman:
        main = {"ar dzemdi, piedēkļiem un maksts priekšējo sienu": "Cistektomija, histerektomija ar piedēkļiem, maksts "
                                                                   "priekšējās sienas rezekcija",
                "ar maksts priekšējo sienu (dzemde iepriekš izņemta)": "Cistektomija, maksts priekšējās sienas rezekcija",
                "orgānus saudzējoša": "Orgānus saudzējoša cistektomija"}[kind]
    else:
        main = ("Nervus saudzējoša cistoprostatektomija" if nerve else "Cistoprostatektomija")
    name = main + (", uretrektomija" if urethra else "") + \
        {"ileum conduit": ", ileum conduit izveide", "ureterokutanostomija": ", ureterokutanostomija",
         "nav (nefrostomas)": ""}[diversion] + (", iegurņa limfadenektomija" if lnd != "nav" else "")
    return text, name


def cystoscopy_operation(c1, c2, c3, pt_acc):
    """Diagnostic cystoscopy, with or without biopsy. Returns (paragraphs, name)."""
    st.caption("EAU NMIBC: flexible cystoscopy where available (strong); describe size, site and appearance of every "
               "lesion (a bladder diagram helps); cold-cup biopsy of suspicious areas.")
    pt_nom = "Paciente" if pt_acc == "Pacienti" else "Pacients"
    anesthesia = c1.selectbox("Anestēzija", ["Vietējā anestēzijā (lidokaīna gels)", "Spinālā anestēzijā",
                                             "Intravenozā anestēzijā"], key="op-cy-an")
    scope = c2.selectbox("Instruments", ["fleksiblu cistoskopu", "rigīdu cistoskopu"], key="op-cy-scope")
    urethra = c3.text_input("Uretra", "brīvi caurejama", key="op-cy-ure")
    mucosa = c1.selectbox("Gļotāda", ["bez patoloģiskām izmaiņām", "hiperemēta", "trabekulāra",
                                      "ar pēcoperācijas rētu izmaiņām"], key="op-cy-muc")
    orifices = c2.text_input("Urīnvadu atveres", "tipiskā vietā, izdalās dzidrs urīns", key="op-cy-ori")
    biopsy = c3.checkbox("Biopsija", key="op-cy-bx")
    finding = st.text_input("Atrade (optional)", key="op-cy-find",
                            placeholder="e.g. Kreisajā sānu sienā ~1 cm papillārs jaunveidojums.")
    where = st.text_input("Biopsijas vieta", "aizdomīgās gļotādas vietas", key="op-cy-where") if biopsy else ""
    wash = st.checkbox("Urīnpūšļa skalojums citoloģijai", key="op-cy-wash")
    text = [f"{pt_nom} atrodas dorsālās litotomijas guļā. {anesthesia}, pēc operācijas lauka apdares urīnpūslī ievada "
            f"{scope}. Uretra {urethra.strip()}. Urīnpūšļa gļotāda {mucosa}." +
            (f" {finding.strip().rstrip('.')}." if finding.strip() else " Jaunveidojumus nekonstatē.") +
            f" Urīnvadu atveres {orifices.strip()}." +
            (f" Ar biopsijas knaiblēm paņem biopsiju no {where.strip()}, biopsijas vietu koagulē." if biopsy else "") +
            (" Paņem urīnpūšļa skalojumu citoloģiskai izmeklēšanai." if wash else "") +
            " Evakuē instrumentu. Manipulācija noritēja bez komplikācijām."]
    return text, "Cistoskopija" + (", biopsija" if biopsy else "")


def scrotal_infection_operation(op, c1, c2, c3, pt_acc):
    """Orchiectomy, inguinal orchifuniculectomy, epididymectomy, revision / drainage (department wording)."""
    side = c1.radio("Puse", ["Kreisā", "Labā"], horizontal=True, key="op-so-side")
    gen, acc, short, adj = {"Kreisā": ("kreisā", "kreiso", "kr.", "Kreisās"),
                            "Labā": ("labā", "labo", "lb.", "Labās")}[side]
    anesthesia = c2.selectbox("Anestēzija", ["spinālā anestēzijā", "vispārējā anestēzijā"], key="op-so-an")
    pus = c3.checkbox("Strutas – uzsējums", key="op-so-pus")
    pus_text = " Izdalās strutas, paņem materiālu uzsējumam." if pus else ""
    if op == "Orchiectomy (scrotal)":
        length = c1.number_input("Griezums (cm)", 2, 12, 6, key="op-so-len")
        drain = c2.checkbox("Drena ar aktīvo aspirāciju", True, key="op-so-drain")
        text = [f"{anesthesia.capitalize()}, pēc op. lauka apdares izdara griezienu scrotum {gen} sēklinieka rajonā "
                f"{length} cm.{pus_text} Izdala funiculus spermaticus, to pārdala un liģē. Izdala atsevišķi ductus "
                f"deferens un a. et v. testiculares, tās cauršuj. Ekstirpē {acc} testis. Rūpīga hemostāze." +
                (" Caur kontrapertūru atstāj drenu ar aktīvo aspirāciju." if drain else "") +
                f" Brūci sašuj pa kārtām. Sterili pārsēji. {pt_acc} vispārēji apmierinošā stāvoklī pārved uz nodaļu.",
                "Preparātu nosūta histoloģiskai izmeklēšanai."]
        name = f"{adj} puses orhektomija" + (", drenāža" if drain else "")
    elif op == "Inguinal orchifuniculectomy":
        text = [f"{anesthesia.capitalize()}, pēc op. lauka apdares izdara griezienu {short} cirkšņa rajonā. Šķeļ "
                f"ingvinālo kanālu, eksplorē {short} sēklinieku pie kanāla ārējā gredzena, funiculus spermaticus, "
                "pārdala atsevišķi ductus deferens un a. et v. testiculares iekšējās atveres līmenī." + pus_text +
                f" Veic {short} sēklinieka izdalīšanu un ekstirpāciju. Rūpīga hemostāze. Slēdz ingvinālo kanālu ar "
                "atsevišķām šuvēm. Brūci slēdz pa kārtām. Sterils pārsējs. "
                f"{pt_acc} vispārēji apmierinošā stāvoklī pārved uz nodaļu.",
                "Preparātu nosūta histoloģiskai izmeklēšanai."]
        name = f"{adj} puses orhofunikulektomija"
    elif op == "Epididymectomy":
        look = st.text_input("Piedēklis", "prāvāku, cietāku, iekaisīgi izmainītu", key="op-so-look")
        text = [f"{anesthesia.capitalize()}, pēc op. lauka trīsreizējas apdares veic griezienu skrotum ādā {short} "
                f"pusē, šķeļ tunica vaginalis testis, izdala {short} sēklinieku. Vizualizē {look.strip()} sēklinieka "
                f"piedēkli.{pus_text} Izdala sēklinieka piedēkli, pārdala un atsevišķi nošuj {short} ductus deferens, "
                "parastā kārtā veic epididimektomiju. Rūpīga hemostāze. Brūci slēdz pa kārtām. Sterils pārsējs. "
                f"{pt_acc} vispārēji apmierinošā stāvoklī pārved uz nodaļu.",
                "Preparātu nosūta histoloģiskai izmeklēšanai."]
        name = f"Epididimektomija {short} pusē"
    else:  # revision / drainage
        where = c1.selectbox("Piekļuve", ["sēklinieku maisiņa", "cirkšņa"], key="op-so-where")
        found = st.text_input("Atrade", "sastrutojušu hematomu", key="op-so-found",
                              placeholder="e.g. abscesu, sastrutojušu hematomu, piocēli")
        text = [f"{anesthesia.capitalize()}, pēc op. lauka apdares izdara griezienu {short} {where} rajonā. Eksplorē "
                f"{'funiculus spermaticus' if where == 'cirkšņa' else 'sēklinieku maisiņa saturu, sēklinieks vitāls'}, "
                f"evakuē {found.strip()}." + (" Paņem materiālu uzsējumam." if pus else "") + " Dobumu skalo ar "
                "antiseptisku šķīdumu. Atstāj drenu. Brūci slēdz pa kārtām. Sterils pārsējs. "
                f"{pt_acc} vispārēji apmierinošā stāvoklī pārved uz nodaļu."]
        name = f"{adj} puses {'funikula' if where == 'cirkšņa' else 'sēklinieku maisiņa'} revīzija, drenāža"
    return text, name


def prostatic_abscess_operation(c1, c2, c3, pt_acc):
    """Prostatic abscess drainage: TRUS-guided aspiration or transurethral deroofing. Returns (paragraphs, name)."""
    st.caption("Systematic review 2020: TRUS-guided aspiration for simple abscesses in stable patients; transurethral "
               "deroofing for unwell patients, multiloculated abscesses, outlet obstruction or failed aspiration. "
               "Send pus for culture.")
    method = c1.selectbox("Metode", ["TRUS vadīta punkcija", "Transuretrāla atvēršana (deroofing)"], key="op-pa-m")
    lobe = c2.selectbox("Lokalizācija", ["labajā daivā", "kreisajā daivā", "abās daivās", "vidējā daivā"],
                        key="op-pa-lobe")
    size = c3.number_input("Izmērs (cm)", 0.5, 10.0, 2.5, step=0.5, key="op-pa-size")
    ml = c1.number_input("Strutas (ml)", 0, 200, 10, key="op-pa-ml")
    if method.startswith("TRUS"):
        anesthesia = c2.selectbox("Anestēzija", ["Vietējā anestēzijā", "Intravenozā anestēzijā", "Spinālā anestēzijā"],
                                  key="op-pa-an")
        route = c3.selectbox("Pieeja", ["transrektāli", "transperineāli"], key="op-pa-route")
        drain = st.checkbox("Atstāj drenu (pigtail)", key="op-pa-drain")
        text = [f"{anesthesia}, pacients sānu guļā. Pēc op. lauka apdares TRUS kontrolē prostatas {lobe} vizualizē "
                f"~{_comma(size)} cm hipoehogēnu šķidruma kolekciju. {route.capitalize()} ar punkcijas adatu veic "
                f"abscesa punkciju, aspirē ~{ml} ml strutaina satura, materiālu nosūta uzsējumam. Dobumu skalo ar "
                "fizioloģisko šķīdumu." + (" Abscesa dobumā atstāj pigtail drenu." if drain else "") +
                f" Manipulācija noritēja bez komplikācijām. {pt_acc} vispārēji apmierinošā stāvoklī pārved uz nodaļu."]
        name = "Prostatas abscesa punkcija un drenāža TRUS kontrolē"
    else:
        anesthesia = c2.selectbox("Anestēzija", ["Spinālā anestēzijā", "Intubācijas narkozē"], key="op-pa-an2")
        fr = c3.selectbox("3-kanālu katetrs (Fr)", ["18", "20", "22"], index=1, key="op-pa-fr")
        text = [f"{anesthesia}, pacients dorsālās litotomijas guļā. Pēc op. lauka apdares uretrā ievada rezektoskopu. "
                f"Prostatas {lobe} vizualizē izspīlējumu. Ar cilpu atver abscesa dobumu (deroofing), izdalās ~{ml} ml "
                "strutu, materiālu nosūta uzsējumam. Abscesa dobuma sienas rezecē, dobumu atskalo. Rūpīga hemostāze. "
                f"Urīnpūslī ievada {fr} Fr 3-kanālu Foley katetru. {pt_acc} vispārēji apmierinošā stāvoklī pārved uz "
                "nodaļu."]
        name = "Transuretrāla prostatas abscesa atvēršana un drenāža"
    return text, name


def stricture_operation(op, c1, c2, c3, pt_dat, pt_acc):
    """Operation descriptions for urethral strictures. Returns (paragraphs, name)."""
    notes = {
        "Optical urethrotomy (DVIU)": "EAU: DVIU/dilatation for a primary, single, short (< 2 cm) bulbar stricture; not "
                                      "for penile strictures (strong); not more than twice if urethroplasty is possible "
                                      "(strong); laser and cold knife are equivalent. Catheter ≤ 72 h after an "
                                      "uncomplicated DVIU (weak).",
        "Urethral dilation": "EAU: visually controlled dilatation rather than blind dilatation (weak); same indications "
                             "as DVIU.",
        "Meatotomy": "EAU: meatoplasty or distal urethroplasty for meatal stenosis (weak); send tissue if lichen "
                     "sclerosus is suspected.",
        "Urethroplasty (EPA / buccal graft)": "EAU: transecting EPA for short post-traumatic bulbar strictures (strong); "
                                              "free graft for bulbar strictures not suitable for EPA (strong); dorsal or "
                                              "ventral graft as per expertise; urethrogram before catheter removal "
                                              "(strong). Check this template against your surgeons' practice.",
    }
    st.caption(notes[op])
    anesthesia = c1.selectbox("Anestēzija", ["spinālā anestēzijā", "vispārējā anestēzijā", "vietējā anestēzijā"],
                              key=f"op-st-an-{op}")
    an = anesthesia[0].upper() + anesthesia[1:]
    fr = c2.selectbox("Katetrs (Ch)", ["14", "16", "18", "20"], index=2 if op != "Urethroplasty (EPA / buccal graft)"
                      else 1, key=f"op-st-fr-{op}")
    if op != "Meatotomy":
        loc_key = c3.selectbox("Striktūras vieta", [k for k in STRICTURE_LOC if k != "Meatus"],
                               index=1 if op == "Urethroplasty (EPA / buccal graft)" else 0, key=f"op-st-loc-{op}")
        loc = STRICTURE_LOC[loc_key]
    if op == "Optical urethrotomy (DVIU)":
        tool = c1.selectbox("Instruments", ["Sachse nazis (aukstais nazis)", "lāzers (Ho:YAG)", "lāzers (tūlija)"],
                            key="op-st-tool")
        clock = c2.selectbox("Incīzija", ["pl. 12", "pl. 11 un 1", "pl. 12 un 6"], key="op-st-clock")
        wire = c3.checkbox("Drošības stīga", True, key="op-st-wire")
        bladder = st.text_input("Urīnpūšļa atrade (optional)", key="op-st-bladder",
                                placeholder="e.g. urīnpūslī konkrements Ø 1 cm – evakuē")
        laser = tool.startswith("lāzers")
        name = "Lāzeruretrotomija" if laser else "Optiskā uretrotomija (pēc Sachse)"
        cut = (f"Ievada lāzera šķiedru ({tool.split('(')[1].rstrip(')')}) un veic striktūras pārdalīšanu {clock}"
               if laser else f"Ar Sachse nazi veic striktūras pārdalīšanu {clock}")
        text = [f"Pēc operācijas lauka tualetes {anesthesia}, {pt_dat} atrodoties dorsālās litotomijas guļā, uretrā "
                f"ievada uretrotomu. Uretras distālā daļa neizmainīta. Uretras {loc} konstatē striktūru."
                + (" Caur striktūru urīnpūslī ievada drošības stīgu." if wire else "")
                + f" {cut} līdz veseliem audiem, rētaudus pārdalot pilnā biezumā. Instrumentu brīvi ievada "
                "urīnpūslī; proksimāli citas striktūras nekonstatē."
                + (f" {bladder.strip().rstrip('.')[0].upper() + bladder.strip().rstrip('.')[1:]}." if bladder.strip()
                   else " Urīnpūšļa gļotāda bez patoloģiskas atrades.")
                + " Evakuē instrumentu" + (" un drošības stīgu" if wire else "") + ".",
                f"Operāciju nobeidzot urīnpūslī ievada Ch{fr} Foley tipa urīnpūšļa ilgkatetru.",
                f"Vispārējā apmierinošā stāvoklī ar funkcionējošu urīnpūšļa ilgkatetru {pt_acc.lower()} pārved uz "
                "nodaļu."]
    elif op == "Urethral dilation":
        method = c1.selectbox("Dilatācija", ["ar metāla bužiem", "ar fasciālajiem dilatatoriem pa stīgu",
                                             "ar balona dilatatoru"], key="op-st-method")
        upto = c2.selectbox("Līdz (Ch)", ["18", "20", "22", "24", "26"], index=2, key="op-st-upto")
        name = "Uretras dilatācija"
        text = [f"Pēc operācijas lauka tualetes {anesthesia}, {pt_dat} atrodoties dorsālās litotomijas guļā, uretrā "
                f"ievada cistoskopu. Uretras {loc} konstatē striktūru. Caur striktūru urīnpūslī ievada vadītājstīgu. "
                f"Vizuālā kontrolē veic pakāpenisku striktūras dilatāciju {method} līdz Ch{upto}. Kontroles "
                "uretroskopijā lūmenis brīvs, dziļus gļotādas plīsumus un viltus ceļu nekonstatē. Evakuē instrumentu.",
                f"Urīnpūslī ievada Ch{fr} Foley tipa urīnpūšļa ilgkatetru.",
                f"Vispārējā apmierinošā stāvoklī {pt_acc.lower()} pārved uz nodaļu."]
    elif op == "Meatotomy":
        length = c3.number_input("Incīzija (mm)", 2, 20, 8, key="op-st-mm")
        calibre = c1.selectbox("Kalibrē līdz (Ch)", ["18", "20", "22", "24"], index=2, key="op-st-cal")
        catheter = c2.checkbox("Ilgkatetrs", True, key="op-st-mcath")
        biopsy = c3.checkbox("Audi histoloģijai", False, key="op-st-mbiopsy")
        name = "Meatotomija"
        text = [f"Pēc operācijas lauka apdares {anesthesia} vizualizē sašaurinātu uretras ārējo atveri (meatus). Veic "
                f"meatus ventrālo incīziju {length} mm garumā līdz normāla platuma uretras lūmenim. Uretras gļotādas "
                "malas ar atsevišķām uzsūcošām šuvēm (Vicryl 5-0) piešuj glans ādai. Uretru brīvi kalibrē līdz "
                f"Ch{calibre}. Rūpīga hemostāze."
                + (f" Urīnpūslī ievada Ch{fr} Foley tipa ilgkatetru." if catheter else "") + " Sterils pārsējs.",
                f"Vispārējā apmierinošā stāvoklī {pt_acc.lower()} pārved uz nodaļu."]
        if biopsy:
            text.append("Izgrieztos audus nosūta histoloģiskai izmeklēšanai.")
    else:  # urethroplasty
        kind = c1.selectbox("Tehnika", ["Ekscīzija un primāra anastomoze (EPA)",
                                        "Vaiga gļotādas transplantāts – dorsāls onlay",
                                        "Vaiga gļotādas transplantāts – ventrāls onlay"], key="op-st-kind")
        length = c2.number_input("Striktūras garums (cm)", 0.5, 20.0, 2.0, step=0.5, key="op-st-len")
        drain = c3.checkbox("Drena", True, key="op-st-drain")
        cheek = c1.radio("Vaigs", ["Kreisais", "Labais"], horizontal=True, key="op-st-cheek") if "Vaiga" in kind else ""
        donor = c2.selectbox("Donorvieta", ["šuj ar Vicryl 4-0", "atstāj atvērtu"], key="op-st-donor")             if "Vaiga" in kind else ""
        len_text = f"{length:g}".replace(".", ",")
        opening = (f"{an}, {pt_dat} atrodoties litotomijas guļā, pēc operācijas lauka apdares veic perineālu griezienu "
                   "pa viduslīniju. Pārdala m. bulbospongiosus, izdala un mobilizē bulbāro uretru. Uretrā ievadītā "
                   f"katetra līmenī lokalizē striktūru (garums ~{len_text} cm).")
        if kind.startswith("Ekscīzija"):
            name = "Uretroplastika ar striktūras ekscīziju un primāru anastomozi (EPA)"
            body = ("Striktūras segmentu pilnībā ekscidē līdz veseliem, labi apasiņotiem uretras galiem. Abus galus "
                    "spatulē un veido beztensijas anastomozi ar atsevišķām uzsūcošām šuvēm (PDS 5-0). Pirms "
                    f"anastomozes priekšējās sienas slēgšanas urīnpūslī ievada Ch{fr} silikona katetru.")
        else:
            dorsal = "dorsāls" in kind
            name = f"Uretroplastika ar vaiga gļotādas transplantātu ({'dorsāls' if dorsal else 'ventrāls'} onlay)"
            body = (f"No {'kreisā' if cheek == 'Kreisais' else 'labā'} vaiga iekšējās virsmas, saudzējot Stenona vada "
                    f"atveri, izgriež gļotādas transplantātu atbilstoši striktūras garumam; donorvietu {donor}. "
                    "Transplantātu atbrīvo no zemgļotādas audiem. Uretru atver striktūras garumā "
                    + ("pa dorsālo sienu (pēc Barbagli)" if dorsal else "pa ventrālo sienu") + " līdz veseliem audiem. "
                    "Transplantātu novieto " + ("dorsāli uz corpora cavernosa" if dorsal else "ventrāli") + " un fiksē "
                    "ar nepārtrauktu uzsūcošu šuvi (PDS 5-0), papildinot uretras lūmenu. Pirms uretras slēgšanas "
                    f"urīnpūslī ievada Ch{fr} silikona katetru.")
        text = [opening + " " + body,
                "Rūpīga hemostāze." + (" Brūcē atstāj drenu." if drain else "")
                + " M. bulbospongiosus un brūci slēdz pa kārtām ar uzsūcošām šuvēm. Sterils spiedošs pārsējs.",
                f"Vispārējā apmierinošā stāvoklī {pt_acc.lower()} pārved uz nodaļu.",
                "Izgrieztos audus nosūta histoloģiskai izmeklēšanai." if kind.startswith("Ekscīzija") else ""]
        text = [t for t in text if t]
    return text, name


def urs_operation(flexible, c1, c2, c3, pt_dat, pt_acc):
    """Operation description of URS or FURS, from the steps that were done. Returns (paragraphs, name)."""
    st.caption("EAU urolithiasis: use a safety wire, take stones out only under direct vision, and a stent is not "
               "needed after an uncomplicated, complete URS (place one for residual fragments, ureteral injury, "
               "oedema, infection or when in doubt). Tick only what was done.")
    side = c1.radio("Puse", ["Kreisā", "Labā"], horizontal=True, key="op-u-side")
    left = side == "Kreisā"
    ost, ost_loc, adj, m_acc = (("kreiso", "kreisajā", "kreisās", "kreiso") if left
                                else ("labo", "labajā", "labās", "labo"))
    tag = "f" if flexible else "s"
    anesthesia = c2.selectbox("Anestēzija", ["vispārējā anestēzijā", "spinālā anestēzijā"],
                              index=0 if flexible else 1, key=f"op-u{tag}-an")
    stent_out = c3.checkbox("Iepriekšējā stenta evakuācija sākumā", flexible, key=f"op-u{tag}-stentout")
    cysto = st.text_input("Cistoskopijas atrade", "Urīnpūšļa gļotāda neizmainīta, ostijas simetriskas.",
                          key=f"op-u{tag}-cysto")
    a1, a2, a3 = st.columns(3)
    wire = a1.checkbox("Drošības stīga", True, key=f"op-u{tag}-wire")
    pyelo = a2.checkbox("Retrogrāda pielogrāfija", False, key=f"op-u{tag}-pyelo")
    analysis = a3.checkbox("Fragmenti ķīmiskai analīzei", True, key=f"op-u{tag}-analysis")

    text = [f"{anesthesia[0].upper() + anesthesia[1:]}, {pt_dat} atrodoties dorsālās litotomijas guļā, pēc operācijas "
            f"lauka apdares urīnpūslī ievada cistoskopu. Instrumenta ievadīšana brīva. {cysto.strip().rstrip('.')}."]
    name_parts, extracted = [], False
    if stent_out:
        text.append(f"Vizualizē {ost} ostiju un DJ stenta distālo galu. Ar standziņām satver stenta galu un "
                    "evakuē stentu.")
        name_parts.append(f"{adj} puses DJ stenta evakuācija")
    access = []
    if wire:
        access.append(f"Rentgenoskopijas kontrolē caur {ost} ostiju urīnvadā līdz nieres bļodiņai ievada hidrofilo "
                      "drošības stīgu.")
    if pyelo:
        access.append("Veic retrogrādu pielogrāfiju: kontrastviela aizpilda nieres bļodiņu un kausiņus, "
                      "ekstravazāciju nekonstatē.")
    if access:
        text.append(" ".join(access))
    along = "Gar drošības stīgu" if wire else "Vizuālā kontrolē"

    if not flexible:
        st.markdown("**Ureterorenoskopija**")
        b1, b2, b3 = st.columns(3)
        reach = b1.selectbox("Cik tālu veikta URS", URS_REACH, key="op-us-reach")
        finding = b2.selectbox("Atrade", URS_FINDINGS, key="op-us-find") if reach in URS_REACH[:2] else None
        where = b3.selectbox("Vieta urīnvadā", list(URS_LOCATIONS), key="op-us-where") \
            if finding in ("Stone", "Stricture") else None
        actions, size, impacted = [], None, False
        if finding == "Stone":
            d1, d2 = st.columns([1, 2])
            size = d1.number_input("Konkrements (mm)", 0, 50, value=None, key="op-us-size")
            impacted = d1.checkbox("Iesprūdis", key="op-us-impacted")
            actions = d2.pills("Ko darīja ar konkrementu", URS_ACTIONS, selection_mode="multi", key="op-us-actions") or []
            result = d2.selectbox("Rezultāts", list(URS_RESULTS), key="op-us-result") if actions else None
        side_word = "" if stent_out else f"{adj} puses "
        if reach == "No URS – stent only":
            name_parts.append(f"cistoskopija, {side_word}retrogrāda DJ stenta ievietošana")
        else:
            name_parts.append(f"{side_word}ureterorenoskopija (URS)")
        if reach == "Not possible (narrow ureter)":
            text.append(f"{along} {ost_loc} ostijā ievada semirigīdo ureterorenoskopu. Šaura urīnvada dēļ instrumentu "
                        "tālāk virzīt nav iespējams. Evakuē ureterorenoskopu.")
        elif reach in URS_REACH[:2]:
            urs = [f"{along} {ost_loc} ostijā ievada semirigīdo ureterorenoskopu un to virza pa urīnvadu."]
            if finding == "Stone":
                stone = ("iesprūdušu konkrementu" if impacted else "konkrementu") + (f" Ø {size} mm" if size else "")
                urs.append(f"Urīnvada {URS_LOCATIONS[where]} vizualizē {stone}.")
                broken = any(a in actions for a in ("Laser lithotripsy", "Pneumatic lithotripsy"))
                if "Laser lithotripsy" in actions:
                    urs.append("Caur darba kanālu ievada lāzera šķiedru un veic konkrementa kontakta lāzerlitotripsiju.")
                if "Pneumatic lithotripsy" in actions:
                    urs.append("Veic konkrementa pneimatisko litotripsiju.")
                if "Stone migrated to the kidney" in actions:
                    urs.append("Konkrements (tā fragments) migrē uz nieres bļodiņu.")
                what = "konkrementa fragmentu" if broken else "konkrementa"
                for action, tool in (("Basket extraction", "Dormia grozu"), ("Loop (cilpa) extraction", "cilpu")):
                    if action in actions:
                        urs.append(f"Ievada {tool} un vizuālā kontrolē veic {what} litoekstrakciju.")
                        extracted = True
                if actions:  # nothing done with the stone (e.g. URS only to place a stent): no result line
                    urs.append(URS_RESULTS[result])
                if "Laser lithotripsy" in actions:
                    name_parts.append("kontakta lāzerlitotripsija")
                if "Pneumatic lithotripsy" in actions:
                    name_parts.append("pneimatiskā litotripsija")
                if extracted:
                    name_parts.append("litoekstrakcija")
            elif finding == "No stone (passed)":
                urs.append("Pa urīnvada gaitu konkrementu nekonstatē (iespējams, izdalījies spontāni).")
            elif finding == "Stricture":
                urs.append(f"Urīnvada {URS_LOCATIONS[where]} konstatē sašaurinājumu (striktūru).")
            else:
                urs.append("Urīnvada gļotāda tūskaina, ar kontaktasiņošanu.")
            if reach == "Whole ureter, up to the renal pelvis":
                urs.append("Ureterorenoskopu virza līdz nieres bļodiņai; pa urīnvada gaitu "
                           + ("citus konkrementus nekonstatē." if finding == "Stone" else "patoloģiju nekonstatē."))
            urs.append("Evakuē ureterorenoskopu.")
            text.append(" ".join(urs))
        stent_default = True
    else:
        st.markdown("**Fleksibla ureterorenoskopija**")
        b1, b2, b3 = st.columns(3)
        semirigid = b1.checkbox("Sākumā semirigīda URS", True, key="op-uf-semi")
        sheath = b2.checkbox("Piekļuves apvalks (UAS)", True, key="op-uf-uas")
        sheath_size = b3.selectbox("UAS izmērs", ["9,5/11,5 Fr", "10/12 Fr", "11/13 Fr", "12/14 Fr"], index=1,
                                   key="op-uf-uas-size") if sheath else ""
        places = st.multiselect("Konkrementa vieta", list(FURS_PLACES), default=["Renal pelvis"], key="op-uf-places")
        d1, d2, d3 = st.columns(3)
        several = d1.checkbox("Vairāki konkrementi", key="op-uf-several")
        size = d1.number_input("Lielākais (mm)", 0, 60, value=None, key="op-uf-size")
        actions = d2.pills("Ko darīja ar konkrementu", FURS_ACTIONS, selection_mode="multi",
                           default=["Laser lithotripsy"], key="op-uf-actions") or []
        mode = d3.selectbox("Lāzera režīms", list(LASER_MODE), key="op-uf-mode") if "Laser lithotripsy" in actions else None
        result = d3.selectbox("Rezultāts", list(URS_RESULTS), key="op-uf-result")
        name_parts.append(("" if stent_out else f"{adj} puses ") + "fleksibla ureterorenoskopija (FURS)")
        if semirigid:
            text.append(f"{along} {ost_loc} ostijā ievada semirigīdo ureterorenoskopu un izmeklē {m_acc} urīnvadu – "
                        "konkrementus un sašaurinājumus nekonstatē. Ievada darba stīgu, evakuē ureterorenoskopu.")
        kidney = []
        if sheath:
            kidney.append(f"Pa darba stīgu rentgenoskopijas kontrolē ievada uretera piekļuves apvalku {sheath_size}.")
        kidney.append("Ievada fleksiblo ureterorenoskopu un sistemātiski apskata nieres bļodiņu un kausiņu sistēmu.")
        where = lv_list([FURS_PLACES[p] for p in places]) or "nieres bļodiņā"
        stone = ("konkrementus" if several else "konkrementu") + (
            (f" Ø līdz {size} mm" if several else f" Ø {size} mm") if size else "")
        kidney.append(f"{where[0].upper() + where[1:]} vizualizē {stone}.")
        if "Move lower-pole stone up (basket)" in actions:
            kidney.append(f"Ar Dormia grozu {'konkrementus' if several else 'konkrementu'} pārvieto uz augšējo "
                          "kausiņu grupu.")
        if "Laser lithotripsy" in actions:
            kidney.append("Caur darba kanālu ievada lāzera šķiedru un veic "
                          f"{'konkrementu' if several else 'konkrementa'} kontakta lāzerlitotripsiju {LASER_MODE[mode]}.")
            name_parts.append("kontakta lāzerlitotripsija")
        if "Basket extraction" in actions:
            what = ("konkrementu fragmentus" if several else "konkrementa fragmentus") if "Laser lithotripsy" in actions \
                else ("konkrementus" if several else "konkrementu")
            kidney.append(f"Ar Dormia grozu vizuālā kontrolē evakuē {what}.")
            name_parts.append("litoekstrakcija")
            extracted = True
        kidney.append(URS_RESULTS[result].replace("Kontrolē", "Kontroles apskatē kausiņu sistēmā"))
        kidney.append("Evakuē fleksiblo ureterorenoskopu" + (" un piekļuves apvalku, apskatot urīnvadu – bojājumus "
                                                             "nekonstatē." if sheath else "."))
        text.append(" ".join(kidney))
        stent_default = True

    st.markdown("**Beigās**")
    e1, e2, e3 = st.columns(3)
    stent_in = e1.checkbox("DJ stenta ievietošana", stent_default, key=f"op-u{tag}-stent")
    if stent_in:
        fr = e2.selectbox("Stents (Fr)", ["4,8", "6", "7"], index=1, key=f"op-u{tag}-stent-fr")
        cm = e3.selectbox("Garums (cm)", ["22", "24", "26", "28"], index=2, key=f"op-u{tag}-stent-cm")
        string = e1.checkbox("Ar diegu (izņemšanai)", False, key=f"op-u{tag}-string")
    foley = e2.checkbox("Foley katetrs", False, key=f"op-u{tag}-foley")
    if foley:
        foley_fr = e3.selectbox("Katetrs (Fr)", ["14", "16", "18", "20"], index=1, key=f"op-u{tag}-foley-fr")
    end = []
    if stent_in:
        end.append(f"{'Gar drošības stīgu' if wire else 'Pa stīgu'} rentgenoskopijas kontrolē ievada DJ stentu "
                   f"{fr} Fr / {cm} cm" + (" ar diegu" if string else "") + ". Stenta proksimālā cilpa nieres bļodiņā, "
                   "distālā – urīnpūslī. Evakuē stīgu.")
        name_parts.append("DJ stenta ievietošana" if not any("retrogrāda DJ" in p for p in name_parts) else "")
    elif wire:
        end.append("Evakuē drošības stīgu.")
    if foley:
        end.append(f"Urīnpūslī ievada Foley {foley_fr} Fr katetru.")
    if extracted and analysis:
        end.append("Konkrementa fragmentus nosūta ķīmiskā sastāva analīzei.")
    if end:
        text.append(" ".join(end))
    text.append(f"Vispārēji apmierinošā stāvoklī {pt_acc.lower()} pārved uz nodaļu.")
    name = ", ".join(p for p in name_parts if p)
    return text, name[0].upper() + name[1:]


def render_operation(sex):
    woman = sex == "Woman"
    pt_dat = "pacientei" if woman else "pacientam"
    pt_acc = "Pacienti" if woman else "Pacientu"
    st.subheader("Operācijas apraksts")
    op = st.selectbox("Operation", ["TURP", "TURU", "Re-TURU (scar resection)", "Radical cystectomy", "Cystoscopy (diagnostic) ± biopsy", "Circumcision", "Dorsumcision (dorsal slit)", "Cystoscopy, stent removal", "URS (semirigid)", "FURS (flexible)",
                                    "Cystolithotripsy", "Cystoscopy – clot evacuation, coagulation (hematuria)", "Optical urethrotomy (DVIU)", "Urethral dilation", "Meatotomy",
                                    "Urethroplasty (EPA / buccal graft)",
                                    "Epididymal tumour extirpation", "Prostatic abscess drainage", "Orchiectomy (scrotal)", "Inguinal orchifuniculectomy",
                                    "Epididymectomy", "Scrotal / funicular revision, drainage", "Hydrocele plasty", "Scrotal mass excision",
                                    "Laparoscopic radical prostatectomy", "Laparoscopic nephrectomy",
                                    "Laparoscopic kidney resection",
                                    "Laparoscopic pyeloplasty",
                                    "Percutaneous trocar epicystostomy (48)",
                                    "Nephrostomy change under US (48)"], key="op-type")
    text, name = [], {"TURP": "TURP", "TURU": "TURU", "Circumcision": "Cirkumcīzija"}.get(op, "")
    with st.expander("Details", expanded=True):
        c1, c2, c3 = st.columns(3)
        if op == "TURP":
            scope = c1.selectbox("Rezektoskops (Ch)", ["24", "26"], index=1, key="op-turp-scope")
            mucosa = st.text_input("Urīnpūšļa atrade", "gļotāda vizuāli neizmainīta, uretera ostijas brīvas", key="op-turp-mucosa")
            neck = c2.checkbox("Urīnpūšļa kakliņa incīzija", True, key="op-turp-neck")
            hemostasis = c3.selectbox("Hemostāze", ["ar vaporizācijas rullīti", "ar koagulācijas cilpu"], key="op-turp-hemo")
            fr = c1.selectbox("Katetrs (Fr)", ["18", "20", "22", "24"], index=1, key="op-turp-fr")
            balloon = c2.number_input("Balons (ml)", 5, 80, 40, key="op-turp-balloon")
            traction = c3.checkbox("Vieglā ekstenzijā", True, key="op-turp-traction")
            urine = c1.selectbox("Urīns pa katetru", ["viegli sārtas krāsas", "sārtas krāsas", "dzidrs"], key="op-turp-urine")
            ward = c2.text_input("Pārved uz", "4. nodaļu", key="op-turp-ward")
            histology = c3.checkbox("Čipsi uz histoloģiju", True, key="op-turp-hist")
            urine_text = "dzidrs urīns" if urine == "dzidrs" else f"{urine} urīns"
            text = [
                f"Pēc operācijas lauka tualetes, {pt_dat} atrodoties dorsālas litotomijas guļā urīnpūslī ievadīts "
                f"{scope} Ch rezektoskops; {mucosa.strip().rstrip('.')}.",
                "Veic prostatas hiperplazēto audu rezekciju no urīnpūšļa kakliņa līdz Colliculus seminalis. Līdzīgā "
                "kārtā veic hiperplazēto audu rezekciju lb. daivai un sekojoši kr. daivai līdz nosacītiem pl. 12:00. "
                "Rezekciju veic līdz prostatas ķirurģiskai kapsulai. Atskaloti rezekcijas čipsi."
                + (" Ar endoskopijas āķīti veikta urīnpūšļa kakliņa incīzija proksimāli no Colliculus seminalis." if neck else ""),
                f"Rūpīga hemostāze {hemostasis}, sākot no urīnpūšļa kakliņa un tālāk par prostatas ložu.",
                f"Operāciju nobeidzot urīnpūslī ievadīts 3-kanālu Foley tipa {fr} Fr katetrs. Balons uzpildīts līdz "
                f"{balloon} ml." + (" Vieglā ekstenzijā." if traction else ""),
                f"{pt_acc} vispārēji apmierinošā stāvoklī pārved uz {ward.strip()}. Pa katetru {urine_text}.",
            ]
            if histology:
                text.append("Rezekcijas čipsi nosūtīti histoloģiskai izmeklēšanai.")
        elif op == "TURU":
            scope = c1.selectbox("Rezektoskops (Ch)", ["24", "26"], index=1, key="op-turu-scope")
            number = c2.radio("Veidojumi", ["Viens", "Vairāki"], index=1, horizontal=True, key="op-turu-number")
            kind = c3.selectbox("Izskats", ["papillārs", "solīds", "plakans (CIS aizdomas)"], key="op-turu-kind")
            where = st.multiselect("Lokalizācija", list(TURU_LOCATION_LOC), default=["Right lateral wall"], key="op-turu-where")
            detrusor = c1.checkbox("Paraugā detruzora muskulis", True, key="op-turu-det")
            fr = c2.selectbox("Katetrs (Fr)", ["18", "20", "22"], index=1, key="op-turu-fr")
            balloon = c3.number_input("Balons (ml)", 5, 60, 20, key="op-turu-balloon")
            ward = c1.text_input("Pārved uz", "nodaļu", key="op-turu-ward")
            instill = c2.checkbox("Agrīna instilācija", True, key="op-turu-inst")
            drug = st.text_input("Instilācija", "Sol. Epirubicini 50 mg", key="op-turu-drug") if instill else ""
            single = number == "Viens"
            adjective = {"papillārs": ("papillāru", "papillārus"), "solīds": ("solīdu", "solīdus"),
                         "plakans (CIS aizdomas)": ("plakanu, sārtu", "plakanus, sārtus")}[kind][0 if single else 1]
            noun = "veidojumu" if single else "veidojumus"
            place = lv_list([TURU_LOCATION_LOC[w] for w in where]) or "—"
            text = [
                f"Pēc operācijas lauka tualetes, {pt_dat} atrodoties dorsālas litotomijas guļā urīnpūslī ievada "
                f"{scope} Ch rezektoskopu. Urīnpūšļa {place} vizualizē {adjective} {noun}.",
                ("Veikta veidojuma rezekcija" if single else "Veikta veidojumu rezekcija")
                + (", iegūstot arī detruzora muskuļa paraugu." if detrusor else ".") + " Rūpīga hemostāze.",
                f"{fr} Fr 3-kanālu ilgkatetrs, balons uzpildīts līdz {balloon} ml. {pt_acc} apmierinošā stāvoklī "
                f"pārved uz {ward.strip()}. Rezekcijas materiāls nosūtīts histoloģiskai izmeklēšanai.",
            ]
            if instill and drug.strip():
                text.append(f"Agrīnā pēcoperācijas periodā saņēma {drug.strip()} intravezikālu instilāciju.")
        elif op == "Re-TURU (scar resection)":
            st.caption("EAU NMIBC: a second resection 2–6 weeks after the first (T1, incomplete resection, no detrusor "
                       "muscle); it must include the primary tumour site. A single early instillation is meant for the "
                       "first TURU, not routinely for re-TURU.")
            scope = c1.selectbox("Rezektoskops (Ch)", ["24", "26"], index=1, key="op-rt-scope")
            where = st.multiselect("Rētas lokalizācija", list(TURU_LOCATION_LOC), default=["Right lateral wall"],
                                   key="op-rt-where")
            finding = c2.selectbox("Rētas atrade", ["bez redzama atlikušā audzēja", "ar atlikušiem papillāriem audiem",
                                                    "ar aizdomīgu, sārtu gļotādu"], key="op-rt-find")
            new = c3.text_input("Jauni veidojumi (empty = nav)", key="op-rt-new",
                                placeholder="e.g. jumtā papillārs veidojums 5 mm")
            separate = c1.checkbox("Rezekcijas pamatne atsevišķi", True, key="op-rt-sep")
            fr = c2.selectbox("Katetrs (Fr)", ["18", "20", "22"], index=1, key="op-rt-fr")
            balloon = c3.number_input("Balons (ml)", 5, 60, 20, key="op-rt-balloon")
            ward = c1.text_input("Pārved uz", "nodaļu", key="op-rt-ward")
            name = "Atkārtota TURU (re-TURU)"
            place = lv_list([TURU_LOCATION_LOC[w] for w in where]) or "—"
            mucosa = (f"Konstatēti jauni veidojumi: {new.strip().rstrip('.')}; tie rezecēti." if new.strip()
                      else "Pārējā urīnpūšļa gļotāda bez jauniem veidojumiem.")
            text = [
                f"Pēc operācijas lauka tualetes, {pt_dat} atrodoties dorsālas litotomijas guļā urīnpūslī ievada "
                f"{scope} Ch rezektoskopu. Urīnpūšļa gļotādu pārskata sistemātiski. Urīnpūšļa {place} vizualizē "
                f"iepriekšējās rezekcijas rētu {finding}. {mucosa}",
                "Veic rētaudu un primārā audzēja vietas rezekciju dziļumā līdz detruzora muskulim, iegūstot detruzora "
                "muskuļa paraugu." + (" Rezekcijas pamatnes paraugu nosūta atsevišķi." if separate else "")
                + " Rūpīga hemostāze.",
                f"{fr} Fr 3-kanālu ilgkatetrs, balons uzpildīts līdz {balloon} ml. {pt_acc} apmierinošā stāvoklī "
                f"pārved uz {ward.strip()}. Rezekcijas materiāls nosūtīts histoloģiskai izmeklēšanai.",
            ]
        elif op == "Cystoscopy, stent removal":
            side = c1.radio("Puse", ["Kreisā", "Labā", "Abas"], horizontal=True, key="op-cs-side")
            findings = st.text_input("Atrade", "uretra normāla platuma, urīnpūšļa gļotāda nav izmainīta, urēteru ostijas "
                                     "normālas, simetriskas", key="op-cs-find")
            removal = c2.checkbox("Stenta evakuācija", True, key="op-cs-remove")
            ostium = "kreiso" if side == "Kreisā" else "labo"
            if side == "Abas":
                name = "Cistoskopija, abpusēja pielovezikālo stentu evakuācija" if removal else "Cistoskopija"
            else:
                name = ("Cistoskopija, " + ("kr." if side == "Kreisā" else "lb.") + " puses pielovezikāla stenta "
                        "evakuācija" if removal else "Cistoskopija")
            text = [f"Pēc operācijas lauka tualetes, {pt_dat} atrodoties dorsālas guļā tiek ievadīts cistoskops, "
                    f"{findings.strip().rstrip('.')}."]
            if removal and side == "Abas":
                text.append("Vizualizē abas ostijas un pielovezikālo stentu galus. Ievada standziņas, stentus pēc kārtas "
                            "evakuē. Evakuācija veiksmīga.")
            elif removal:
                text.append(f"Vizualizē {ostium} ostiju un pielovezikālā stenta galu. Ievada standziņas, stentu evakuē. "
                            "Evakuācija veiksmīga.")
            text.append("Evakuē instrumentu.")
        elif op == "Percutaneous trocar epicystostomy (48)":
            incision = c1.number_input("Griezums virs simfīzes (cm)", 1, 5, 2, key="op-ec-cm")
            ch = c2.selectbox("Katetrs (Ch)", ["Not stated", "12", "14", "16", "18"], key="op-ec-ch")
            balloon = c3.number_input("Balons (ml NaCl 0,9%)", 2, 20, 5, key="op-ec-balloon")
            volume = c1.number_input("Izdalās urīns (ml)", 0, 5000, 400, step=50, key="op-ec-vol")
            look = c2.selectbox("Urīns", ["dzidra", "duļķaina", "ar asiņu piejaukumu"], key="op-ec-look")
            name = "Perkutāna troakāra epicistostomija"
            catheter = "epicistostomas katetru" if ch == "Not stated" else f"Ch {ch} epicistostomas katetru"
            text = [f"Pēc operācijas lauka apdares ar polispirtu šķīdumu lokālā anestēzijā ar Sol. Lidocaini 2%, veic "
                    f"ādas griezienu {incision} cm virs simfīzes. US kontrolē veic troakāra cistostomiju. Ievada "
                    f"{catheter}, iegūst urīnu, evakuē troakāru. Uzpilda katetra balonu ar {balloon} ml NaCl 0,9% "
                    "šķīduma. Fiksē katetru ar ādas šuvi. Pārsējs. "
                    f"Pa epicistostomu izdalās {volume} ml {look} urīna."]
        elif op == "Nephrostomy change under US (48)":
            side = c1.radio("Puse", ["Kreisā", "Labā", "Abas"], horizontal=True, key="op-nc-side")
            ch = c2.selectbox("Nefrostoma (Ch)", ["8", "10", "12", "14"], index=1, key="op-nc-ch")
            position = c3.selectbox("Novietojums", ["adekvāts", "neadekvāts – koriģēts"], key="op-nc-pos")
            if side == "Abas":
                name = "Abpusēja nefrostomu nomaiņa USS kontrolē"
                text = ["Pēc operācijas lauka apdares abās pusēs secīgi caur nefrostomu ievada stīgu. Uz stīgas evakuē "
                        f"nefrostomu. Uz stīgas ievada Ch {ch} nefrostomas katetru. Evakuē stīgu. USS kontrolē konstatē "
                        f"nefrostomu novietojumu. Nefrostomu novietojums {position}. Fiksē nefrostomu katetrus."]
            else:
                name = f"{'Kreisās' if side == 'Kreisā' else 'Labās'} nefrostomas nomaiņa USS kontrolē"
                text = ["Pēc operācijas lauka apdares caur nefrostomu ievada stīgu. Uz stīgas evakuē nefrostomu. Uz "
                        f"stīgas ievada Ch {ch} nefrostomas katetru. Evakuē stīgu. USS kontrolē konstatē nefrostomas "
                        f"novietojumu. Nefrostomas novietojums {position}. Fiksē nefrostomas katetru."]
        elif op == "Epididymal tumour extirpation":
            side = c1.radio("Puse", ["Labā", "Kreisā"], horizontal=True, key="op-ep-side")
            part = c2.selectbox("Daļa", ["Cauda epididimis", "Corpus epididimis", "Caput epididimis"], key="op-ep-part")
            size = c3.text_input("Izmērs", "~ 1 x 2 cm", key="op-ep-size")
            look = st.text_input("Veidojums", "blīvu, labi norobežotu veidojumu", key="op-ep-look")
            histology = st.text_input("Histoloģija (optional)", key="op-ep-hist")
            short, word = ("lb.", "lb.") if side == "Labā" else ("kr.", "kr.")
            name = f"Epididimis jaunveidojuma ekstirpācija {short} pusē"
            text = [f"Pēc op. lauka trīsreizējas apdares, veic griezienu skrotum ādā {short} pusē, šķeļ tunica vaginalis "
                    f"testis, izdala {word} sēklinieku. {part} rajonā konstatē {look.strip()} {size.strip()}, ko izdala un "
                    "rezecē. Rūpīga hemostāze. Brūci slēdz pa kārtām. Sterils pārsējs. "
                    f"{pt_acc} vispārēji apmierinošā stāvoklī pārved uz nodaļu."]
            if histology.strip():
                text.append(f"\nHistoloģija: {histology.strip()}")
        elif op == "Cystoscopy – clot evacuation, coagulation (hematuria)":
            text, name = hematuria_operation(c1, c2, c3, pt_acc)
        elif op == "Cystolithotripsy":
            text, name = cystolithotripsy_operation(c1, c2, c3, pt_acc)
        elif op in ("Optical urethrotomy (DVIU)", "Urethral dilation", "Meatotomy", "Urethroplasty (EPA / buccal graft)"):
            text, name = stricture_operation(op, c1, c2, c3, pt_dat, pt_acc)
        elif op in ("URS (semirigid)", "FURS (flexible)"):
            text, name = urs_operation(op == "FURS (flexible)", c1, c2, c3, pt_dat, pt_acc)
        elif op == "Laparoscopic radical prostatectomy":
            lnd = c1.checkbox("Iegurņa limfadenektomija", True, key="op-rp-lnd")
            nerve = c2.selectbox("Nervu saudzēšana", ["Nav", "Labajā pusē", "Kreisajā pusē", "Abās pusēs"],
                                 key="op-rp-nerve")
            fr = c3.selectbox("Katetrs (Fr)", ["16", "18", "20"], index=1, key="op-rp-fr")
            balloon = c1.number_input("Balons (ml)", 5, 40, 20, key="op-rp-balloon")
            drain = c2.checkbox("Drena pie anastomozes", True, key="op-rp-drain")
            extras = (["iegurņa limfadenektomiju"] if lnd else []) + \
                     ([f"nervu saudzēšanu ({nerve.lower()})"] if nerve != "Nav" else [])
            name = "Laparoskopiska radikāla prostatektomija" + (" ar " + lv_list(extras) if extras else "")
            text = [
                f"Pēc operācijas lauka 3x apdares ar polialkohola šķīdumu intubācijas narkozē, {pt_dat} atrodoties "
                "muguras guļā, veic 1,5 cm griezienu zem umbo, iekļūst prevezikālajā telpā. Neasā ceļā atbīdot "
                "taukaudus, izveido telpu priekšpusē peritonejam. Ievieto 10 mm troakāru, piešuj, ievada optiku, "
                "vizualizējot ekstraperitoneālo telpu, uzpilda līdz 15 mmHg. Redzes kontrolē hipogastrijā pa kreisi un "
                "pa labi ievada divus 12 mm troakārus. 4 cm pa labi no umbo ievieto 5 mm darba troakāru.",
            ]
            if lnd:
                text.append("Vizualizē a. iliaca externa dxt. Mobilizē un izdala limfmezglu paketi labajā iliakālajā un "
                            "obturatorajā zonā. Vizualizē a. iliaca externa sin. Mobilizē un izdala limfmezglu paketi "
                            "kreisajā iliakālajā un obturatorajā zonā.")
            text.append("Mobilizē prostatu un urīnpūšļa kaklu, incidē fascia endopelvina, pārdala ligamenta "
                        "puboprostatica" + (f", trulā ceļā atejot no nervu-asinsvadu kūlīša {nerve.lower()}"
                                            if nerve != "Nav" else "") + ".")
            text.append("Ar bipolāro elektrokoagulāciju pārdala plexus dorsalis, rūpīga hemostāze.")
            text.append("Atver urīnpūsli, vizualizē ostijas un prostatas malu. Pārdala urīnpūšļa gļotādu, mobilizē "
                        "priekšdziedzeri, abus ductus deferens un sēklas pūslīšus izdala, klipē un pārdala. Apeksa "
                        "rajonā atdala prostatu no uretras. Prostatas preparātu ievieto laparoskopijas maisiņā un uz "
                        "laiku novieto labajā apakšējā kvadrantā.")
            text.append(f"Izveido vezikouretrālo anastomozi ar nepārtrauktu šuvi. Ievieto urīnpūslī {fr} Fr katetru, "
                        f"balons {balloon} ml, savelk anastomozes šuves, pārbauda šuvju hermētismu. Hemostāze laba. "
                        "Izvelk prostatas preparātu. Evakuē troakārus."
                        + (" Retropubiski caur kontrapertūru pie anastomozes atstāj drenu." if drain else "")
                        + " Brūces sašuj pa kārtām. Sterils pārsējs.")
            text.append("Preparāts: " + ("lb. puses l/m; kr. puses l/m; prostata." if lnd else "prostata.")
                        + " Preparātu nosūta histoloģiskai izmeklēšanai.")
        elif op in ("Laparoscopic nephrectomy", "Laparoscopic kidney resection", "Laparoscopic pyeloplasty"):
            side = c1.radio("Puse", ["Labā", "Kreisā"], horizontal=True, key="op-lap-side")
            right = side == "Labā"
            adj = "labās" if right else "kreisās"        # labās puses ..., labās lumbotomijas guļā
            loc = "labajā" if right else "kreisajā"     # labajā pararektālajā rajonā
            lat = "dx." if right else "sin."
            kidney_acc = "labo nieri" if right else "kreiso nieri"
            access = (f"Pēc operācijas lauka 3x apdares ar polialkohola šķīdumu intubācijas narkozē, {pt_dat} atrodoties "
                      f"{adj} lumbotomijas guļā, veic griezienu 2 cm virs umbo {loc} pararektālajā rajonā.")
            if op == "Laparoscopic nephrectomy":
                fr = c2.selectbox("Katetrs (Fr)", ["14", "16", "18", "20"], index=2, key="op-nx-fr")
                drain = c3.checkbox("Drena nieres gultnē", False, key="op-nx-drain")
                name = f"Laparoskopiska {adj} puses nefrektomija"
                text = [
                    f"Ievieto urīnpūšļa ilgkatetru Foley tipa {fr} Fr, izdalās dzeltens urīns.",
                    access + " Ievada Veress adatu, uzpilda vēdera dobumu līdz 14 mmHg. Ievada 10 mm troakāru. "
                    f"Vizualizē vēdera dobumu. Veic griezienu un ievada 12 mm troakāru pa lin. parasternalis {lat} "
                    "Pa l. axillaris anterior subkostāli ievada 5 mm darba troakāru. 2 cm virs spina iliaca "
                    f"{lat} ievada 5 mm darba troakāru.",
                    f"Pārdala peritoneja parietālo lapiņu, vizualizē {kidney_acc}. Mobilizē un atbrīvo {adj} nieres "
                    "kājiņu, kuru pārdala ar endoskopisku griezējšuvēju. Mobilizē ureteru, ko klipē un pārdala. Veic "
                    f"nefrektomiju. Hemostāzes kontrole. Pagarinot griezienu, evakuē {kidney_acc} pa pararektālā "
                    "troakāra atveri." + (" Nieres gultnē pievada drenu." if drain else ""),
                    "Evakuē troakārus. Brūces sašuj pa kārtām.",
                    f"{pt_acc} vispārēji apmierinošā stāvoklī pārved uz nodaļu.",
                    "",
                    f"Preparāts: {'labā' if right else 'kreisā'} niere.",
                    "Preparātu nosūta histoloģiskai izmeklēšanai.",
                ]
            elif op == "Laparoscopic kidney resection":
                where = c2.text_input("Tumors (kur nierē)", "vidējā sienā, dorsāli", key="op-pn-where")
                margin = c3.number_input("Atkāpe (cm)", 0.1, 2.0, 0.5, step=0.1, format="%.1f", key="op-pn-margin")
                clamp = c1.checkbox("Nieres artērijas klemmēšana", False, key="op-pn-clamp")
                ischemia = (c2.number_input("Siltā išēmija (min)", 1, 60, 15, key="op-pn-isch") if clamp else 0)
                drain = c3.checkbox("Drena", True, key="op-pn-drain")
                name = f"Laparoskopiska {adj} nieres rezekcija"
                margin_text = f"{margin:.1f}".replace(".", ",")
                text = [
                    access + " Paņemot uz turētājiem, pārdala aponeirozi un peritoneju. Ievada Veress adatu, uzpilda "
                    "vēdera dobumu līdz 14 mmHg. Ievada 10 mm troakāru. Veic griezienu un ievada 12 mm troakāru pa lin. "
                    f"parasternalis {lat} Pa l. axillaris anterior subkostāli ievada 5 mm darba troakāru. 2 cm virs spina "
                    f"iliaca {lat} ievada 5 mm darba troakāru. Vizualizē vēdera dobumu.",
                    f"Pārdala peritoneja parietālo lapiņu, vizualizē {kidney_acc}. Nieres {where.strip().rstrip('.')} "
                    "vizualizē tumoru."
                    + (" Izdala nieres artēriju, uzliek asinsvadu klemmi." if clamp else "")
                    + f" Veic peritumorozu veidojuma rezekciju ar {margin_text} cm atkāpi. Rūpīga hemostāze, ar Z veida "
                    "šuvēm apšujot nieres defektu."
                    + (f" Noņem klemmi, siltās išēmijas laiks {ischemia} min." if clamp else "")
                    + " Hemostāze apmierinoša."
                    + (f" Caur darba kanālu virs spina iliaca {lat} ievieto drenu." if drain else "")
                    + " Evakuē rezecētos audus pa pararektālā troakāra atveri.",
                    "Evakuē troakārus. Brūces sašuj pa kārtām.",
                    "",
                    f"Preparāts: {adj} nieres rezekcijas audi.",
                    "Preparātu nosūta histoloģiskai izmeklēšanai.",
                ]
            else:
                stent = c2.checkbox("Uretera stents", True, key="op-pp-stent")
                drain = c3.checkbox("Drena pie anastomozes", True, key="op-pp-drain")
                suture = c1.text_input("Šuve", "3.0 Vicryl", key="op-pp-suture")
                name = f"{'Labās' if right else 'Kreisās'} puses laparoskopiska pieloplastika"
                text = [
                    access + " Paņemot uz turētājiem, pārdala aponeirozi un peritoneju. Ievada Veresa adatu, uzpilda "
                    "vēdera dobumu līdz 14 mmHg. Ievada 10 mm troakāru. 5 cm laterālāk no optikas troakāra veic "
                    "griezienu un ievada 12 mm troakāru. Pa l. axillaris anterior subkostāli ievada 5 mm darba "
                    f"troakāru. 2 cm virs spina iliaca {lat} ievada 5 mm darba troakāru. Vizualizē vēdera dobumu.",
                    f"Pārdala peritoneja parietālo lapiņu, vizualizē {kidney_acc}. Asā un neasā ceļā eksplorē nieres "
                    f"bļodiņu. Mobilizē ureteru. Ar {suture.strip()} uzšuj uretera dorsālo sienu. Pārdala ureteru, "
                    "iešķeļ to pa mugurējo sienu. Veic bļodiņas rezekciju. Ar nepārtrauktu šuvi veido anastomozes "
                    "mugurējo sienu." + (" Veic uretera stenta intubāciju." if stent else "")
                    + " Ar otru nepārtrauktu šuvi veido anastomozes priekšējo sienu un slēdz atlikušo defektu nieres "
                    "bļodiņā." + (" Pie anastomozes pievada drenu." if drain else ""),
                    "Evakuē troakārus. Brūces sašuj pa kārtām.",
                    f"{pt_acc} vispārēji apmierinošā stāvoklī pārved uz nodaļu.",
                ]
        elif op == "Scrotal mass excision":
            st.caption("Every scrotal mass needs an ultrasound first. A solid intratesticular mass is treated with an "
                       "inguinal orchidectomy (EAU), and a suspected paratesticular sarcoma with an inguinal approach – "
                       "not with a scrotal excision.")
            anesthesia = c1.selectbox("Anestēzija", ["spinālā anestēzijā", "vispārējā anestēzijā",
                                                     "vietējā anestēzijā ar Sol. Lidocaini 1%"], key="op-sm-an")
            position = c2.selectbox("Guļa", ["dorsālas litotomijas guļā", "muguras guļā"], key="op-sm-pos")
            place = c3.selectbox("Vieta", list(SCROTAL_PLACE), key="op-sm-place")
            side = c1.selectbox("Puse", list(SCROTAL_SIDE), key="op-sm-side")
            many = c2.radio("Veidojumi", ["Viens", "Vairāki"], horizontal=True, key="op-sm-many") == "Vairāki"
            size = c3.number_input("Izmērs (cm)", 0.0, 30.0, value=None, step=0.5, key="op-sm-size")
            incision = c1.selectbox("Griezums", ["virs veidojuma", "pa scrotum viduslīniju (raphe)",
                                                 "scrotum/perineālā rajonā", "elipsveida ādas griezums ap veidojumu"],
                                    key="op-sm-inc")
            looks = {  # (before the noun, singular / plural, after the noun)
                "iekapsulēts, mīksti elastīgs": ("iekapsulētu, mīksti elastīgu", "iekapsulētus, mīksti elastīgus", ""),
                "cistisks, ar biezu, baltu saturu": ("cistisku", "cistiskus", " ar biezu, baltu saturu"),
                "mīksts, lipomatozs": ("mīkstu, lipomatozu", "mīkstus, lipomatozus", ""),
                "blīvs, ar skaidrām robežām": ("blīvu", "blīvus", " ar skaidrām robežām"),
            }
            look = c2.selectbox("Veidojuma raksturs", list(looks), key="op-sm-look")
            free = c3.checkbox("Nav saistīts ar sēklinieku un sēklas virvīti", True, key="op-sm-free")
            drain = c1.checkbox("Drenāža (gumijas strēmele)", False, key="op-sm-drain")
            closure = c2.selectbox("Šuves", ["Zemādas šuve. Ādas šuves.", "Zemādas šuve. Intrakutāna uzsūcoša šuve."],
                                   key="op-sm-close")
            name = scrotal_words({"Scrotal mass location": place, "Scrotal mass number": "Multiple" if many else "Single"})[2]
            name = name[0].upper() + name[1:]
            noun = "veidojumus" if many else "veidojumu"
            size_text = (f" Ø {size:g} cm".replace(".", ",") if size else "")
            where = "Zemādā" if place != "Spermatic cord / paratesticular" else "Paratestikulāri, sēklas virvītes apvidū,"
            cut = ("veic elipsveida ādas griezienu ap veidojumu" if incision.startswith("elipsveida")
                   else f"izdara griezienu {incision.replace('virs veidojuma', 'virs veidojumiem') if many else incision}")
            text = [
                f"{pt_acc.replace('Pacientu', 'Pacients').replace('Pacienti', 'Paciente')} atrodas {position}. Pēc "
                f"operācijas lauka apdares {anesthesia} {cut} ({SCROTAL_SIDE[side]}). {where} vizualizē "
                f"{looks[look][1 if many else 0]} {noun}{size_text}{looks[look][2]}"
                + ((", kas nav saistīti ar sēklinieku un sēklas virvīti" if many else
                    ", kas nav saistīts ar sēklinieku un sēklas virvīti") if free else "")
                + f". Neasā ceļā {'tos' if many else 'to'} izdala un evakuē. Rūpīga hemostāze."
                + (" Brūcē atstāj gumijas strēmeli." if drain else "") + f" {closure} Sterils pārsējs.",
                f"{'Veidojumus' if many else 'Veidojumu'} nosūta patohistoloģiskai izmeklēšanai.",
            ]
        elif op == "Hydrocele plasty":
            side = c1.radio("Puse", ["Labā", "Kreisā"], horizontal=True, key="op-h-side")
            technique = c2.selectbox("Tehnika", ["Izvēršana (Winkelmann)", "Ekscīzija (Bergmann)", "Plikācija (Lord)"],
                                     key="op-h-tech")
            fluid = c3.selectbox("Šķidrums", ["serozs", "hemorāģisks", "duļķains"], key="op-h-fluid")
            drain = c1.checkbox("Drenāža", True, key="op-h-drain")
            ward = c2.text_input("Pārved uz", "nodaļu", key="op-h-ward")
            short = "lb." if side == "Labā" else "kr."
            step = {"Izvēršana (Winkelmann)": "Veic tunica vaginalis sieniņas incīziju un izvēršanu uz āru, fiksē ar "
                                              "nepārtrauktu šuvi.",
                    "Ekscīzija (Bergmann)": "Veic tunica vaginalis parietālās lapiņas ekscīziju, brīvās malas apšuj ar "
                                            "nepārtrauktu šuvi.",
                    "Plikācija (Lord)": "Veic tunica vaginalis plikāciju ar atsevišķām šuvēm."}[technique]
            name = f"{'Labās' if side == 'Labā' else 'Kreisās'} puses hidrocēles plastika"
            text = [f"Pēc op. lauka apdares veic griezienu scrotum ādā pa viduslīniju. Izdala {short} puses testis līdz "
                    f"apvalkiem. Konstatē hidrocēli. Pārdala to, izdalās {fluid} šķidrums. {step} Rūpīga hemostāze."
                    + (" Drena." if drain else "") + " Ādu slēdz ar atsevišķām šuvēm. Sterils pārsējs.",
                    f"{pt_acc} vispārēji apmierinošā stāvoklī pārved uz {ward.strip()}."]
        elif op == "Prostatic abscess drainage":
            text, name = prostatic_abscess_operation(c1, c2, c3, pt_acc)
        elif op in ("Orchiectomy (scrotal)", "Inguinal orchifuniculectomy", "Epididymectomy",
                    "Scrotal / funicular revision, drainage"):
            text, name = scrotal_infection_operation(op, c1, c2, c3, pt_acc)
        elif op == "Radical cystectomy":
            text, name = cystectomy_operation(woman, c1, c2, c3, pt_dat, pt_acc)
        elif op == "Cystoscopy (diagnostic) ± biopsy":
            text, name = cystoscopy_operation(c1, c2, c3, pt_acc)
        elif op == "Dorsumcision (dorsal slit)":
            reason = c1.selectbox("Indikācija", ["parafimoze", "fimoze"], key="op-ds-reason")
            anesthesia = c2.selectbox("Anestēzija", ["lokālā anestēzijā", "spinālā anestēzijā", "vispārējā anestēzijā"],
                                      key="op-ds-an")
            suture = c3.text_input("Šuves", "4.0 Vicryl", key="op-ds-suture")
            ward = c1.selectbox("Pēc operācijas", ["pārved uz palātu", "izraksta ambulatori"], key="op-ds-ward")
            name = "Dorsumcīzija"
            ring = "parafimozes (iespīlējuma) gredzenu" if reason == "parafimoze" else "fimozes gredzenu"
            text = [
                f"Pēc op. lauka apdares ar \"Braunol\" {anesthesia} iešķeļ {ring} uz plkst. 12.00. "
                + ("Priekšādiņu reponē pāri dzimumlocekļa galviņai. " if reason == "parafimoze" else "Atver priekšādiņu. ")
                + f"Brūci apstrādā ar \"Braunol\". Brūces malas apšuj ar {suture.strip()} atsevišķām šuvēm. "
                "Hemostāze apmierinoša. Sterils pārsējs. "
                + (f"{pt_acc} vispārējā apmierinošā stāvoklī pārved uz palātu." if ward == "pārved uz palātu"
                   else f"{pt_acc} vispārējā apmierinošā stāvoklī izraksta ambulatori.")
            ]
        elif op == "Circumcision":
            anesthesia = c1.selectbox("Anestēzija", ["spinālā anestēzijā", "vispārējā anestēzijā", "vietējā anestēzijā"], key="op-c-an")
            antiseptic = c2.text_input("Lauka apdare ar", "polispirtu šķīdumu", key="op-c-anti")
            margin = c3.number_input("Atkāpe no corona glandis (mm)", 2, 15, 5, key="op-c-margin")
            suture = c1.text_input("Šuves", "SafilQuick 4.0", key="op-c-suture")
            frenulum = c2.checkbox("Frenuloplastika", key="op-c-fren")
            ward = c3.text_input("Pārved uz", "nodaļu", key="op-c-ward")
            text = [
                f"Pacients {anesthesia}. Pēc op. lauka apdares ar {antiseptic.strip()} izdara cirkulāru griezienu "
                "Corona glandis rajonā, izpreparē preputium ārējo lapiņu. Iešķeļ preputium pret 12.00, atvirza to pār "
                "glans penis, notīra smegmas kārtu ar Sol. Braunoli. Veic preputium iekšējās lapiņas rezekciju kopā ar "
                f"preputium ar aptuveni {margin} mm atkāpi no corona glandis."
                + (" Veic frenuloplastiku." if frenulum else "")
                + f" Rūpīga hemostāze. Brūci slēdz ar atsevišķām {suture.strip()} šuvēm. Sterils pārsējs. Pacientu "
                f"vispārēji apmierinošā stāvoklī pārved uz {ward.strip()}."
            ]
        complications = st.text_input("Intraoperatīvas komplikācijas / papildus atrade (optional)", key=f"op-extra-{op}")
    if complications.strip():
        text.append(complications.strip().rstrip(".") + ".")
    if op == "TURU" and not st.session_state.get("op-turu-where"):
        st.info("Choose where the tumour is.")
        return
    header = render_doc_header("op")
    output = ((f"Operācija: {name}\n" if name else "")
              + (f"Anesteziologs: {header['anesth']}\n" if header.get("anesth") else "")
              + ("\n" if name or header.get("anesth") else "") + "Operācijas apraksts:\n\n" + "\n".join(text))
    st.divider()
    st.text_area("Operācijas apraksts", output, height=420, key=f"op-out-{hash(output)}")
    word_button("Word file (Operācijas apraksts)", lambda: operation_docx(header, name, text),
                f"Operacijas_apraksts_{header.get('date', '').strip('.').replace('.', '-') or 'bez_datuma'}.docx",
                "op-word")


# ---------------------------------------------------------------
# 3. PAGE
# ---------------------------------------------------------------

st.set_page_config(page_title="Discharge helper", layout="wide")
st.markdown("<h1 style='text-align: center;'>Discharge helper</h1>", unsafe_allow_html=True)
st.caption(f"Build: 2026-10-26-prostatitis · Running file: {Path(__file__).resolve()}")

# Make the two sex buttons tall rectangles
st.markdown(
    "<style>"
    ".st-key-sex_woman button, .st-key-sex_man button {height: 3.4rem; font-size: 1.25rem;}"
    "</style>",
    unsafe_allow_html=True,
)

# Bar 0: which document to write
doc_type = st.segmented_control(
    "Document", ["Discharge", "Primary evaluation", "Decursus", "Operation description"], default="Discharge",
    selection_mode="single", key="doc_type", width="stretch",
) or "Discharge"

# Bar 1: where the patient is discharged from (only for discharge papers)
department = None if doc_type != "Discharge" else st.segmented_control(
    "Discharged from",
    options=list(DEPARTMENTS),
    format_func=lambda key: DEPARTMENTS[key],
    selection_mode="single",
    default=None,
    key="department",
    width="stretch",
)

# Bar 2: patient sex, two equal rectangles (woman left, man right)
def set_sex(value):
    st.session_state["sex"] = value


st.session_state.setdefault("sex", None)
sex = st.session_state["sex"]
st.caption("Patient")
col_woman, col_man = st.columns(2)
with col_woman:
    st.button("♀  Woman", key="sex_woman", on_click=set_sex, args=("Woman",),
              type="primary" if sex == "Woman" else "secondary", width="stretch")
with col_man:
    st.button("♂  Man", key="sex_man", on_click=set_sex, args=("Man",),
              type="primary" if sex == "Man" else "secondary", width="stretch")

# Keep the Word header fields (and the surgeons) when you switch between pages
for _key in HDR_KEYS + ["pe-surgeons", "pe-nurse"]:
    if _key in st.session_state:
        st.session_state[_key] = st.session_state[_key]

# Primary evaluation and decursus have their own pages
if doc_type != "Discharge":
    if sex is None:
        st.info("Choose the patient's sex first.")
    elif doc_type == "Primary evaluation":
        render_primary_evaluation(sex)
    elif doc_type == "Operation description":
        render_operation(sex)
    else:
        render_decursus(sex)
    st.stop()

# The central bar: type a name or a code, pick one or several
selected = st.multiselect(
    "Search diagnosis",
    options=list(DIAGNOSES),
    format_func=lambda code: f"{code} - {DIAGNOSES[code]}",
    placeholder="Start typing a diagnosis or ICD-10 code, e.g. prostate or N40",
)

choices = {}
chosen_blocks = {}

if selected:
    col_main, col_side = st.columns([3, 2], gap="large")

    # -----------------------------------------------------------
    # Right side: recommendation blocks you can drag in
    # -----------------------------------------------------------
    with col_side:
        st.subheader("Recommendation blocks")
        classic = st.checkbox("Use the old panel (if the new one does not load)", key="classic_panel_v2")
        any_blocks = False
        for code in selected:
            blocks = BLOCKS.get(code)
            if not blocks:
                continue
            any_blocks = True
            st.markdown(f"**{code} - {DIAGNOSES[code]}**")
            by_id = {b["id"]: b for b in blocks}
            label_to_id = {block_label(b): b["id"] for b in blocks}

            if not classic:
                st.caption("Drag blocks into Chosen, or hold Ctrl and click a block to move it "
                           "between the lists. Hover a block to read its notes.")
                chosen_ids = block_picker(code, blocks)
            elif HAS_SORTABLES:
                st.caption("Drag blocks from Available to Chosen. "
                           "The order in Chosen is the order in the text.")
                result = sort_items(
                    [{"header": "Available", "items": list(label_to_id)},
                     {"header": "Chosen", "items": []}],
                    multi_containers=True,
                    direction="vertical",
                    key=f"blocks-{code}",
                )
                try:
                    chosen_ids = [label_to_id[t] for t in result[1]["items"] if t in label_to_id]
                except (TypeError, IndexError, KeyError):
                    chosen_ids = []
            else:
                st.caption("Drag and drop is not installed, so pick the blocks from a list.")
                picked = st.multiselect("Choose blocks", list(label_to_id), key=f"blocks-{code}")
                chosen_ids = [label_to_id[t] for t in picked]

            chosen_blocks[code] = [by_id[i] for i in chosen_ids if i in by_id]

            for warning in BLOCK_WARNINGS.get(code, []):
                if all(i in chosen_ids for i in warning["ids"]):
                    st.warning(warning["text"])

            with st.expander("Clinical notes"):
                for b in blocks:
                    st.markdown(f"**{b['title']}** ({b['group']})  \n{b['note']}")
        if not any_blocks:
            st.caption("No blocks for the selected diagnoses yet.")

    # -----------------------------------------------------------
    # Left side: modifiers and the discharge text
    # -----------------------------------------------------------
    with col_main:
        for code in selected:
            fixed = fixed_values(code, department)
            with st.container(border=True):
                st.subheader(f"{code} - {DIAGNOSES[code]}")
                # lets sections and rules depend on the department and the patient's sex
                choices[code] = {"_department": department, "_sex": sex}
                layout = DIAGNOSIS_MODIFIERS.get(code, DEFAULT_MODIFIERS)
                # A list of (title, [modifiers]) pairs gives sections; a plain list is one block
                sections = layout if layout and isinstance(layout[0], tuple) else [(None, layout)]

                for section in sections:
                    title, groups = section[0], section[1]
                    condition = section[2] if len(section) > 2 else {}
                    if not all(condition_met(choices[code].get(other), wanted)
                               for other, wanted in condition.items()):
                        continue  # this section does not apply to the current choices
                    holder = st.expander(title, expanded=True) if title else st.container()
                    with holder:
                        for group in groups:
                            spec = MODIFIERS[group]

                            # Skip modifiers that only appear when another modifier has a value
                            show_if = spec.get("show_if", {})
                            if not all(condition_met(choices[code].get(other), wanted)
                                       for other, wanted in show_if.items()):
                                continue

                            if group in fixed:  # set automatically, so do not ask
                                choices[code][group] = fixed[group]
                                continue

                            key = f"{code}-{group}"
                            if spec["type"] == "single":
                                value = st.radio(group, spec["options"], key=key,
                                                 index=spec.get("default_index"), horizontal=True)
                            elif spec["type"] == "date":
                                value = st.date_input(group, value=None, key=key, format="DD.MM.YYYY")
                                if value is None and spec.get("optional"):
                                    value = ""  # an optional date may stay empty
                            elif spec["type"] == "number":
                                value = st.number_input(group, min_value=spec["min"], max_value=spec["max"],
                                                        value=None, step=spec["step"], format="%g", key=key)
                                if value is None and spec.get("optional"):
                                    value = ""  # an optional number may stay empty
                            elif spec["type"] == "multi":
                                picked = st.multiselect(group, spec["options"],
                                                        default=spec.get("default_selected", []), key=key)
                                if spec.get("custom"):
                                    extra = st.text_input(f"{group} - other (optional)", key=f"{key}-custom",
                                                          help=spec.get("custom_help"))
                                    if extra.strip():
                                        picked = picked + [extra.strip()]
                                value = picked if (picked or spec.get("optional")) else None
                            elif spec["type"] == "text":
                                text = st.text_input(group, key=key, placeholder=spec.get("placeholder", ""))
                                value = text.strip() if (text.strip() or spec.get("optional")) else None
                            elif spec["type"] == "table":
                                st.markdown(f"**{group}**")
                                value = table_input(spec, key)
                            else:  # toggle
                                value = st.checkbox(group, key=key, value=spec.get("default", False))
                            choices[code][group] = value

                if fixed:
                    shown = {g: v for g, v in fixed.items() if g in choices[code]}
                    if shown:
                        st.caption("Set automatically: " + ", ".join(f"{g}: {v}" for g, v in shown.items()))

                # Follow-up tests: suggested from the choices above, nothing is ticked by default
                tests, tests_by = render_tests(code, choices[code])
                choices[code]["_tests"] = tests
                choices[code]["_test_ids"] = [t["id"] for t in tests]
                choices[code]["_tests_by"] = tests_by

        # -------------------------------------------------------
        # Discharge text
        # -------------------------------------------------------
        st.divider()
        st.header("Discharge text")

        if department is None or sex is None:
            st.info("Choose where the patient is discharged from and the patient's sex to see the text.")
        else:
            for code in selected:
                st.subheader(f"{code} - {DIAGNOSES[code]}")
                values = choices[code]
                values["_blocks"] = [b["id"] for b in chosen_blocks.get(code, [])]  # for the safety checks

                missing = [g for g, v in values.items() if v is None and not g.startswith("_")]
                if missing:
                    st.info("Choose all options above to see the text. Not chosen yet: " + ", ".join(missing))
                    continue

                template = find_template(code, department, values)
                if template is None:
                    st.warning("No text template for this combination of choices yet.")
                    continue

                # Safety checks: some situations need a clinician's review instead of a routine text
                checker = CHECKS.get(code)
                blockers, cautions = checker(values) if checker else ([], [])
                if blockers:
                    st.error("Clinical review is needed before a routine discharge text can be written:\n\n"
                             + "\n".join(f"- {b}" for b in blockers))
                    continue
                for caution in cautions:
                    st.warning(caution)

                block_texts = [b["text"] for b in chosen_blocks.get(code, [])]
                if template.get("builder"):
                    parts = BUILDERS[template["builder"]](sex, values, block_texts)
                    dx_text = parts["clinical_diagnosis"]
                    epi_text = parts["epicrisis"]
                    recs = parts["recommendations"]
                else:
                    forms = build_forms(sex, values)
                    dx_text = template.get("clinical_diagnosis", "").format(**forms)
                    epi_text = template["epicrisis"].format(**forms)
                    recs = add_blocks(template["recommendations"], block_texts)
                block_sig = "|".join(b["id"] for b in chosen_blocks.get(code, []))
                # The text boxes are editable, so you can adjust the wording before
                # copying it. Changing the choices above regenerates the texts.
                digest = hashlib.md5(repr(sorted(values.items())).encode("utf-8")).hexdigest()
                sig = f"{template['id']}|{sex}|{digest}"

                tab_dx, tab_epi, tab_rec = st.tabs(["Clinical diagnosis", "Epicrisis", "Recommendations"])
                with tab_dx:
                    if not dx_text:
                        st.info("No clinical diagnosis text for this template yet. You can type it in the box.")
                    st.text_area("Clinical diagnosis", dx_text, height=170, key=f"dx-{sig}")
                with tab_epi:
                    st.text_area("Epicrisis", epi_text, height=420, key=f"epi-{sig}")
                with tab_rec:
                    st.text_area("Recommendations", recs, height=600,
                                 key=f"rec-{sig}-{block_sig}")

        with st.expander("Current selection (technical view)"):
            technical = {"department": department, "sex": sex, "diagnoses": choices}
            st.json(json.loads(json.dumps(technical, default=str)))
