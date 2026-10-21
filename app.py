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
    "N20": "Calculus of kidney and ureter (kidney stone disease)",
    "N23": "Unspecified renal colic",
    "N30.0": "Acute cystitis",
    "N32.0": "Bladder-neck obstruction",
    "N35": "Urethral stricture",
    "N39.0": "Urinary tract infection, site not specified",
    "N40": "Benign prostatic hyperplasia",
    "N41.0": "Acute prostatitis",
    "N43": "Hydrocele and spermatocele",
    "N45": "Orchitis and epididymitis",
    "N47": "Phimosis / redundant prepuce",
    "R31": "Hematuria",
    "C61": "Malignant neoplasm of prostate",
    "C64": "Malignant neoplasm of kidney",
    "C67": "Malignant neoplasm of bladder",
    "Z43.6": "Attention to artificial opening of urinary tract (nephrostomy change)",
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
    "Lithotripsy": {
        "type": "single",
        "options": ["Laser lithotripsy", "Pneumatic lithotripsy", "Basket extraction", "None"],
    },
    "DJ stent placed": {"type": "toggle"},
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
        ("1. Admission", ["Stay", "Surgery", "Conservative reason"]),
        ("2. Findings", ["Phimosis degree", "Balanoposthitis", "Frenulum breve"]),
        ("3. Operation", ["Surgery type", "Frenuloplasty", "Operation date", "Anesthesia"], {"Surgery": True}),
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
    "C67": [
        ("1. Admission and tumours",
         ["Admission", "Tumour verification", "Tumour number", "Tumour location"]),
        ("2. Operation and course",
         ["TURU type", "Procedure date", "Operation details", "Intravesical treatment",
          "Intravesical treatment (other)", "Postoperative course", "Postoperative complications",
          "Catheter removal", "Catheter removed on", "Planned catheter removal (optional)",
          "Voiding after catheter removal", "Histology", "Tumour stage", "N stage", "M stage",
          "Tumour grade", "Concomitant CIS", "Histology details"]),
        ("3. Recommendations",
         ["Discharge paper", "Activity restriction", "Urology visit", "Oncology council", "Oncology council date"]),
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
         ["Admission", "Stricture location", "Stricture length (cm)", "Stricture cause", "Previous stricture treatment"]),
        ("2. Epicystostomy at NMC",
         ["Failed urethral catheterisation", "Ultrasound guided", "Procedure date", "Catheter Fr (N35)",
          "Drained volume (ml)", "Urine appearance", "Creatinine raised (N35)", "Fever at NMC", "Urology consultation"],
         {"_department": "NMC"}),
        ("2. Operation",
         ["Stricture operation", "N35 anesthesia", "Procedure date", "Postoperative course",
          "Postoperative complications", "Tissue for histology", "Histology result (N35)"],
         {"_department": "4"}),
        ("3. Catheter and voiding",
         ["Catheter Fr (N35)", "Catheter at discharge (N35)", "Planned catheter removal (N35)",
          "Voiding after removal (N35)", "Self-dilation taught"],
         {"_department": "4"}),
        ("4. Other", ["Investigations (N35)"], {"_department": "4"}),
    ],
    "Z43.6": [
        ("1. Nephrostomy and co-diagnoses", ["Nephrostomy side", "Nephrostomy co-diagnoses"]),
        ("2. Nephrostomy change",
         ["Procedure date", "Nephrostomy Ch", "Nephrostomy function", "Irrigation performed",
          "Fixed with skin sutures", "Infection signs at change", "Fever (48)"]),
        ("3. Recommendations", ["Supervising doctor", "Next nephrostomy change", "Investigations (48)"]),
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
    "URS": "ureterorenoskopija",
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
STRICTURE_DX = {"Meatus": "Meatus striktūra", "Penile urethra": "Urīnizvadkanāla penilās daļas striktūra",
                "Bulbar urethra": "Urīnizvadkanāla bulbārās daļas striktūra",
                "Membranous (posterior) urethra": "Urīnizvadkanāla membranozās daļas striktūra",
                "Not specified": "Urīnizvadkanāla striktūra"}
STRICTURE_EPI = {"Meatus": "urīnizvadkanāla ārējās atveres sašaurinājumu",
                 "Penile urethra": "urīnizvadkanāla penilās daļas striktūru",
                 "Bulbar urethra": "urīnizvadkanāla bulbārās daļas striktūru",
                 "Membranous (posterior) urethra": "urīnizvadkanāla membranozās daļas striktūru",
                 "Not specified": "urīnizvadkanāla striktūru"}
STRICTURE_OP = {  # (as done, after "st. pēc")
    "Meatotomy": ("meatotomija", "meatotomijas"),
    "Meatoplasty": ("meatoplastika", "meatoplastikas"),
    "Urethral dilation": ("urīnizvadkanāla dilatācija", "urīnizvadkanāla dilatācijas"),
    "Internal urethrotomy (DVIU)": ("iekšējā optiskā uretrotomija", "iekšējās optiskās uretrotomijas"),
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
]
BLOCK_WARNINGS["N47"] = BLOCK_WARNINGS["N10"]
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
        "when": {},
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
    if v["Lithotripsy"] in LITHO_TEXT:
        with_parts.append(LITHO_TEXT[v["Lithotripsy"]][0])
        name_parts.append(LITHO_TEXT[v["Lithotripsy"]][1])
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


def c67_checks(v):
    blockers, cautions = [], []
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
    recs.append(f"Vizīte pie urologa ambulatori: {WEEKS_TEXT[v['Urology visit']]} pēc operācijas.")
    recs += tests_lines(v, "plain", after_op=True)  # e.g. cystoscopy 3 months after the operation
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
CHECKS.update({"N10": n10_checks, "N47": n47_checks, "T83": t83_checks})


# ---------------------------------------------------------------
# 2h. URETHRAL STRICTURE (N35): checks and text builder (EAU guideline based)
# ---------------------------------------------------------------


def n35_parts(sex, v, blocks=()):
    f = SEX_FORMS["Man"]
    op = v["Stricture operation"]
    op_nom, op_gen = STRICTURE_OP[op]
    day = fmt_date(v["Procedure date"])
    length = v.get("Stricture length (cm)")
    catheter_in = v["Catheter at discharge (N35)"] == "In place"

    # Clinical diagnosis (department format: "Meatus striktūra, st. pēc meatotomijas, urīna ilgkatetrs (date).")
    first = STRICTURE_DX[v["Stricture location"]]
    if v.get("Previous stricture treatment"):
        first = "Recidivējoša " + first[0].lower() + first[1:]
    if length not in (None, ""):
        first += f" ({length:g} cm)"
    diagnosis = f"{first}, st. pēc {op_gen}" + (", urīna ilgkatetrs" if catheter_in else "") + f" ({day})."
    cause = v.get("Stricture cause")
    if cause in STRICTURE_CAUSE_DX:
        diagnosis += "\n" + STRICTURE_CAUSE_DX[cause]

    # Epicrisis
    reason = STRICTURE_EPI[v["Stricture location"]]
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
        f"{f['Pacients']} tiek {f['gatavots']} operatīvai terapijai, {day} {N35_ANESTHESIA[v['N35 anesthesia']]} tiek "
        f"veikta {op_nom}, tiek ievadīts Fr{v['Catheter Fr (N35)']} izmēra urīna ilgkatetrs.",
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
    epicrisis = " ".join(sentences) + f"\n\nIzmeklējumi: {investigations}.\n\nOperācijas: {op_nom} ({day})."

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
        recs += ["• Režīms – saudzējošs 1 nedēļu no operācijas dienas, izvairīties no pārmērīgas fiziskas slodzes."]
    if catheter_in:
        recs += ["• Urīna ilgkatetra aprūpe: uzturēt noslēgtu drenāžas sistēmu, maisu novietot zem urīnpūšļa līmeņa, "
                 "sekot urīna attecei."]
        removal = v.get("Planned catheter removal (N35)")
        when = fmt_date(removal) if removal else "pēc ārstējošā urologa norādījumiem"
        recs.append(f"• Urīna ilgkatetra evakuācija plānota {when}.")
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
            cautions.append("Recurrent stricture after endoscopic treatment: repeat urethrotomy/dilation rarely "
                            "succeeds; EAU advises offering urethroplasty.")
        if location in ("Penile urethra", "Membranous (posterior) urethra") or (length not in (None, "") and length > 2):
            cautions.append("Urethrotomy works best for short (under 2 cm) bulbar strictures; for penile or longer "
                            "strictures EAU favours urethroplasty.")
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
    "panel": {"group": "Blood", "title": "Extended blood panel (department list)",
              "lv": "CRO, LDH, sārmainā fosfatāze, kreatinīns, urīnskābe, Ca, P, Na, K, glikoze, ASAT, ALAT, kopējais "
                    "bilirubīns, 25-OH-vitamīns D (D3+D2)"},
    # imaging
    "us_kb": {"group": "Imaging", "title": "Kidney and bladder ultrasound", "lv": "nieru un urīnpūšļa USS"},
    "us_k": {"group": "Imaging", "title": "Kidney ultrasound", "lv": "nieru USS"},
    "us_scrotum": {"group": "Imaging", "title": "Scrotal ultrasound", "lv": "sēklinieku maisiņa USS"},
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
        out.append(_s("urethrocysto", "consider", "symptoms", "If the stream weakens or the symptoms return."))
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
    "Circumcision – phimosis": ("Planned", "cirkumcīzija", "fimozi", "Fimoze", "", ""),
    "Hydrocele plasty": ("Planned", "kreisās puses hidrocēles plastika", "kreisās puses hidrocēli",
                         "Kreisās puses hidrocēle", "", ""),
    "URS/FURS – kidney stones": ("Planned", "URS/FURS ar kontakta litotripsiju", "nierakmeņu slimību",
                                 "Nierakmeņu slimība", "", ""),
    "Laparoscopic radical prostatectomy – prostate cancer": ("Planned", "laparoskopiska radikāla prostatektomija",
                                                             "prostatas CA", "Prostatas CA", "", ""),
    "Nephrectomy – kidney tumour": ("Planned", "nefrektomija", "nieres jaunveidojumu", "Nieres jaunveidojums", "", ""),
    "Laparoscopic pyeloplasty – PUJ stricture": ("Planned", "laparoskopiska pieloplastika",
                                                 "labās puses pieloureterālā segmenta striktūru",
                                                 "Labā pieloureterālā segmenta striktūra", "", ""),
    "Urethroplasty – urethral stricture": ("Planned", "uretroplastika", "urīnizvadkanāla striktūru",
                                           "Urīnizvadkanāla striktūra", "", ""),
    "Hematuria – acute admission (no operation)": (
        "Acute", "", "", "Makrohematūrija", "asins piejaukumu urīnā",
        "Izmeklēts NMC. Veic CT ar i/v kontrastvielu."),
    "Acute admission, other (no operation)": ("Acute", "", "", "", "", "Izmeklēts NMC."),
    "Other (type it)": ("Planned", "", "", "", "", ""),
}
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
HDR_KEYS = ["hdr-patient", "hdr-history", "hdr-dept", "hdr-surgeons", "hdr-nurse", "hdr-date-op", "hdr-date-pe"]


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
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Urologa apskate")
    run.bold, run.font.size = True, Pt(14)
    _docx_line(doc, [("Datums", h.get("date", ""))])
    _docx_line(doc, [("Pacients", h.get("patient", ""))], indent=Cm(0))
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
    """Header fields that only go into the Word file. kind: "op" or "pe"."""
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
        if kind == "op":
            h["history"] = c1.text_input("Vēstures nr", key="hdr-history").strip()
            h["dept"] = c2.text_input("Nodaļa", key="hdr-dept").strip()
            h["surgeons"] = c1.text_input("Ķirurgs/ Asistents", key="hdr-surgeons").strip()
            h["nurse"] = c2.text_input("Operācijas māsa", key="hdr-nurse").strip()
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
        admission = st.radio("Admission", ["Planned", "Acute"], index=0 if adm_d == "Planned" else 1,
                             horizontal=True, key=f"pe-admission-{preset}")
        operation = st.text_input("Operation (empty if no operation)", op_d, key=f"pe-op-{preset}")
        if operation.strip():
            dx_acc = st.text_input("Reason, as it follows 'sakarā ar iepriekš diagnosticētu ...'", acc_d,
                                   key=f"pe-dxacc-{preset}")
            complaints = st.text_input("Complaints (optional)", key=f"pe-compl-{preset}")
            course = ""
        else:
            dx_acc = ""
            complaints = st.text_input("Complaints, as they follow 'ar sūdzībām par ...'", compl_d,
                                       key=f"pe-compl-{preset}")
            course = st.text_area("Before admission (NMC, investigations, findings)", course_d,
                                  key=f"pe-course-{preset}", height=90)
        diagnosis = st.text_area("Diagnosis (Slēdziens)", dx_d, key=f"pe-dx-{preset}", height=70)
    with st.expander("2. Past diseases", expanded=True):
        other = st.text_area("Citas slimības (empty = Noliedz)", key="pe-other", height=70)
        c1, c2, c3 = st.columns(3)
        tbc = c1.radio("TBC", ["Noliedz", "Apstiprina"], key="pe-tbc")
        hep = c2.radio("Hepatīti", ["Noliedz", "Apstiprina"], key="pe-hep")
        allergy = c3.radio("Medikamentu nepanesība", ["Noliedz", "Ir"], key="pe-allergy")
        allergy_text = st.text_input("Which medications", key="pe-allergy-text") if allergy == "Ir" else ""
        operations = st.text_area("Pārciestas operācijas (empty = Nav)", key="pe-ops", height=70)
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
        genitals = c3.text_input("Ārējie dzimumorgāni", "neizmainīti", key="pe-genitals")
    with st.expander("4. Conclusion and plan", expanded=True):
        if operation.strip():
            indication = st.text_input("Indication (empty = the diagnosis)", key="pe-indication",
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
        [("Citas slimības:", other.strip() or "Noliedz")],
        [("TBC", tbc)],
        [("Hepatīti", hep)],
        [("Medikamentu nepanesība", "Noliedz" if allergy == "Noliedz" else (allergy_text.strip() or "Ir"))],
        [("Pārciestas operācijas:", operations.strip() or "Nav")],
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
    ]
    as_text = lambda line: "  ".join(f"{label}  {value}" for label, value in line)
    dx_line = diagnosis.strip().rstrip(".") + "."
    lines = ([reason, "", "Citas pārciestas slimības:"] + [as_text(l) for l in history]
             + ["", "Objektīvais stāvoklis:"] + [as_text(l) for l in objective]
             + ["", f"Diagnoze, Slēdziens: {dx_line}"])
    if operation.strip():
        op = operation.strip()
        dx = diagnosis.strip().rstrip(".")
        parts = [f"{'Pacientei' if sex == 'Woman' else 'Pacientam'} indicēta {op}.",
                 f"Indikācijas operācijai – {(indication.strip().rstrip('.') or dx[:1].lower() + dx[1:])}."]
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
    "complaints": "Pacients jūtas apmierinoši, sūdzību nav.",
    "pulse": 70, "rhythm": "ritmisks", "bp": "120/80", "resp": 16, "breath": "vezikulāra", "temp": 36.6,
    "abdomen": "Vēders palpatori mīksts, nesāpīgs.", "voiding": "", "notes": "",
    "cath_ml": None, "cath_look": "(no catheter)", "drain_ml": None, "drain_look": "(no drain)", "drain_out": False,
    "dressing": False,
    "ending": "Turpina ārstēšanu stacionārā.",
}
CATH_LOOKS = ["(no catheter)", "dzeltens", "dzidrs", "viegli sārts", "viegli sangvinozs", "sangvinozs", "ar recekļiem", "duļķains"]
DRAIN_LOOKS = ["(no drain)", "serozi", "serozi-sangvinozi", "sangvinozi-serozi", "serosangvinozi", "sangvinozi"]
DRESSING_TEXT = ("Pārsiešana. Operācijas brūces tīras, sausas, dzīst primāri. Ādu apstrādā ar “Braunol”. "
                 "Sterili pārsēji.")
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


def render_decursus_cards(sex):
    """The decursus as cards (one per day). Returns the decursus text."""
    st.session_state.setdefault("dd-count", 1)
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
            complaints = st.text_input("Complaints", key=f"dd{i}-complaints", label_visibility="collapsed")
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
            k = st.columns([2, 3, 2, 3, 2, 2], vertical_alignment="bottom")
            cath_ml = k[0].number_input("Katetrs ml/dnn", 0, 10000, step=100, key=f"dd{i}-cath_ml")
            cath_look = k[1].selectbox("Urīns pa katetru", CATH_LOOKS, key=f"dd{i}-cath_look")
            drain_ml = k[2].number_input("Drena ml/dnn", 0, 5000, step=10, key=f"dd{i}-drain_ml")
            drain_look = k[3].selectbox("Izdalījumi pa drenu", DRAIN_LOOKS, key=f"dd{i}-drain_look")
            drain_out = k[4].checkbox("Drenu evakuē", key=f"dd{i}-drain_out")
            dressing = k[5].checkbox("Pārsiešana", key=f"dd{i}-dressing")
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
        if complaints.strip():
            parts.append(sexed(complaints.strip(), sex))
        parts.append(f"Obj. Vispārējais stāvoklis {status}.")
        parts.append(f"P {pulse:g} x min, {rhythm}, TA {bp.strip()} mmHg, E {resp:g} x min, {breath}, T {temp:.1f}°C.")
        tubes = []
        if cath_look != "(no catheter)":
            tubes.append("Pa urīnpūšļa ilgkatetru " + cath_look + " urīns"
                         + (f" ap {cath_ml:g} ml/dnn" if cath_ml is not None else ""))
        if drain_look != "(no drain)":
            tubes.append("Pa drenu " + drain_look + " izdalījumi" + (f" ap {drain_ml:g} ml/dnn" if drain_ml is not None else "")
                         + (", drenu evakuē" if drain_out else ""))
        for extra in [abdomen, voiding] + tubes + [DRESSING_TEXT if dressing else "", notes]:
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
        st.text_area("Decursus text", text, height=900, key=f"dec-out-{hash(text)}")
        if text:
            def make():
                doc = _docx_new()
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


def render_operation(sex):
    woman = sex == "Woman"
    pt_dat = "pacientei" if woman else "pacientam"
    pt_acc = "Pacienti" if woman else "Pacientu"
    st.subheader("Operācijas apraksts")
    op = st.selectbox("Operation", ["TURP", "TURU", "Circumcision", "Cystoscopy, stent removal", "URS / FURS",
                                    "Epididymal tumour extirpation", "Hydrocele plasty",
                                    "Laparoscopic radical prostatectomy", "Laparoscopic nephrectomy",
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
        elif op == "Cystoscopy, stent removal":
            side = c1.radio("Puse", ["Kreisā", "Labā"], horizontal=True, key="op-cs-side")
            findings = st.text_input("Atrade", "uretra normāla platuma, urīnpūšļa gļotāda nav izmainīta, urēteru ostijas "
                                     "normālas, simetriskas", key="op-cs-find")
            removal = c2.checkbox("Stenta evakuācija", True, key="op-cs-remove")
            ostium = "kreiso" if side == "Kreisā" else "labo"
            name = ("Cistoskopija, " + ("kr." if side == "Kreisā" else "lb.") + " puses pielovezikāla stenta evakuācija"
                    if removal else "Cistoskopija")
            text = [f"Pēc operācijas lauka tualetes, {pt_dat} atrodoties dorsālas guļā tiek ievadīts cistoskops, "
                    f"{findings.strip().rstrip('.')}."]
            if removal:
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
            side = c1.radio("Puse", ["Kreisā", "Labā"], horizontal=True, key="op-nc-side")
            ch = c2.selectbox("Nefrostoma (Ch)", ["8", "10", "12", "14"], index=1, key="op-nc-ch")
            position = c3.selectbox("Novietojums", ["adekvāts", "neadekvāts – koriģēts"], key="op-nc-pos")
            name = f"{'Kreisās' if side == 'Kreisā' else 'Labās'} nefrostomas nomaiņa USS kontrolē"
            text = ["Pēc operācijas lauka apdares caur nefrostomu ievada stīgu. Uz stīgas evakuē nefrostomu. Uz stīgas "
                    f"ievada Ch {ch} nefrostomas katetru. Evakuē stīgu. USS kontrolē konstatē nefrostomas novietojumu. "
                    f"Nefrostomas novietojums {position}. Fiksē nefrostomas katetru."]
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
        elif op == "URS / FURS":
            side = c1.radio("Puse", ["Kreisā", "Labā"], horizontal=True, key="op-u-side")
            stent_out = c2.checkbox("Stenta evakuācija sākumā", True, key="op-u-stentout")
            ureter_stone = c3.checkbox("Konkrements urīnvadā", key="op-u-ustone")
            u_laser = c1.checkbox("Urīnvadā: lāzera litotripsija", key="op-u-ulaser") if ureter_stone else False
            u_extract = c2.checkbox("Urīnvadā: šķembu ekstrakcija", True, key="op-u-uext") if ureter_stone else False
            flexible = c3.checkbox("Fleksibla ureterorenoskopija (nierē)", True, key="op-u-flex")
            if flexible:
                place = c1.selectbox("Konkrements(-i) nierē", ["Nieres bļodiņā vizualizē konkrementu",
                                     "Nieres bļodiņā vizualizē vairākus konkrementus",
                                     "Nieres kausiņos vizualizē multiplus konkrementus"], key="op-u-place")
                k_laser = c2.checkbox("Nierē: lāzera litotripsija", True, key="op-u-klaser")
                k_extract = c3.checkbox("Nierē: ekstrakcija", True, key="op-u-kext")
            else:
                place, k_laser, k_extract = "", False, False
            tool = c1.selectbox("Ekstrakcijas instruments", ["cilpu", "Dormia grozu"], key="op-u-tool") \
                if (u_extract or k_extract) else "cilpu"
            difficulty = st.text_input("Papildus atrade / grūtības (optional)",
                                       placeholder="e.g. Kontaktsangvinācijas dēļ ir apgrūtināta vizualizācija.", key="op-u-diff")
            stent_in = c2.checkbox("Stenta reintubācija beigās", key="op-u-stentin")
            contrast = c3.checkbox("RTG kontrole ar kontrastvielu", True, key="op-u-contrast") if stent_in else False
            ostium = "kreiso" if side == "Kreisā" else "labo"
            short = "Kr." if side == "Kreisā" else "Lb."
            tool_gen = "cilpas" if tool == "cilpu" else "groza"

            name_parts = []
            if stent_out:
                name_parts.append(f"{short} puses pielovezikāla stenta evakuācija")
            main = "fleksibla ureterorenoskopija" if flexible else "ureterorenoskopija"
            withs = []
            if u_laser or k_laser:
                withs.append("lāzera litotripsiju (LLT)")
            if u_extract or k_extract:
                withs.append(f"{tool_gen} litoekstrakciju")
            name_parts.append(main + (" ar " + lv_list(withs) if withs else ""))
            if stent_in:
                name_parts.append("pielovezikāla stenta reintubācija")
            name = "; ".join(name_parts)
            name = name[0].upper() + name[1:]

            text = [f"Pēc operācijas lauka tualetes, {pt_dat} atrodoties dorsālas guļā urīnpūslī ievadīts cistoskops. "
                    "Cistoskopa ievadīšana brīva, ostijas vizualizē. Gļotāda neizmainīta. Ostijas brīvas."]
            if stent_out:
                text.append(f"Vizualizē {ostium} ostiju, stenta galu. Stenta galu satver ar standziņām, evakuē stentu.")
            if ureter_stone:
                ureter = [f"Ievada semirigīdo ureterorenoskopu. Kanulē {ostium} ostiju, ievada semirigīdo instrumentu "
                          "ureterī. Vizualizē konkrementu."]
                if u_laser:
                    ureter.append("Ievieto lāzera stīgu. Vizuālā kontrolē veic konkrementa skaldīšanu. Evakuē lāzera stīgu.")
                if u_extract:
                    ureter.append(f"Ievieto {tool}. Vizuālā kontrolē veic konkrementa šķembu ekstrakciju. "
                                  f"Evakuē {tool}.")
                ureter.append("Aiziet līdz nierei, vairāk konkrementu pa urētera gaitu nevizualizē.")
                text.append(" ".join(ureter))
            else:
                text.append(f"Atkārtoti ievada instrumentu. Kanulē {ostium} ostiju, ievada semirigīdo instrumentu "
                            "ureterī līdz nierei.")
            if flexible:
                kidney = ["Ievada stīgu nierē, evakuē semirigīdo instrumentu un uz stīgas ievada vadītāju. Ievada "
                          "fleksiblo endoskopu. Eksplorē nieri.", place + "."]
                if k_laser:
                    kidney.append("Ievada lāzera stīgu. Vizuālā kontrolē veic konkrementa skaldīšanu ar lāzeri. "
                                  "Evakuē lāzera stīgu.")
                if k_extract:
                    kidney.append(f"Ievada {tool}. Vizuālā kontrolē veic konkrementa ekstrakciju ar {tool}. "
                                  f"Evakuē {tool}.")
                text.append(" ".join(kidney))
            if difficulty.strip():
                text.append(difficulty.strip().rstrip(".") + ".")
            if stent_in:
                end = (f"Ievada semirigīdo ureterorenoskopu, vizualizē {ostium} ostiju. Uz stīgas ievada stentu vizuālā "
                       "kontrolē.")
                if contrast:
                    end += " RTG kontrolē ievada kontrastvielu. Ekstravazāciju nekonstatē."
                text.append(end + " Evakuē instrumentu, darba un drošības stīgu.")
            else:
                text.append("Evakuē instrumentu.")
            text.append(f"Pēc operācijas vispārēji apmierinošā stāvoklī {pt_acc.lower()} pārved uz nodaļu.")
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
        elif op in ("Laparoscopic nephrectomy", "Laparoscopic pyeloplasty"):
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
    output = (f"Operācija: {name}\n\n" if name else "") + "Operācijas apraksts:\n\n" + "\n".join(text)
    header = render_doc_header("op")
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
st.caption(f"Build: 2026-10-21-word · Running file: {Path(__file__).resolve()}")

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