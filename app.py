import hashlib
import json
from pathlib import Path

import pandas as pd
import streamlit as st
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
    "OFFICE": "Ambulatory office",
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
    "Lab follow-up": {
        "type": "multi", "optional": True,
        "options": ["Urinalysis and urine culture with susceptibility",
                    "Complete blood count, CRO, creatinine and eGFR",
                    "Kidney and bladder ultrasound",
                    "Other tests as advised by the treating urologist"],
        "default_selected": ["Urinalysis and urine culture with susceptibility",
                             "Complete blood count, CRO, creatinine and eGFR",
                             "Other tests as advised by the treating urologist"],
    },
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
    "Activity restriction": {"type": "single", "options": ["1 week", "2 weeks", "3 weeks", "4 weeks", "6 weeks"]},
    "Urology visit": {"type": "single", "options": ["2 weeks", "4 weeks", "6 weeks", "3 months"]},
    "Tests before the visit": {
        "type": "multi", "optional": True,
        "options": ["Urinalysis", "Urine cytology", "Creatinine and eGFR",
                    "Full blood panel (previous list)", "Kidney and bladder ultrasound"],
    },
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
    "First PSA after surgery": {
        "type": "single", "options": ["6 weeks", "3 months"], "default_index": 1,
        "show_if": {"Operation": "Radical prostatectomy"},
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
    "Urethrogram before removal": {
        "type": "toggle", "default": True,
        "show_if": {"Stricture operation": ["Urethroplasty – excision and primary anastomosis",
                                            "Urethroplasty – buccal mucosa graft"],
                    "Catheter at discharge (N35)": "In place"},
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
    "Failed urethral catheterisation": {"type": "toggle", "default": True},
    "Ultrasound guided": {"type": "toggle", "default": True},
    "Creatinine raised (N35)": {"type": "toggle"},
    "Investigations (N35)": {
        "type": "text", "optional": True,
        "placeholder": "optional; empty = 'veikti ambulatorā kārtā, pieļauj operāciju'",
    },
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
         ["Urology follow-up", "Urology follow-up date", "Fluid intake", "Lab follow-up", "Warning signs",
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
         ["Discharge paper", "Activity restriction", "Urology visit", "Tests before the visit",
          "Oncology council", "Oncology council date"]),
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
          "Intraoperative complications"],
         {"Operation": "Radical prostatectomy"}),
        ("3. TURP",
         ["TURP anesthesia", "Procedure date", "Intraoperative course",
          "Intraoperative complications", "Bladder irrigation after surgery"],
         {"Operation": "TURP"}),
        ("4. Course and histology",
         ["Postoperative course", "Postoperative complications", "Catheter removal", "Catheter removed on",
          "Planned catheter removal (optional)", "Voiding after catheter removal", "Histology",
          "Pathological T stage", "Pathological N stage", "Surgical margin", "Final Gleason score",
          "Prostate histology details"],
         {"_department": "4"}),
        ("5. Recommendations",
         ["Discharge paper", "Activity restriction", "Urology visit", "First PSA after surgery",
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
          "Urethrogram before removal", "Voiding after removal (N35)", "Self-dilation taught"],
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
LAB_FOLLOWUP_TEXT = {
    "Urinalysis and urine culture with susceptibility":
        "Urīna analīze un urīna uzsējums ar antibakteriālās jutības noteikšanu.",
    "Complete blood count, CRO, creatinine and eGFR":
        "Pilna asinsaina, CRO, kreatinīns un GFĀ.",
    "Kidney and bladder ultrasound": "USS nierēm un urīnpūslim.",
    "Other tests as advised by the treating urologist":
        "Citi izmeklējumi pēc ārstējošā urologa norādījumiem.",
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
C67_TESTS_TEXT = {  # "Pirms vizītes veikt: ..."
    "Urinalysis": "Urīna analīzi.",
    "Urine cytology": "Urīna citoloģisko izmeklēšanu.",
    "Creatinine and eGFR": "Kreatinīnu un GFĀ.",
    "Full blood panel (previous list)": "Asins analīzes: pilna asinsaina, CRO, LDH, sārmainā fosfatāze, "
                                        "kreatinīns, urīnskābe, Ca, P, Na, K, glikoze, ASAT, ALAT, kopējais "
                                        "bilirubīns un 25-OH-vitamīns D (D3+D2).",
    "Kidney and bladder ultrasound": "Nieru un urīnpūšļa ultrasonogrāfiju.",
}
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
              "6 weeks": "6 nedēļas", "3 months": "3 mēnešus"}

# ----- Wording for prostate cancer (C61) -----
RP_TEXT = {  # (as the operation done, in "St. pēc ...")
    "Open retropubic": ("atvērta retropubiska radikāla prostatektomija", "atvērtas retropubiskas radikālas prostatektomijas"),
    "Laparoscopic": ("laparoskopiska radikāla prostatektomija", "laparoskopiskas radikālas prostatektomijas"),
    "Robot-assisted": ("robotasistēta radikāla prostatektomija", "robotasistētas radikālas prostatektomijas"),
}
LND_TEXT = (" ar iegurņa limfadenektomiju", " un iegurņa limfadenektomijas")
ANESTHESIA_TEXT = {"Spinal": "spinālā anestēzijā", "General": "vispārējā anestēzijā"}
PSA_TIME_TEXT = {"6 weeks": "6 nedēļas", "3 months": "3 mēnešus"}
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
FIRST_CT_TEXT = {"Low": "6 mēnešus", "Intermediate": "6 mēnešus", "High": "3 mēnešus"}  # EAU Table 9.1
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



# Text templates. A template is used when the diagnosis and department
# match AND every modifier listed under "when" has that value.
# Each template has: clinical_diagnosis (optional), epicrisis, recommendations.
TEMPLATES = [
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

    tests = [LAB_FOLLOWUP_TEXT[t] for t in LAB_FOLLOWUP_TEXT if t in (v.get("Lab follow-up") or [])]
    if tests:
        lines.append(numbered("Kontroles izmeklējumi atbilstoši klīniskajai situācijai:"))
        lines.extend("    o " + t for t in tests)
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
    lines += [
        numbered("Ambulatori urologa konsultācija pēc mēneša ar iepriekš veiktiem izmeklējumiem: "
                 "PAA, bioķīmija, urīna analīze."),
        numbered("Atkārtoti veikt USS vēdera dobumam pēc mēneša"),
        numbered("Stāvokļa pasliktināšanās gadījumā vērsties pēc palīdzības NMC."),
    ]
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
    visit = f"Vizīte pie urologa ambulatori: {WEEKS_TEXT[v['Urology visit']]} pēc operācijas."
    tests = [C67_TESTS_TEXT[t] for t in C67_TESTS_TEXT if t in (v.get("Tests before the visit") or [])]
    if tests:
        recs.append(visit + " Pirms vizītes veikt:")
        recs += [f"{number}. {test}" for number, test in enumerate(tests, start=1)]
    else:
        recs.append(visit)
    recs.append("Ģimenes ārsta nosūtījums plānotai cistoskopijai.")
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


def c61_common_recs(f, v, blocks, extra_regime):
    """Recommendation lines shared by prostatectomy and TURP."""
    recs = []
    if v["Discharge paper"] == "Temporary":
        recs.append(TEMPORARY_PAPER_TEXT)
    recs += [
        "Atrasties ģimenes ārsta un ārstējošā urologa uzraudzībā.",
        "Diēta: vispusīga. Lietot šķidrumu aptuveni 2,0 L diennaktī, ja nav medicīnisku kontrindikāciju.",
        f"Režīms: saudzējošs {WEEKS_TEXT[v['Activity restriction']]} no operācijas dienas. Atturēties no "
        "pārmērīgas fiziskas slodzes, smagumu celšanas, pirts un baseina apmeklējuma.",
    ] + extra_regime
    if v["Catheter removal"] == "Discharged with catheter":
        recs += CATHETER_CARE_LINES
        if v.get("Planned catheter removal (optional)"):
            recs.append(f"Katetra evakuācija plānota {fmt_date(v['Planned catheter removal (optional)'])}.")
    else:
        recs.append("Regulāri, ik pēc 1–2 stundām, iztukšot urīnpūsli.")
    if blocks:
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
    epicrisis = "\n".join([
        f"{f['Pacients']} {f['stacionets']} P. Stradiņa KUS 4. uroloģijas nodaļā {admission} operatīvai "
        f"ārstēšanai sakarā ar verificētu prostatas vēzi{c61_basis(v)}. {f['Pacients']} tiek {f['gatavots']} "
        "operatīvai terapijai.",
        f"{day} vispārējā anestēzijā veikta {operation}. Operācija noritēja {intraop}.",
        f"Pēcoperācijas periods {postop}. {c61_catheter_sentence(f, v)}",
        f"{c61_histology(v)} Uzlabojoties vispārējam stāvoklim, {f['pacients']} {f['izrakstits']} turpmākai "
        "aprūpei ģimenes ārsta uzraudzībā. Sniegtas rekomendācijas.",
    ])
    recs = c61_common_recs(f, v, list(blocks), [])
    recs.insert(recs.index(next(r for r in recs if r.startswith(("Urīnpūšļa katetra maiņa", "Regulāri")))),
                "Pēc katetra evakuācijas regulāri veikt iegurņa pamatnes muskulatūras vingrinājumus urīna "
                "noturēšanas atjaunošanai.")
    recs.append(f"PSA kontrole {PSA_TIME_TEXT[v['First PSA after surgery']]} pēc operācijas; turpmāk "
                "saskaņā ar ārstējošā urologa norādījumiem.")
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
    recs.append(visit + ". Pirms vizītes noteikt kreatinīnu un GFĀ.")
    if benign:
        recs.append("Turpmākā novērošana pēc ārstējošā urologa norādījumiem.")
    elif v["Histology"] == "Known":
        recs.append(f"Kontroles CT krūškurvim un vēdera dobumam {FIRST_CT_TEXT[v['Follow-up risk group']]} pēc "
                    "operācijas; turpmākā novērošana atbilstoši riska grupai pēc ārstējošā urologa norādījumiem.")
    else:
        recs.append("Kontroles izmeklējumu (CT) plāns tiks noteikts pēc histoloģiskās atbildes saņemšanas.")
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
    recs = (
        "Atrasties ģimenes ārsta uzraudzībā.\n"
        "• Diēta – vispusīga, lietot šķidrumu 2.0 – 2.5 L/dnn, atturēties no saldinātu, "
        "kofeīnu vai alkoholu saturošu dzērienu lietošanas.\n"
        "• Režīms – saudzējošs, izvairīties no pārmērīgas fiziskas slodzes, pirts vai baseina apmeklējuma.\n"
        "• Regulāri, vismaz reizi 1 – 2 stundās iztukšot urīnpūsli.\n"
        "{blocks}\n"
        "• Kontrole pie urologa ambulatori pēc 3 mēnešiem ar sekojošiem izmeklējumiem:\n"
        "    o Urīna analīze;\n"
        "    o Urīna uzsējums ar a/b jūtību;\n"
        "    o Asins aina, CRO, kreatinīns, GFĀ;\n"
        "    o USS nierēm un urīnpūslim.\n"
        "• Stāvokļa pasliktināšanās gadījumā griezties NMC."
    )
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
        numbered("Stāvokļa pasliktināšanās gadījumā vērsties pēc palīdzības NMC."),
    ]
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
            "• Stāvokļa pasliktināšanās gadījumā griezties NMC.",
        ]
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
        numbered("Stāvokļa pasliktināšanās gadījumā vērsties pēc palīdzības NMC."),
    ]
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
        if op in URETHROPLASTY and v.get("Urethrogram before removal"):
            recs.append(f"• Urīna ilgkatetra evakuācija plānota {when}, pirms tam veicot pericatetrālu retrogrādu "
                        "uretrogrāfiju (vai mikcijas cistouretrogrāfiju), lai izslēgtu kontrasta ekstravazāciju.")
        else:
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
        recs.append("• Kontrole pie urologa ambulatori pēc 3 mēnešiem ar urīna plūsmas mērījumu (uroflowmetriju), "
                    "simptomu un seksuālās funkcijas novērtējumu un uretrocistoskopiju; turpmāk pēc 12 un 24 mēnešiem.")
    else:
        recs.append("• Kontrole pie urologa ambulatori pēc 3 mēnešiem ar urīna plūsmas mērījumu (uroflowmetriju) un "
                    "simptomu novērtējumu.")
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
    if op in URETHROPLASTY and v.get("Catheter at discharge (N35)") == "In place" and not v.get("Urethrogram before removal"):
        cautions.append("EAU (strong recommendation): perform a urethrogram before catheter removal after urethroplasty.")
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
                 "(retrogrāda uretrogrāfija, uretrocistoskopija) un plānotu definitīvo ārstēšanu."),
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


# ---------------------------------------------------------------
# 3. PAGE
# ---------------------------------------------------------------

st.set_page_config(page_title="Discharge helper", layout="wide")
st.markdown("<h1 style='text-align: center;'>Discharge helper</h1>", unsafe_allow_html=True)
st.caption(f"Build: 2026-10-18-neph · Running file: {Path(__file__).resolve()}")

# Make the two sex buttons tall rectangles
st.markdown(
    "<style>"
    ".st-key-sex_woman button, .st-key-sex_man button {height: 3.4rem; font-size: 1.25rem;}"
    "</style>",
    unsafe_allow_html=True,
)

# Bar 1: where the patient is discharged from
department = st.segmented_control(
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