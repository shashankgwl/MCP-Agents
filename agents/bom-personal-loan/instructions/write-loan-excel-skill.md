---
name: write-loan-excel-skill
description: Generate and validate one Excel workbook from one application's JSON using the embedded clean template.
---
## Responsibility and inputs
Receive extraction_json, manifest_path and application_id from the orchestrator.
Do not invoke another skill or the workflow. Do not read another application's
JSON or invent missing fields. Only the orchestrator sequences the steps.

## Deterministic Excel creation
Run the supplied Python implementation unchanged. It embeds the approved blank
four-sheet template, verifies its checksum, validates application/source identity
and PDF page count, maps fields by exact template labels, clears all editable
cells, preserves formatting/dropdowns, rejects conflicting values and safely
stores literal text (never formulas) for identifiers and extracted strings.
Use openpyxl and pypdf from the sandbox toolkit. If a required module is missing,
install only that dependency if the runtime permits it; otherwise report the
specific missing dependency. Do not replace the implementation with guessed
cell addresses, CSV output or model-generated workbook/base64 content.

Copy the Python block verbatim to bom_writer.py (or extract it programmatically
from this skill's SKILL.md when available). Execute:
python bom_writer.py "<extraction_json>" "<manifest_path>" "<application_id>"
Invoke using Python arguments or a subprocess argument list with shell=False.

Keep all four tabs in template order: personal details; employment details;
DETAILS OF EXISTING LOAN; CHECKLIST FOR PERSONAL LOAN.
Missing application values and unused loan slots remain blank. Up to two
existing loans per party are supported; overflow is a blocking issue.
Checklist points 3-5 stay blank; supporting documents are out of scope.
Retain an explicit application answer for points 1-2 only with source evidence.
Do not alter label, header or guidance cells or add unrequested financial sheets.
Source fields without a template destination must be disclosed in issues.

## File handoff
The helper reopens the saved XLSX, verifies written values and dropdowns, rejects
Excel errors/formulas, and emits workbook_path and upload_payload_path plus
per-application status and issues. It encodes the real XLSX bytes exactly once
into the upload JSON, whose fields are:
  text: the unchanged manifest batch_folder
  file: {name: unique_workbook_filename, contentBytes: complete_raw_base64}
Return these file paths and status to the orchestrator. Never print/retype the
base64 payload or pass the upload JSON itself as Excel file content.
The orchestrator uses the supported file/tool bridge to call storage.
If that bridge cannot transport the actual bytes, report an upload limitation
instead of substituting a sandbox path or claiming the workbook was uploaded.

```python
import argparse
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.comments import Comment


PERSONAL = "personal details"
EMPLOYMENT = "employment details"
LOANS = "DETAILS OF EXISTING LOAN"
CHECKLIST = "CHECKLIST FOR PERSONAL LOAN"
SHEETS = [PERSONAL, EMPLOYMENT, LOANS, CHECKLIST]
MAPPINGS = {
    "full_name": (PERSONAL, "Full name"),
    "age_years": (PERSONAL, "Age (years)"),
    "date_of_birth": (PERSONAL, "Date of birth"),
    "father_husband_name": (PERSONAL, "Father's / Husband's name"),
    "current_address": (PERSONAL, "Current residential address"),
    "permanent_address": (PERSONAL, "Permanent residential address"),
    "residential_telephone": (PERSONAL, "Residential telephone number"),
    "mobile_number": (PERSONAL, "Mobile number"),
    "email": (PERSONAL, "Email ID"),
    "aadhar_number": (PERSONAL, "Aadhar number"),
    "pan": (PERSONAL, "PAN allotted by I.T. Department"),
    "passport_number": (PERSONAL, "Passport number"),
    "passport_issue_date": (PERSONAL, "Passport date of issuance"),
    "education": (PERSONAL, "Educational qualification"),
    "salary_account": (PERSONAL, "Salary account number"),
    "marital_status": (PERSONAL, "Marital status"),
    "dependents": (PERSONAL, "Number of dependents"),
    "housing_status": (PERSONAL, "Housing status"),
    "employer_name": (EMPLOYMENT, "Name of employer"),
    "employer_category": (EMPLOYMENT, "Category of employer"),
    "employer_category_other": (EMPLOYMENT, "Other employer category - specify"),
    "employer_address": (EMPLOYMENT, "Employer address"),
    "employer_telephone": (EMPLOYMENT, "Employer telephone number"),
    "employer_website": (EMPLOYMENT, "Employer website"),
    "employer_email": (EMPLOYMENT, "Employer email ID"),
    "designation": (EMPLOYMENT, "Designation"),
    "years_present_occupation": (EMPLOYMENT, "Years in present occupation"),
    "retirement_date": (EMPLOYMENT, "Date of retirement"),
    "previous_employment": (EMPLOYMENT, "Previous employment details"),
    "length_of_service": (EMPLOYMENT, "Length of service"),
    "loan_bank_branch": (LOANS, "Name of bank / branch"),
    "loan_type": (LOANS, "Type of loan availed"),
    "loan_amount": (LOANS, "Amount of loan availed (INR)"),
    "loan_outstanding": (LOANS, "Present outstanding (INR)"),
    "loan_security": (LOANS, "Security offered (if any)"),
    "loan_repayment": (LOANS, "Repayment plan"),
    "loan_default": (LOANS, "Amount of default, if any (INR)"),
}
CHECKLIST_LABELS = [
    "Application form duly complete and signed", "Two passport size photographs",
    "Photo Identity Proof", "Residence Address Proof", "Income Proof",
]
CHECKLIST_OPTIONS = [
    ["Yes", "No"], ["Yes", "No"],
    ["Passport", "Pan Card", "Voters ID Card", "Driving License", "Aadhar Card"],
    ["Passport", "Ration Card", "Utility Bill: Electricity Bill/ Land Line Bill",
     "Aadhar Card", "Voters ID card", "Driving License"],
    ["Last 3 months salary slip", "Latest Form 16/ ITR for two years"],
]
OPTIONS = {
    "marital_status": ["Single", "Married"],
    "housing_status": ["Own", "Rented", "Quarters provided by employer"],
    "employer_category": [
        "Govt./Semi & Quasi Govt.", "Public Ltd Co.", "Pvt. Ltd Co.", "Multi National Co.",
        "Reputed Institutions", "Proprietorship/Partnership firms", "Others",
    ],
}
NUMERIC = {"age_years", "dependents", "loan_amount", "loan_outstanding", "loan_default"}


def populate(template, data, output):
    import math
    import re

    if not isinstance(data, dict) or data.get("is_personal_loan_application") is not True:
        raise ValueError("A confirmed personal-loan application is required.")
    if not isinstance(data.get("source_file"), str) or not data["source_file"].strip():
        raise ValueError("The PDF source filename is required.")
    for key in ("fields", "checklist", "issues", "blocking_issues"):
        if not isinstance(data.get(key), list):
            raise ValueError(f"{key} must be an array.")
    if not all(isinstance(issue, str) for issue in data["issues"] + data["blocking_issues"]):
        raise ValueError("Issues must contain strings.")
    if data["blocking_issues"]:
        raise ValueError("Blocking extraction issues require manual review.")
    workbook = load_workbook(template)
    if workbook.sheetnames != SHEETS:
        raise ValueError("Unexpected template sheets or order.")
    rows = {}
    for sheet in workbook:
        labels = {}
        for row in range(7, sheet.max_row + 1):
            value = sheet.cell(row, 2 if sheet.title == CHECKLIST else 1).value
            if value:
                if value in labels:
                    raise ValueError(f"Duplicate template label: {value}")
                labels[value] = row
        rows[sheet.title] = labels
    for sheet, label in MAPPINGS.values():
        if label not in rows[sheet]:
            raise ValueError(f"Missing template label: {sheet} / {label}")
    if list(rows[CHECKLIST]) != CHECKLIST_LABELS:
        raise ValueError("Expected the corrected five-row checklist.")
    validations = {sheet.title: str(sheet.data_validations) for sheet in workbook}
    for sheet, label in MAPPINGS.values():
        for column in range(2, 6 if sheet == LOANS else 4):
            cell = workbook[sheet].cell(rows[sheet][label], column)
            cell.value, cell.comment = None, None
    for label in CHECKLIST_LABELS:
        for column in (3, 4, 5):
            cell = workbook[CHECKLIST].cell(rows[CHECKLIST][label], column)
            cell.value, cell.comment = None, None
    writes, seen = [], set()
    for field in data["fields"]:
        if not isinstance(field, dict) or field.get("field_key") not in MAPPINGS:
            raise ValueError("Unknown or malformed extracted field.")
        key, party, index = field["field_key"], field.get("party"), field.get("loan_index")
        value, page = field.get("value"), field.get("source_page")
        sheet, label = MAPPINGS[key]
        if party not in ("borrower", "guarantor") or type(index) is not int:
            raise ValueError("Invalid party or loan index.")
        if index not in ((1, 2) if sheet == LOANS else (0,)):
            raise ValueError("Template supports two loans per party; non-loan fields use index 0.")
        if not isinstance(value, str) or type(page) is not int or page < 0:
            raise ValueError("Field value must be text; source_page must be a non-negative integer.")
        if len(value) > 32767:
            raise ValueError("Extracted value exceeds the Excel cell text limit.")
        identity = (key, party, index)
        if identity in seen:
            raise ValueError(f"Duplicate or conflicting field: {identity}")
        seen.add(identity)
        if not value.strip():
            continue
        if key in OPTIONS and value not in OPTIONS[key]:
            raise ValueError(f"Invalid dropdown option: {key}")
        if key in NUMERIC:
            clean = re.sub(r"^(INR|Rs\.?)\s*", "", value, flags=re.I).replace(",", "").strip()
            if not re.fullmatch(r"\d+(?:\.\d+)?", clean):
                raise ValueError(f"Invalid non-negative number: {key}")
            value = float(clean)
            if not math.isfinite(value):
                raise ValueError(f"Non-finite number: {key}")
            if key in ("age_years", "dependents"):
                if not value.is_integer():
                    raise ValueError(f"Expected whole number: {key}")
                value = int(value)
        column = ((2 if party == "borrower" else 4) + index - 1
                  if sheet == LOANS else (2 if party == "borrower" else 3))
        writes.append((sheet, rows[sheet][label], column, value,
                       f"Source: {data['source_file']}; PDF page {page if page else 'not supplied'}."
                       + (f" Evidence: {field['evidence']}" if field.get("evidence") else "")))
    application_fields = len(writes)
    if not application_fields:
        raise ValueError("No application fields were extracted; refusing an empty success-shaped workbook.")
    for entry in data["checklist"]:
        if not isinstance(entry, dict):
            raise ValueError("Malformed checklist entry.")
        point, party, value = entry.get("point"), entry.get("party"), entry.get("value")
        evidence = entry.get("evidence")
        if type(point) is not int or not 1 <= point <= 5 or party not in ("borrower", "guarantor"):
            raise ValueError("Invalid checklist point or party.")
        if not isinstance(value, str) or not isinstance(evidence, str):
            raise ValueError("Checklist value and evidence must be strings.")
        identity = ("checklist", point, party)
        if identity in seen:
            raise ValueError("Duplicate checklist selection.")
        seen.add(identity)
        if not value:
            continue
        if value not in CHECKLIST_OPTIONS[point - 1] or not evidence.strip():
            raise ValueError("Checklist choice must be allowed and supported by evidence.")
        writes.append((CHECKLIST, rows[CHECKLIST][CHECKLIST_LABELS[point - 1]],
                       3 if party == "borrower" else 4, value,
                       f"Source: {data['source_file']}. {evidence}"))
    for sheet, row, column, value, source in writes:
        cell = workbook[sheet].cell(row, column)
        if isinstance(value, str):
            cell.value = value
            cell.data_type = "s"
            cell.number_format = "@"
        else:
            cell.value = value
            cell.number_format = "0" if isinstance(value, int) else '#,##0.00;(#,##0.00);"-"'
        cell.comment = Comment(source, "Document evidence")
    note = "; ".join(data["issues"] + ["Application-only extraction; supporting documents were not reviewed."])
    cell = workbook[CHECKLIST].cell(rows[CHECKLIST][CHECKLIST_LABELS[4]], 5)
    cell.value, cell.data_type = note, "s"
    output = Path(output)
    if output.exists():
        raise FileExistsError(f"Refusing to overwrite: {output}")
    workbook.save(output)
    check = load_workbook(output)
    if check.sheetnames != SHEETS:
        raise ValueError("Saved workbook sheet validation failed.")
    for sheet in check:
        if str(sheet.data_validations) != validations[sheet.title]:
            raise ValueError("Template data validation changed.")
        for row in sheet:
            for cell in row:
                if cell.data_type in ("e", "f"):
                    raise ValueError("Unexpected Excel error or formula.")
    for sheet, row, column, value, _ in writes:
        if check[sheet].cell(row, column).value != value:
            raise ValueError("Saved workbook does not match extracted fields.")
    return {"status": "needs_review" if data["issues"] else "populated",
            "application_fields_written": application_fields, "cells_written": len(writes)}



import argparse
import base64
import hashlib
import json
from pathlib import Path
import re
import zlib

from pypdf import PdfReader



def create_workbook(template, extraction_json, manifest_path, application_id):
    manifest_path = Path(manifest_path)
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    application = next(
        (item for item in manifest["applications"] if item["application_id"] == application_id),
        None,
    )
    if application is None:
        raise ValueError("Application ID is not in this ZIP manifest.")
    source = Path(application["pdf_path"])
    if hashlib.sha256(source.read_bytes()).hexdigest() != application["sha256"]:
        raise ValueError("The PDF changed after ZIP extraction.")
    data = json.loads(Path(extraction_json).read_text(encoding="utf-8-sig"))
    if data.get("schema_version") != "1.0":
        raise ValueError("Unsupported extraction schema version.")
    if data.get("source_file") != application["source_file"]:
        raise ValueError("Extraction belongs to a different PDF.")
    if data.get("application_id") != application_id:
        raise ValueError("Extraction belongs to a different application.")
    if data.get("source_sha256") != application["sha256"]:
        raise ValueError("Extraction does not match the manifest PDF hash.")
    page_count = data.get("page_count")
    if type(page_count) is not int or page_count < 1:
        raise ValueError("A positive PDF page count is required.")
    reader = PdfReader(source)
    if reader.is_encrypted or len(reader.pages) != page_count:
        raise ValueError("Extraction page count does not match an accessible source PDF.")
    for field in data.get("fields", []):
        page = field.get("source_page")
        if type(page) is not int or not 1 <= page <= page_count:
            raise ValueError("Every field needs a valid 1-based source page.")
        if not isinstance(field.get("evidence"), str) or not field["evidence"].strip():
            raise ValueError("Every field needs a quoted source value or visual evidence.")
    for entry in data.get("checklist", []):
        if entry.get("point") not in (1, 2):
            raise ValueError("Supporting-proof checklist points 3-5 are out of scope.")
    name = application["output_name"]
    if not re.fullmatch(r"[A-Za-z0-9_-]+\.xlsx", name):
        raise ValueError("Unsafe output workbook filename.")
    output_folder = manifest_path.parent / "workbooks"
    output_folder.mkdir(exist_ok=True)
    output = output_folder / name
    result = populate(template, data, output)
    raw = output.read_bytes()
    handoff = {
        "text": manifest["batch_folder"],
        "file": {"name": output.name, "contentBytes": base64.b64encode(raw).decode("ascii")},
    }
    handoff_path = output.with_suffix(".upload.json")
    with handoff_path.open("x", encoding="utf-8") as handle:
        json.dump(handoff, handle)
    return {
        **result,
        "application_id": application_id,
        "source_file": application["source_file"],
        "workbook_path": str(output),
        "upload_payload_path": str(handoff_path),
        "file_name": output.name,
        "bytes": len(raw),
        "sha256": hashlib.sha256(raw).hexdigest(),
        "issues": data["issues"],
    }




TEMPLATE_SHA256 = "e72ea2ac04909c31fcc224d1bd514c58366a7b4c3523260a1d03a1945b2b9fe9"
TEMPLATE_ZLIB_BASE64 = (
    "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"

)


if __name__ == "__main__":
    import tempfile

    parser = argparse.ArgumentParser()
    parser.add_argument("extraction_json")
    parser.add_argument("manifest_path")
    parser.add_argument("application_id")
    arguments = parser.parse_args()
    template_bytes = zlib.decompress(base64.b64decode(TEMPLATE_ZLIB_BASE64, validate=True))
    if hashlib.sha256(template_bytes).hexdigest() != TEMPLATE_SHA256:
        raise ValueError("Embedded template checksum mismatch.")
    with tempfile.TemporaryDirectory(prefix="bom-template-") as directory:
        template_path = Path(directory) / "template.xlsx"
        template_path.write_bytes(template_bytes)
        result = create_workbook(template_path, arguments.extraction_json,
                                 arguments.manifest_path, arguments.application_id)
    print(json.dumps(result, ensure_ascii=False))

```
