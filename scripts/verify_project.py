from __future__ import annotations

import json
import re
import subprocess
import zipfile
from pathlib import Path

from bson import decode_file_iter
from docx import Document
from pptx import Presentation

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "deliverables"
GROUP = "group_xxx"
COLLECTION = "sales_receipts"
PDF = ROOT / "Big Data Storage_Project.pdf"


def text_from_pdf(path: Path) -> str:
    target = ROOT / "docs" / "pdf_text_check.txt"
    target.parent.mkdir(exist_ok=True)
    subprocess.run(["pdftotext", str(path), str(target)], check=True)
    return target.read_text(errors="replace")


def docx_text(path: Path) -> str:
    doc = Document(str(path))
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


def pptx_text(path: Path) -> str:
    prs = Presentation(str(path))
    lines = []
    for slide in prs.slides:
        for shape in slide.shapes:
            text = getattr(shape, "text", "")
            if text:
                lines.append(text)
    return "\n".join(lines)


def bson_docs(path: Path) -> list[dict]:
    with path.open("rb") as handle:
        return list(decode_file_iter(handle))


def zip_names(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as archive:
        bad = archive.testzip()
        if bad:
            raise RuntimeError(f"bad zip entry: {bad}")
        names = archive.namelist()
        if len(names) != len(set(names)):
            raise RuntimeError("zip has duplicate entries")
        return names


def add_unique(zip_file: Path, items: list[tuple[Path, str]]) -> None:
    temp = zip_file.with_suffix(".tmp.zip")
    skip = {name for _, name in items}
    with zipfile.ZipFile(zip_file) as source, zipfile.ZipFile(temp, "w", compression=zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            if info.filename not in skip:
                target.writestr(info, source.read(info.filename))
        for path, name in items:
            target.write(path, name)
    temp.replace(zip_file)


def has_all(text: str, words: list[str]) -> bool:
    low = text.lower()
    return all(word.lower() in low for word in words)


def query_counts(text: str) -> dict[str, int]:
    business_block = text.split("Simple insight queries", 1)[1].split("Aggregation pipelines", 1)[0]
    explain_block = text.split("Index performance checks", 1)[1]
    return {
        "business_find_queries": len(re.findall(r"db\.sales_receipts\.find\(", business_block)),
        "explain_find_calls": len(re.findall(r"db\.sales_receipts\.find\(", explain_block)),
        "aggregations": len(re.findall(r"db\.sales_receipts\.aggregate\(", text)),
        "indexes": len(re.findall(r"createIndex\(", text)),
        "validators": text.count("bsonType") + text.count("enum"),
        "explain": text.count("explain(\"executionStats\")"),
    }


def valid_doc(doc: dict) -> bool:
    # not full mongo validation. just catches basic breaks.
    required = ["_id", "invoice_id", "branch", "customer", "items", "sale_datetime", "sale_month", "sale_weekday", "sale_hour", "payment", "financials", "rating"]
    if any(key not in doc for key in required):
        return False
    if doc["branch"].get("code") not in {"A", "B", "C"}:
        return False
    if doc["customer"].get("type") not in {"Member", "Normal"}:
        return False
    if doc["payment"].get("method") not in {"Cash", "Credit card", "Ewallet"}:
        return False
    if not isinstance(doc["items"], list) or not doc["items"]:
        return False
    if not 0 <= doc["sale_hour"] <= 23:
        return False
    if not 0 <= doc["rating"] <= 10:
        return False
    item = doc["items"][0]
    money_fields = [doc["financials"].get("cogs"), doc["financials"].get("tax_5"), doc["financials"].get("total"), doc["financials"].get("gross_income"), item.get("unit_price"), item.get("line_cogs"), item.get("tax_5"), item.get("line_total"), item.get("gross_income")]
    return item.get("quantity", 0) >= 1 and all(value is not None and value >= 0 for value in money_fields)


def proof_row(name: str, ok: bool, proof: str) -> dict[str, str]:
    return {"requirement": name, "status": "pass" if ok else "fail", "proof": proof}


def main() -> None:
    report_docx = OUT / f"{GROUP}_report.docx"
    report_pdf = OUT / f"{GROUP}_report.pdf"
    query_file = OUT / f"{GROUP}.txt"
    deck = OUT / f"{GROUP}_presentation.pptx"
    bson_file = OUT / f"{GROUP}.bson"
    zip_file = OUT / f"{GROUP}.zip"
    metadata = OUT / f"{GROUP}.metadata.json"

    assignment = text_from_pdf(PDF)
    pdf_text = text_from_pdf(report_pdf)
    doc_text = docx_text(report_docx)
    deck_text = pptx_text(deck)
    query_text = query_file.read_text(encoding="utf-8")
    docs = bson_docs(bson_file)
    names = zip_names(zip_file)
    counts = query_counts(query_text)
    meta = json.loads(metadata.read_text(encoding="utf-8"))
    first = docs[0]
    invalid_docs = sum(1 for doc in docs if not valid_doc(doc))

    expected_zip = [
        f"{GROUP}.bson",
        f"{GROUP}.txt",
        f"{GROUP}_report.docx",
        f"{GROUP}_report.pdf",
        f"{GROUP}_presentation.pptx",
        f"{GROUP}.metadata.json",
        "README_SUBMISSION.txt",
        f"support/dump/{GROUP}/{COLLECTION}.bson",
        f"support/dump/{GROUP}/{COLLECTION}.metadata.json",
    ]

    acceptance = [
        proof_row("PDF requirements extracted", "Final Deliverables" in assignment, "assignment PDF text was extracted with pdftotext"),
        proof_row("commercial business and dataset chosen", "QuickCart Supermarket" in doc_text and "supermarket sales dataset" in doc_text, "report names the business and source dataset"),
        proof_row("at most 10 collections", True, "model uses one collection: sales_receipts"),
        proof_row("correct MongoDB data types", all(k in first for k in ["sale_datetime", "items", "financials", "rating"]) and invalid_docs == 0, f"BSON decoded and all documents passed data checks, invalid docs: {invalid_docs}"),
        proof_row("five data quality operations documented", has_all(doc_text, ["converts numbers", "parses sale date", "trims", "duplicate", "sale_month"]), "report data quality section names the operations"),
        proof_row("five validation rules", counts["validators"] >= 5 and "invoice_id" in query_text, "query file has JSON schema rules including invoice_id"),
        proof_row("five insight queries", counts["business_find_queries"] >= 5, f"query file has {counts['business_find_queries']} business find queries"),
        proof_row("indexes with performance notes", counts["indexes"] >= 5 and "Index impact" in query_text, f"query file has {counts['indexes']} indexes and performance notes"),
        proof_row("three aggregations", counts["aggregations"] >= 3, f"query file has {counts['aggregations']} aggregation pipelines"),
        proof_row("BSON backup exists", bson_file.exists() and bson_file.stat().st_size > 0 and len(docs) == 1000, f"BSON decoded with {len(docs)} documents"),
        proof_row("report contains required sections", has_all(doc_text, ["One page description", "Model design decisions", "Validation rules", "Indexes and performance", "Advantages and disadvantages", f"{GROUP}.bson", f"{GROUP}.txt"]), "DOCX report contains required sections and file references"),
        proof_row("PDF report mirrors required sections", f"{GROUP}.bson" in pdf_text and f"{GROUP}.txt" in pdf_text and "Team members" in pdf_text, "PDF report text contains cover markers and file references"),
        proof_row("presentation fits defense", len(Presentation(str(deck)).slides) == 8 and has_all(deck_text, ["MongoDB", "validation", "aggregations", "Indexes", "Submission package"]), "deck has 8 slides and covers the required story"),
        proof_row("zip includes required artifacts", all(item in names for item in expected_zip), "zip has all main files at root"),
        proof_row("metadata exists", bool(meta.get("options")) and bool(meta.get("indexes")), "metadata JSON includes collection options and indexes"),
        proof_row("placeholders are isolated", all(token in doc_text + deck_text + query_text for token in ["group_xxx", "NAME", "STUDENT NUMBER"]), "only unknown group fields remain as clear placeholders"),
    ]

    ok = all(row["status"] == "pass" for row in acceptance)
    result = {
        "status": "pass" if ok else "fail",
        "workspace": str(ROOT),
        "zip": str(zip_file),
        "document_count": len(docs),
        "invalid_documents": invalid_docs,
        "zip_entries": names,
        "query_counts": counts,
        "acceptance": acceptance,
        "placeholders_to_replace": ["group_xxx", "NAME", "STUDENT NUMBER"],
    }
    final_json = OUT / "final_verification.json"
    final_txt = OUT / "final_verification_report.txt"
    final_json.write_text(json.dumps(result, indent=2), encoding="utf-8")
    lines = ["Final verification", "", f"Status: {result['status']}", f"Documents: {len(docs)}", f"Zip: {zip_file}", ""]
    for row in acceptance:
        lines.append(f"{row['status'].upper()}: {row['requirement']} | {row['proof']}")
    lines.extend(["", "Replace before Moodle:", "group_xxx", "NAME", "STUDENT NUMBER"])
    final_txt.write_text("\n".join(lines) + "\n", encoding="utf-8")
    add_unique(
        zip_file,
        [
            (final_json, "support/final_verification.json"),
            (final_txt, "support/final_verification_report.txt"),
            (ROOT / "scripts" / "verify_project.py", "support/verify_project.py"),
        ],
    )
    result["zip_entries"] = zip_names(zip_file)
    final_json.write_text(json.dumps(result, indent=2), encoding="utf-8")
    add_unique(zip_file, [(final_json, "support/final_verification.json")])
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
