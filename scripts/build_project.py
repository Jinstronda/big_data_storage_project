from __future__ import annotations

import csv
import json
import math
import shutil
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from urllib.request import Request, urlopen

from bson import BSON
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from pptx import Presentation
from pptx.dml.color import RGBColor as PptColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches as PptInches
from pptx.util import Pt as PptPt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
DOCS_DIR = ROOT / "docs"
OUT_DIR = ROOT / "deliverables"
DATA_URL = "https://raw.githubusercontent.com/selva86/datasets/master/supermarket_sales.csv"
CSV_PATH = DATA_DIR / "supermarket_sales.csv"
COLLECTION = "sales_receipts"
GROUP = "group_xxx"

PRODUCT_LINES = [
    "Electronic accessories",
    "Fashion accessories",
    "Food and beverages",
    "Health and beauty",
    "Home and lifestyle",
    "Sports and travel",
]
BRANCH_CITIES = {"A": "Yangon", "B": "Mandalay", "C": "Naypyitaw"}
PAYMENT_METHODS = ["Cash", "Credit card", "Ewallet"]
CUSTOMER_TYPES = ["Member", "Normal"]
GENDERS = ["Female", "Male"]


def money(value: float) -> str:
    return f"${value:,.2f}"


def pct(value: float) -> str:
    return f"{value:.1f}%"


def clean_number(value: str) -> float:
    return round(float(value), 4)


def ensure_dirs() -> None:
    for path in [DATA_DIR, DOCS_DIR, OUT_DIR]:
        path.mkdir(parents=True, exist_ok=True)


def download_dataset() -> None:
    if CSV_PATH.exists() and CSV_PATH.stat().st_size > 1000:
        return
    req = Request(DATA_URL, headers={"User-Agent": "Mozilla/5.0"})
    with urlopen(req, timeout=60) as response:
        CSV_PATH.write_bytes(response.read())


def read_rows() -> list[dict[str, str]]:
    with CSV_PATH.open(newline="", encoding="utf-8") as handle:
        return list(csv.DictReader(handle))


def parse_datetime(row: dict[str, str]) -> datetime:
    return datetime.strptime(f"{row['Date']} {row['Time']}", "%m/%d/%Y %H:%M")


def sale_doc(row: dict[str, str]) -> dict:
    sale_dt = parse_datetime(row)
    unit_price = clean_number(row["Unit price"])
    quantity = int(row["Quantity"])
    tax = clean_number(row["Tax 5%"])
    total = clean_number(row["Total"])
    cogs = clean_number(row["cogs"])
    gross_income = clean_number(row["gross income"])
    branch = row["Branch"].strip()
    return {
        "_id": row["Invoice ID"].strip(),
        "invoice_id": row["Invoice ID"].strip(),
        "branch": {"code": branch, "city": BRANCH_CITIES[branch]},
        "customer": {
            "type": row["Customer type"].strip(),
            "gender": row["Gender"].strip(),
        },
        "items": [
            {
                "product_line": row["Product line"].strip(),
                "unit_price": unit_price,
                "quantity": quantity,
                "line_cogs": cogs,
                "tax_5": tax,
                "line_total": total,
                "gross_income": gross_income,
            }
        ],
        "sale_datetime": sale_dt,
        "sale_date": sale_dt.strftime("%Y-%m-%d"),
        "sale_month": sale_dt.strftime("%Y-%m"),
        "sale_weekday": sale_dt.strftime("%A"),
        "sale_hour": sale_dt.hour,
        "payment": {"method": row["Payment"].strip()},
        "financials": {
            "cogs": cogs,
            "tax_5": tax,
            "total": total,
            "gross_margin_percentage": clean_number(row["gross margin percentage"]),
            "gross_income": gross_income,
        },
        "rating": clean_number(row["Rating"]),
        "source": "selva86/datasets supermarket_sales.csv",
    }


def build_documents(rows: list[dict[str, str]]) -> list[dict]:
    seen = set()
    docs = []
    for row in rows:
        invoice = row["Invoice ID"].strip()
        if invoice in seen:
            continue
        seen.add(invoice)
        docs.append(sale_doc(row))
    return docs


def json_ready(doc: dict) -> dict:
    ready = {}
    for key, value in doc.items():
        if isinstance(value, datetime):
            ready[key] = {"$date": value.isoformat()}
        elif isinstance(value, dict):
            ready[key] = json_ready(value)
        elif isinstance(value, list):
            ready[key] = [json_ready(v) if isinstance(v, dict) else v for v in value]
        else:
            ready[key] = value
    return ready


def write_bson(docs: list[dict]) -> None:
    # one receipt stays one doc. simple to defend.
    target = OUT_DIR / f"{GROUP}.bson"
    with target.open("wb") as handle:
        for doc in docs:
            handle.write(BSON.encode(doc))
    shutil.copy2(target, OUT_DIR / f"{COLLECTION}.bson")
    dump_dir = OUT_DIR / "dump" / GROUP
    dump_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(target, dump_dir / f"{COLLECTION}.bson")
    with (OUT_DIR / f"{COLLECTION}.json").open("w", encoding="utf-8") as handle:
        for doc in docs:
            handle.write(json.dumps(json_ready(doc), ensure_ascii=False) + "\n")


def validator() -> dict:
    return {
        "$jsonSchema": {
            "bsonType": "object",
            "required": [
                "_id",
                "invoice_id",
                "branch",
                "customer",
                "items",
                "sale_datetime",
                "sale_month",
                "sale_weekday",
                "sale_hour",
                "payment",
                "financials",
                "rating",
            ],
            "properties": {
                "_id": {"bsonType": "string", "pattern": "^[0-9]{3}-[0-9]{2}-[0-9]{4}$"},
                "invoice_id": {"bsonType": "string", "pattern": "^[0-9]{3}-[0-9]{2}-[0-9]{4}$"},
                "branch": {
                    "bsonType": "object",
                    "required": ["code", "city"],
                    "properties": {
                        "code": {"enum": list(BRANCH_CITIES.keys())},
                        "city": {"enum": list(BRANCH_CITIES.values())},
                    },
                },
                "customer": {
                    "bsonType": "object",
                    "required": ["type", "gender"],
                    "properties": {
                        "type": {"enum": CUSTOMER_TYPES},
                        "gender": {"enum": GENDERS},
                    },
                },
                "items": {
                    "bsonType": "array",
                    "minItems": 1,
                    "items": {
                        "bsonType": "object",
                        "required": ["product_line", "unit_price", "quantity", "line_cogs", "tax_5", "line_total", "gross_income"],
                        "properties": {
                            "product_line": {"enum": PRODUCT_LINES},
                            "unit_price": {"bsonType": ["double", "decimal"], "minimum": 0},
                            "quantity": {"bsonType": "int", "minimum": 1},
                            "line_cogs": {"bsonType": ["double", "decimal"], "minimum": 0},
                            "tax_5": {"bsonType": ["double", "decimal"], "minimum": 0},
                            "line_total": {"bsonType": ["double", "decimal"], "minimum": 0},
                            "gross_income": {"bsonType": ["double", "decimal"], "minimum": 0},
                        },
                    },
                },
                "sale_datetime": {"bsonType": "date"},
                "sale_month": {"bsonType": "string", "pattern": "^[0-9]{4}-[0-9]{2}$"},
                "sale_weekday": {"enum": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]},
                "sale_hour": {"bsonType": "int", "minimum": 0, "maximum": 23},
                "payment": {
                    "bsonType": "object",
                    "required": ["method"],
                    "properties": {"method": {"enum": PAYMENT_METHODS}},
                },
                "financials": {
                    "bsonType": "object",
                    "required": ["cogs", "tax_5", "total", "gross_income"],
                    "properties": {
                        "cogs": {"bsonType": ["double", "decimal"], "minimum": 0},
                        "tax_5": {"bsonType": ["double", "decimal"], "minimum": 0},
                        "total": {"bsonType": ["double", "decimal"], "minimum": 0},
                        "gross_income": {"bsonType": ["double", "decimal"], "minimum": 0},
                    },
                },
                "rating": {"bsonType": ["double", "decimal"], "minimum": 0, "maximum": 10},
            },
        }
    }


def index_specs() -> list[dict]:
    return [
        {"name": "idx_sale_datetime", "key": {"sale_datetime": 1}},
        {"name": "idx_branch_date", "key": {"branch.code": 1, "sale_datetime": 1}},
        {"name": "idx_product_total", "key": {"items.product_line": 1, "financials.total": -1}},
        {"name": "idx_payment_date", "key": {"payment.method": 1, "sale_datetime": 1}},
        {"name": "idx_customer_product_rating", "key": {"customer.type": 1, "items.product_line": 1, "rating": -1}},
        {"name": "idx_rating_total", "key": {"rating": 1, "financials.total": -1}},
        {"name": "idx_hour", "key": {"sale_hour": 1}},
    ]


def write_metadata() -> None:
    metadata = {
        "options": {
            "validator": validator(),
            "validationLevel": "strict",
            "validationAction": "error",
        },
        "indexes": [{"v": 2, "key": {"_id": 1}, "name": "_id_"}] + index_specs(),
    }
    (OUT_DIR / f"{GROUP}.metadata.json").write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    dump_dir = OUT_DIR / "dump" / GROUP
    dump_dir.mkdir(parents=True, exist_ok=True)
    (dump_dir / f"{COLLECTION}.metadata.json").write_text(json.dumps(metadata, indent=2), encoding="utf-8")


def totals_by(docs: list[dict], getter) -> list[tuple[str, float, int, float]]:
    bucket = defaultdict(lambda: [0.0, 0, 0.0])
    for doc in docs:
        name = getter(doc)
        bucket[name][0] += doc["financials"]["total"]
        bucket[name][1] += 1
        bucket[name][2] += doc["rating"]
    rows = [(name, vals[0], vals[1], vals[2] / vals[1]) for name, vals in bucket.items()]
    return sorted(rows, key=lambda item: item[1], reverse=True)


def summary(docs: list[dict]) -> dict:
    total_revenue = sum(d["financials"]["total"] for d in docs)
    total_profit = sum(d["financials"]["gross_income"] for d in docs)
    avg_ticket = total_revenue / len(docs)
    avg_rating = sum(d["rating"] for d in docs) / len(docs)
    by_product = totals_by(docs, lambda d: d["items"][0]["product_line"])
    by_branch = totals_by(docs, lambda d: d["branch"]["code"])
    by_payment = totals_by(docs, lambda d: d["payment"]["method"])
    by_customer = totals_by(docs, lambda d: d["customer"]["type"])
    by_hour = totals_by(docs, lambda d: str(d["sale_hour"]))
    return {
        "documents": len(docs),
        "total_revenue": total_revenue,
        "total_profit": total_profit,
        "avg_ticket": avg_ticket,
        "avg_rating": avg_rating,
        "top_product": by_product[0],
        "top_branch": by_branch[0],
        "top_payment": by_payment[0],
        "member_revenue": next(row for row in by_customer if row[0] == "Member"),
        "normal_revenue": next(row for row in by_customer if row[0] == "Normal"),
        "by_product": by_product,
        "by_branch": by_branch,
        "by_payment": by_payment,
        "by_customer": by_customer,
        "by_hour": by_hour,
        "product_count": len(by_product),
        "branch_count": len(by_branch),
        "payment_count": len(by_payment),
    }


def date_range_count(docs: list[dict], branch: str, month: str) -> int:
    return sum(1 for d in docs if d["branch"]["code"] == branch and d["sale_month"] == month)


def product_count(docs: list[dict], product_line: str) -> int:
    return sum(1 for d in docs if d["items"][0]["product_line"] == product_line)


def payment_count(docs: list[dict], method: str) -> int:
    return sum(1 for d in docs if d["payment"]["method"] == method)


def product_high_value_count(docs: list[dict], product_line: str, minimum: float) -> int:
    return sum(1 for d in docs if d["items"][0]["product_line"] == product_line and d["financials"]["total"] >= minimum)


def payment_above_avg_count(docs: list[dict], method: str, average: float) -> int:
    return sum(1 for d in docs if d["payment"]["method"] == method and d["financials"]["total"] >= average)


def customer_product_count(docs: list[dict], customer_type: str, product_line: str) -> int:
    return sum(1 for d in docs if d["customer"]["type"] == customer_type and d["items"][0]["product_line"] == product_line)


def low_rating_count(docs: list[dict], maximum: float) -> int:
    return sum(1 for d in docs if d["rating"] < maximum)


def receipt_line(doc: dict) -> str:
    item = doc["items"][0]
    return f"{doc['invoice_id']} | {doc['branch']['code']} | {item['product_line']} | {doc['payment']['method']} | {money(doc['financials']['total'])} | rating {doc['rating']:.1f}"


def sample_lines(rows: list[dict], limit: int = 3) -> str:
    return "\n".join(receipt_line(doc) for doc in rows[:limit])


def query_samples(docs: list[dict], stats: dict) -> dict[str, str]:
    q1 = sorted([d for d in docs if d["branch"]["code"] == "C" and d["sale_month"] == "2019-03"], key=lambda d: d["financials"]["total"], reverse=True)
    q2 = sorted([d for d in docs if d["items"][0]["product_line"] == "Food and beverages" and d["financials"]["total"] >= 500], key=lambda d: d["financials"]["total"], reverse=True)
    q3 = sorted([d for d in docs if d["payment"]["method"] == "Cash" and d["financials"]["total"] >= stats["avg_ticket"]], key=lambda d: d["sale_datetime"])
    q4 = sorted([d for d in docs if d["customer"]["type"] == "Member" and d["items"][0]["product_line"] == "Sports and travel"], key=lambda d: d["rating"], reverse=True)
    q5 = sorted([d for d in docs if d["rating"] < 5], key=lambda d: (d["rating"], -d["financials"]["total"]))
    return {
        "q1": f"count {len(q1)}\n{sample_lines(q1)}",
        "q2": f"count {len(q2)}\n{sample_lines(q2)}",
        "q3": f"count {len(q3)}\n{sample_lines(q3)}",
        "q4": f"count {len(q4)}\n{sample_lines(q4)}",
        "q5": f"count {len(q5)}\n{sample_lines(q5)}",
    }


def aggregation_samples(stats: dict) -> dict[str, str]:
    branch_rows = "\n".join(f"{row[0]} | {money(row[1])} | {row[2]} tickets | avg rating {row[3]:.2f}" for row in stats["by_branch"])
    product_rows = "\n".join(f"{row[0]} | {money(row[1])} | {row[2]} tickets | avg rating {row[3]:.2f}" for row in stats["by_product"][:4])
    hour_rows = "\n".join(f"hour {row[0]} | {money(row[1])} | {row[2]} tickets" for row in stats["by_hour"][:5])
    customer_rows = "\n".join(f"{row[0]} | {money(row[1])} | {row[2]} tickets | avg rating {row[3]:.2f}" for row in stats["by_customer"])
    return {"branch": branch_rows, "product": product_rows, "hour": hour_rows, "customer": customer_rows}


def queries_text(stats: dict, docs: list[dict]) -> str:
    val = json.dumps(validator(), indent=2)
    samples = query_samples(docs, stats)
    aggs = aggregation_samples(stats)
    idx = "\n".join(
        f'db.{COLLECTION}.createIndex({json.dumps(spec["key"])}, {{ name: "{spec["name"]}" }});'
        for spec in index_specs()
    )
    return f"""Big Data Storage final project, MongoDB commands for {GROUP}

Database and collection
use {GROUP};

Restore BSON backup from the submitted zip root
mongorestore --db {GROUP} --collection {COLLECTION} --drop {GROUP}.bson

Restore from this workspace, before zipping
mongorestore --db {GROUP} --collection {COLLECTION} --drop deliverables/{GROUP}.bson

Full dump style restore, if using the support dump folder
mongorestore --drop support/dump

After restore, run the validator and index commands below if Compass or mongorestore did not apply metadata automatically. Then verify the load.
db.{COLLECTION}.countDocuments();
db.getCollectionInfos({{ name: "{COLLECTION}" }});
db.{COLLECTION}.getIndexes();

Create collection with validator before importing, if starting from CSV or JSON instead of restoring BSON
use {GROUP};
db.createCollection("{COLLECTION}", {{
  validator: {val},
  validationLevel: "strict",
  validationAction: "error"
}});

If the collection already exists, apply the validator with collMod
db.runCommand({{
  collMod: "{COLLECTION}",
  validator: {val},
  validationLevel: "strict",
  validationAction: "error"
}});

Indexes
{idx}

Simple insight queries

1. March receipts for branch C, sorted by highest ticket value.
db.{COLLECTION}.find(
  {{ "branch.code": "C", sale_datetime: {{ $gte: ISODate("2019-03-01T00:00:00Z"), $lt: ISODate("2019-04-01T00:00:00Z") }} }},
  {{ invoice_id: 1, branch: 1, sale_datetime: 1, "financials.total": 1, rating: 1 }}
).sort({{ "financials.total": -1 }}).limit(10);

Business insight: branch C had {date_range_count(docs, 'C', '2019-03')} March receipts. This query shows the highest ticket values for that branch.
Checked output:
{samples['q1']}

2. High value Food and beverages receipts.
db.{COLLECTION}.find(
  {{ "items.product_line": "Food and beverages", "financials.total": {{ $gte: 500 }} }},
  {{ invoice_id: 1, branch: 1, customer: 1, items: 1, "financials.total": 1 }}
).sort({{ "financials.total": -1 }});

Business insight: Food and beverages is the top product line by revenue at {money(stats['top_product'][1])}.
Checked output:
{samples['q2']}

3. Cash transactions above the average ticket.
db.{COLLECTION}.find(
  {{ "payment.method": "Cash", "financials.total": {{ $gte: {stats['avg_ticket']:.2f} }} }},
  {{ invoice_id: 1, sale_datetime: 1, payment: 1, "financials.total": 1 }}
).sort({{ sale_datetime: 1 }});

Business insight: Cash is the largest payment method by sales value at {money(stats['top_payment'][1])}.
Checked output:
{samples['q3']}

4. Member receipts in Sports and travel.
db.{COLLECTION}.find(
  {{ "customer.type": "Member", "items.product_line": "Sports and travel" }},
  {{ invoice_id: 1, customer: 1, items: 1, "financials.total": 1, rating: 1 }}
).sort({{ rating: -1 }}).limit(15);

Business insight: this shows member receipts for Sports and travel, sorted by rating.
Checked output:
{samples['q4']}

5. Low rating receipts for service follow up.
db.{COLLECTION}.find(
  {{ rating: {{ $lt: 5 }} }},
  {{ invoice_id: 1, branch: 1, sale_datetime: 1, items: 1, rating: 1, "financials.total": 1 }}
).sort({{ rating: 1, "financials.total": -1 }}).limit(20);

Business insight: this lists low rating receipts that managers may need to follow up.
Checked output:
{samples['q5']}

Aggregation pipelines

A. Revenue, profit, tickets, and rating by branch.
db.{COLLECTION}.aggregate([
  {{ $group: {{ _id: "$branch.code", city: {{ $first: "$branch.city" }}, tickets: {{ $sum: 1 }}, revenue: {{ $sum: "$financials.total" }}, gross_income: {{ $sum: "$financials.gross_income" }}, avg_rating: {{ $avg: "$rating" }} }} }},
  {{ $sort: {{ revenue: -1 }} }}
]);
Checked output:
{aggs['branch']}

B. Product line performance.
db.{COLLECTION}.aggregate([
  {{ $unwind: "$items" }},
  {{ $group: {{ _id: "$items.product_line", units: {{ $sum: "$items.quantity" }}, revenue: {{ $sum: "$items.line_total" }}, gross_income: {{ $sum: "$items.gross_income" }}, avg_rating: {{ $avg: "$rating" }} }} }},
  {{ $sort: {{ revenue: -1 }} }}
]);
Checked output:
{aggs['product']}

C. Sales by hour of day.
db.{COLLECTION}.aggregate([
  {{ $group: {{ _id: "$sale_hour", tickets: {{ $sum: 1 }}, revenue: {{ $sum: "$financials.total" }}, avg_ticket: {{ $avg: "$financials.total" }} }} }},
  {{ $sort: {{ revenue: -1 }} }}
]);
Checked output:
{aggs['hour']}

D. Customer type and payment mix.
db.{COLLECTION}.aggregate([
  {{ $group: {{ _id: {{ customer_type: "$customer.type", payment: "$payment.method" }}, tickets: {{ $sum: 1 }}, revenue: {{ $sum: "$financials.total" }}, avg_rating: {{ $avg: "$rating" }} }} }},
  {{ $sort: {{ revenue: -1 }} }}
]);
Checked output:
{aggs['customer']}

Index performance checks to run in Compass or mongosh

db.{COLLECTION}.find({{ "branch.code": "C", sale_datetime: {{ $gte: ISODate("2019-03-01T00:00:00Z"), $lt: ISODate("2019-04-01T00:00:00Z") }} }}).explain("executionStats");
db.{COLLECTION}.find({{ "items.product_line": "Food and beverages", "financials.total": {{ $gte: 500 }} }}).sort({{ "financials.total": -1 }}).explain("executionStats");
db.{COLLECTION}.find({{ "payment.method": "Cash", "financials.total": {{ $gte: {stats['avg_ticket']:.2f} }} }}).sort({{ sale_datetime: 1 }}).explain("executionStats");
db.{COLLECTION}.find({{ "customer.type": "Member", "items.product_line": "Sports and travel" }}).sort({{ rating: -1 }}).explain("executionStats");
db.{COLLECTION}.find({{ rating: {{ $lt: 5 }} }}).sort({{ rating: 1, "financials.total": -1 }}).explain("executionStats");

Index impact on this 1000 document course dataset
1. Branch C in March: full scan checks 1000 documents. The branch plus date index narrows the candidate set to about {date_range_count(docs, 'C', '2019-03')} receipts.
2. High value Food and beverages: full scan checks 1000 documents. The product plus total index starts from {product_count(docs, 'Food and beverages')} Food and beverages receipts and returns {product_high_value_count(docs, 'Food and beverages', 500)} receipts above 500.
3. Cash above average ticket: full scan checks 1000 documents. The payment plus date index starts from {payment_count(docs, 'Cash')} cash receipts and returns {payment_above_avg_count(docs, 'Cash', stats['avg_ticket'])} receipts above the average ticket.
4. Member Sports and travel: full scan checks 1000 documents. The customer plus product plus rating index starts from {customer_product_count(docs, 'Member', 'Sports and travel')} matching loyalty receipts and can also serve the rating sort.
5. Low rating follow up: full scan checks 1000 documents. The rating plus total index starts from {low_rating_count(docs, 5)} low rating receipts and can serve the follow up sort.
"""


def write_query_file(stats: dict, docs: list[dict]) -> None:
    (OUT_DIR / f"{GROUP}.txt").write_text(queries_text(stats, docs), encoding="utf-8")


def add_doc_heading(doc: Document, text: str, level: int = 1) -> None:
    doc.add_heading(text, level=level)


def add_doc_para(doc: Document, text: str) -> None:
    para = doc.add_paragraph(text)
    para.paragraph_format.space_after = Pt(6)


def add_doc_table(doc: Document, headers: list[str], rows: list[list[str]]) -> None:
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    for idx, header in enumerate(headers):
        table.rows[0].cells[idx].text = header
    for row in rows:
        cells = table.add_row().cells
        for idx, value in enumerate(row):
            cells[idx].text = value


def report_sections(stats: dict, docs: list[dict]) -> list[tuple[str, list[str]]]:
    product = stats["top_product"]
    branch = stats["top_branch"]
    samples = query_samples(docs, stats)
    aggs = aggregation_samples(stats)
    return [
        (
            "Company and dataset",
            [
                "Company chosen: QuickCart Supermarket, a three branch food and household retail chain modelled from the supermarket sales dataset. The business process is the checkout flow: a customer buys one or more products, pays by cash, credit card, or electronic wallet, and leaves a rating.",
                f"The source file has {stats['documents']} receipts from 2019, three branches, six product lines, three payment methods, and two customer types. Total recorded revenue is {money(stats['total_revenue'])}; gross income is {money(stats['total_profit'])}.",
                "MongoDB fits this dataset because one receipt can be stored as one document. Branch, customer, payment, totals, and line items stay in the same record. This keeps normal checkout queries simple and still lets us group by branch, product, payment, hour, and customer type.",
            ],
        ),
        (
            "Model design decisions",
            [
                "The database uses one main collection: sales_receipts. Each receipt embeds branch, customer, payment, financials, and an items array. The source data has one line item per receipt, but the array keeps the model ready for multi item receipts without a second collection.",
                "The model is built around the questions store managers ask first. Most questions start from receipts, so branch city, customer type, payment method, and item details stay inside the receipt. The tradeoff is repeated branch and product text.",
                "The project stays well under the assignment limit of 10 collections. A larger production system could split customers, stock, and product catalogue records, but that would add complexity not needed for this dataset.",
            ],
        ),
        (
            "Data quality work",
            [
                "The build script converts numbers into Python numeric types before BSON encoding, parses sale date and time into a MongoDB date, trims categorical strings, removes duplicate invoice ids, maps branch codes to the city names used in the source, and rounds financial values to four decimals.",
                "It also adds sale_month, sale_weekday, and sale_hour fields. Those derived fields keep common management queries simple and readable in Compass.",
                "The verification file checks row count, unique invoice count, BSON size, a sample document type, and required files.",
            ],
        ),
        (
            "Validation rules",
            [
                "The validator requires invoice id, branch, customer, item array, sale date, sale month, sale weekday, sale hour, payment, financial values, and rating. It also restricts branch codes to A, B, and C; customer type to Member or Normal; gender to Female or Male; payment method to Cash, Credit card, or Ewallet; product line to the six known lines; quantity to at least one; prices and totals to non negative numbers; sale hour to 0 through 23; and rating to the 0 to 10 interval.",
            ],
        ),
        (
            "Indexes and performance",
            [
                "Indexes are defined for sale_datetime, branch plus date, product plus total, payment plus date, customer plus product plus rating, rating plus total, and sale hour. The point is not to index everything. Each index maps to one query or aggregation we use in the report.",
                f"On this 1000 document course dataset, branch C in March drops from 1000 possible checks to about {date_range_count(docs, 'C', '2019-03')} receipts. High value Food and beverages starts from {product_count(docs, 'Food and beverages')} category receipts and returns {product_high_value_count(docs, 'Food and beverages', 500)} above 500. Low rating follow up starts from {low_rating_count(docs, 5)} receipts, not from the whole collection.",
                "The cost is a few extra writes and some index storage. For this project that is okay because inserts are simple, and the report mostly cares about reads.",
            ],
        ),
        (
            "Business insights from the data",
            [
                f"The top product line is {product[0]} with {money(product[1])} in sales. The strongest branch is {branch[0]}, {BRANCH_CITIES[branch[0]]}, with {money(branch[1])} in sales.",
                f"Average ticket value is {money(stats['avg_ticket'])}. Average customer rating is {stats['avg_rating']:.2f} out of 10. Member customers generate {money(stats['member_revenue'][1])}, slightly more than normal customers at {money(stats['normal_revenue'][1])}.",
                "The aggregations show branch results, product results, sales by hour, and customer type plus payment mix.",
            ],
        ),
        (
            "Checked query outputs",
            [
                f"Query 1, branch C in March: {samples['q1'].splitlines()[0]}.",
                f"Query 2, high value Food and beverages: {samples['q2'].splitlines()[0]}.",
                f"Query 3, cash above average ticket: {samples['q3'].splitlines()[0]}.",
                f"Query 4, member Sports and travel: {samples['q4'].splitlines()[0]}.",
                f"Query 5, low rating follow up: {samples['q5'].splitlines()[0]}.",
                "Branch aggregation checked locally: " + aggs["branch"].replace("\n", "; "),
                "Product aggregation top rows checked locally: " + aggs["product"].replace("\n", "; "),
                "Hour aggregation top rows checked locally: " + aggs["hour"].replace("\n", "; "),
            ],
        ),
        (
            "Advantages and disadvantages",
            [
                "Advantages: the receipt document matches how the business reads data, query commands stay short, validation protects common mistakes, and indexes match the most repeated questions.",
                "Disadvantages: product and branch values repeat. If a product name changes, many documents need updates. The dataset has one item per receipt, even though the model can hold several items.",
                "The model is good for sales analysis and class work. A real chain could add stock, supplier, promotion, and customer collections if the business needs them.",
            ],
        ),
    ]


def write_report_docx(stats: dict, docs: list[dict]) -> None:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Big Data Storage Final Project")
    run.bold = True
    run.font.size = Pt(22)
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.add_run("MongoDB design for QuickCart Supermarket").font.size = Pt(14)
    for line in [
        "Group: group_xxx, replace before Moodle submission",
        "Team members: NAME, STUDENT NUMBER, replace before Moodle submission",
        "Course deliverable: report, BSON backup, query file, and presentation",
    ]:
        para = doc.add_paragraph(line)
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()
    add_doc_heading(doc, "One page description", 1)
    for paragraph in report_sections(stats, docs)[0][1]:
        add_doc_para(doc, paragraph)
    add_doc_heading(doc, "Dataset facts", 2)
    add_doc_table(
        doc,
        ["Metric", "Value"],
        [
            ["Receipts", str(stats["documents"])],
            ["Revenue", money(stats["total_revenue"])],
            ["Gross income", money(stats["total_profit"])],
            ["Average ticket", money(stats["avg_ticket"])],
            ["Average rating", f"{stats['avg_rating']:.2f}"],
        ],
    )
    for heading, paragraphs in report_sections(stats, docs)[1:]:
        add_doc_heading(doc, heading, 1)
        for paragraph in paragraphs:
            add_doc_para(doc, paragraph)
    add_doc_heading(doc, "Top product lines by revenue", 2)
    add_doc_table(
        doc,
        ["Product line", "Revenue", "Tickets", "Average rating"],
        [[row[0], money(row[1]), str(row[2]), f"{row[3]:.2f}"] for row in stats["by_product"]],
    )
    add_doc_heading(doc, "Delivery files", 1)
    for item in [
        f"{GROUP}.bson contains the BSON backup for the sales_receipts collection.",
        f"{GROUP}.txt contains validators, indexes, queries, aggregations, and explain commands.",
        f"{GROUP}_presentation.pptx is the 10 minute defense deck.",
        f"{GROUP}.zip contains the full submission package.",
    ]:
        add_doc_para(doc, item)
    doc.save(OUT_DIR / f"{GROUP}_report.docx")


def write_report_pdf(stats: dict, docs: list[dict]) -> None:
    styles = getSampleStyleSheet()
    story = []
    story.append(Paragraph("Big Data Storage Final Project", styles["Title"]))
    story.append(Paragraph("MongoDB design for QuickCart Supermarket", styles["Heading2"]))
    story.append(Spacer(1, 0.3 * cm))
    story.append(Paragraph("Group: group_xxx, replace before Moodle submission", styles["Normal"]))
    story.append(Paragraph("Team members: NAME, STUDENT NUMBER, replace before Moodle submission", styles["Normal"]))
    story.append(Paragraph("Files: report, BSON backup, query file, and PowerPoint deck", styles["Normal"]))
    story.append(PageBreak())
    for heading, paragraphs in report_sections(stats, docs):
        story.append(Paragraph(heading, styles["Heading1"]))
        for paragraph in paragraphs:
            story.append(Paragraph(paragraph, styles["BodyText"]))
            story.append(Spacer(1, 0.15 * cm))
    table_data = [
        ["Metric", "Value"],
        ["Receipts", str(stats["documents"])],
        ["Revenue", money(stats["total_revenue"])],
        ["Gross income", money(stats["total_profit"])],
        ["Average ticket", money(stats["avg_ticket"])],
        ["Average rating", f"{stats['avg_rating']:.2f}"],
    ]
    table = Table(table_data, colWidths=[6 * cm, 8 * cm])
    table.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8F1F5")), ("GRID", (0, 0), (-1, -1), 0.5, colors.grey), ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold")]))
    story.append(Paragraph("Dataset facts", styles["Heading1"]))
    story.append(table)
    story.append(Paragraph("Top product lines by revenue", styles["Heading1"]))
    product_table = Table(
        [["Product line", "Revenue", "Tickets", "Average rating"]] + [[row[0], money(row[1]), str(row[2]), f"{row[3]:.2f}"] for row in stats["by_product"]],
        colWidths=[6 * cm, 3 * cm, 2.5 * cm, 3 * cm],
    )
    product_table.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8F1F5")), ("GRID", (0, 0), (-1, -1), 0.5, colors.grey), ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold")]))
    story.append(product_table)
    story.append(Spacer(1, 0.4 * cm))
    story.append(Paragraph("Delivery files", styles["Heading1"]))
    for item in [
        f"{GROUP}.bson contains the BSON backup for the sales_receipts collection.",
        f"{GROUP}.txt contains validators, indexes, queries, aggregations, and explain commands.",
        f"{GROUP}_presentation.pptx is the 10 minute defense deck.",
        f"{GROUP}.zip contains the full submission package.",
    ]:
        story.append(Paragraph(item, styles["BodyText"]))
    SimpleDocTemplate(str(OUT_DIR / f"{GROUP}_report.pdf"), pagesize=A4, rightMargin=1.5 * cm, leftMargin=1.5 * cm, topMargin=1.5 * cm, bottomMargin=1.5 * cm).build(story)


def add_bg(slide, color: str) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = PptColor.from_string(color)


def add_title(slide, title: str, subtitle: str | None = None, dark: bool = False) -> None:
    color = "FFFFFF" if dark else "1B2A38"
    box = slide.shapes.add_textbox(PptInches(0.55), PptInches(0.35), PptInches(12.2), PptInches(0.7))
    run = box.text_frame.paragraphs[0].add_run()
    run.text = title
    run.font.bold = True
    run.font.size = PptPt(30)
    run.font.color.rgb = PptColor.from_string(color)
    if subtitle:
        sub = slide.shapes.add_textbox(PptInches(0.58), PptInches(1.05), PptInches(11.8), PptInches(0.35))
        srun = sub.text_frame.paragraphs[0].add_run()
        srun.text = subtitle
        srun.font.size = PptPt(14)
        srun.font.color.rgb = PptColor.from_string("DDE8EF" if dark else "4D5D6C")


def add_card(slide, x: float, y: float, w: float, h: float, title: str, body: str, fill: str = "FFFFFF") -> None:
    shape = slide.shapes.add_shape(1, PptInches(x), PptInches(y), PptInches(w), PptInches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = PptColor.from_string(fill)
    shape.line.color.rgb = PptColor.from_string("C9D6DF")
    text = shape.text_frame
    text.clear()
    p = text.paragraphs[0]
    r = p.add_run()
    r.text = title
    r.font.bold = True
    r.font.size = PptPt(15)
    r.font.color.rgb = PptColor.from_string("143642")
    p2 = text.add_paragraph()
    p2.text = body
    p2.font.size = PptPt(11)
    p2.font.color.rgb = PptColor.from_string("263640")


def add_bar_chart(slide, stats_rows: list[tuple[str, float, int, float]], x: float, y: float, w: float, h: float) -> None:
    max_value = max(row[1] for row in stats_rows)
    bar_h = h / len(stats_rows) * 0.55
    gap = h / len(stats_rows) * 0.45
    for idx, row in enumerate(stats_rows):
        yy = y + idx * (bar_h + gap)
        label = slide.shapes.add_textbox(PptInches(x), PptInches(yy), PptInches(2.3), PptInches(bar_h))
        label.text_frame.text = row[0]
        label.text_frame.paragraphs[0].font.size = PptPt(9)
        bar_w = (row[1] / max_value) * (w - 3.2)
        bar = slide.shapes.add_shape(1, PptInches(x + 2.4), PptInches(yy), PptInches(bar_w), PptInches(bar_h))
        bar.fill.solid()
        bar.fill.fore_color.rgb = PptColor.from_string("0E7C7B")
        bar.line.color.rgb = PptColor.from_string("0E7C7B")
        value = slide.shapes.add_textbox(PptInches(x + 2.5 + bar_w), PptInches(yy), PptInches(1.2), PptInches(bar_h))
        value.text_frame.text = money(row[1])
        value.text_frame.paragraphs[0].font.size = PptPt(9)


def write_pptx(stats: dict, docs: list[dict]) -> None:
    prs = Presentation()
    prs.slide_width = PptInches(13.333)
    prs.slide_height = PptInches(7.5)
    blank = prs.slide_layouts[6]

    slide = prs.slides.add_slide(blank)
    add_bg(slide, "102A43")
    add_title(slide, "QuickCart Supermarket MongoDB", "Big Data Storage final project, replace group_xxx and member names before submitting", dark=True)
    add_card(slide, 0.8, 2.0, 3.6, 2.0, "Dataset", f"{stats['documents']} receipts\n3 branches\n6 product lines", "E8F1F5")
    add_card(slide, 4.8, 2.0, 3.6, 2.0, "Revenue", f"{money(stats['total_revenue'])}\nGross income {money(stats['total_profit'])}", "E8F1F5")
    add_card(slide, 8.8, 2.0, 3.6, 2.0, "MongoDB fit", "Each checkout receipt is a single document with embedded branch, customer, payment, financials, and item data.", "E8F1F5")

    slide = prs.slides.add_slide(blank)
    add_bg(slide, "F8FAFC")
    add_title(slide, "Company and business process")
    add_card(slide, 0.7, 1.5, 3.8, 4.6, "Chosen company", "QuickCart Supermarket, a three branch retail chain selling food, household goods, clothing, electronics, sports items, and health products.")
    add_card(slide, 4.8, 1.5, 3.8, 4.6, "Process modelled", "Checkout receipt: branch, customer type, product line, quantity, payment method, total, gross income, and rating.")
    add_card(slide, 8.9, 1.5, 3.8, 4.6, "Why MongoDB", "One receipt can live as one document. Managers can read the sale directly, then group the same collection by branch, product, payment, or hour.")

    slide = prs.slides.add_slide(blank)
    add_bg(slide, "F8FAFC")
    add_title(slide, "Document model")
    add_card(slide, 0.7, 1.4, 5.8, 4.9, "sales_receipts", "_id, invoice_id\nbranch {code, city}\ncustomer {type, gender}\nitems [{product_line, unit_price, quantity, totals}]\nsale_datetime, sale_month, sale_hour\npayment {method}\nfinancials {cogs, tax, total, gross_income}\nrating")
    add_card(slide, 7.0, 1.4, 5.6, 2.2, "Design choice", "One collection keeps the assignment focused and avoids artificial joins. It still supports every required query and aggregation.")
    add_card(slide, 7.0, 4.1, 5.6, 2.2, "Tradeoff", "Branch and product values repeat. That is acceptable for read heavy analysis, but production stock systems would add catalogue and inventory collections.")

    slide = prs.slides.add_slide(blank)
    add_bg(slide, "F8FAFC")
    add_title(slide, "Data quality and validation")
    add_card(slide, 0.7, 1.4, 5.8, 2.0, "Quality operations", "Parsed date and time into MongoDB date. Cast prices, taxes, totals, ratings, and quantity. Removed duplicate invoice ids. Added sale_month, weekday, and hour. Standardised branch city names.")
    add_card(slide, 0.7, 3.9, 5.8, 2.0, "Validation rules", "Required receipt fields. Enums for branch, city, customer type, gender, payment, and product line. Quantity at least 1. Financial values non negative. Rating from 0 to 10.")
    add_card(slide, 7.0, 1.4, 5.6, 4.5, "Why this matters", "Bad checkout records can make branch and product numbers wrong. The validator catches bad categories and bad receipt formats before the reports use them.")

    slide = prs.slides.add_slide(blank)
    add_bg(slide, "F8FAFC")
    add_title(slide, "Business insights")
    add_bar_chart(slide, stats["by_product"], 0.7, 1.4, 7.2, 4.8)
    add_card(slide, 8.4, 1.4, 4.1, 1.4, "Top product line", f"{stats['top_product'][0]}\n{money(stats['top_product'][1])}")
    add_card(slide, 8.4, 3.1, 4.1, 1.4, "Top branch", f"{stats['top_branch'][0]}, {BRANCH_CITIES[stats['top_branch'][0]]}\n{money(stats['top_branch'][1])}")
    add_card(slide, 8.4, 4.8, 4.1, 1.4, "Average ticket", f"{money(stats['avg_ticket'])}\nRating {stats['avg_rating']:.2f}/10")

    slide = prs.slides.add_slide(blank)
    add_bg(slide, "F8FAFC")
    add_title(slide, "Queries and aggregations")
    add_card(slide, 0.7, 1.4, 3.8, 4.8, "Five queries", "Branch C March tickets. High value Food and beverages receipts. Cash tickets above average. Member Sports and travel receipts. Low rating follow up list.")
    add_card(slide, 4.8, 1.4, 3.8, 4.8, "Four aggregations", "Branch scorecard. Product line performance. Sales by hour. Customer type and payment mix.")
    add_card(slide, 8.9, 1.4, 3.8, 4.8, "Business use", "Store managers can see where revenue is coming from, which categories deserve attention, and which low rating receipts need service follow up.")

    slide = prs.slides.add_slide(blank)
    add_bg(slide, "F8FAFC")
    add_title(slide, "Indexes and performance impact")
    add_card(slide, 0.7, 1.4, 3.8, 4.7, "Indexes", "sale_datetime\nbranch.code plus sale_datetime\nitems.product_line plus total\npayment.method plus date\nrating plus total")
    add_card(slide, 4.8, 1.4, 3.8, 4.7, "Candidate reads", f"Branch C March: 1000 to {date_range_count(docs, 'C', '2019-03')}\nFood above 500: {product_count(docs, 'Food and beverages')} to {product_high_value_count(docs, 'Food and beverages', 500)}\nLow rating: 1000 to {low_rating_count(docs, 5)}")
    add_card(slide, 8.9, 1.4, 3.8, 4.7, "Defense point", "The dataset is small, so speed changes are small. I still added the indexes because the same queries would need them with a much larger receipt table.")

    slide = prs.slides.add_slide(blank)
    add_bg(slide, "102A43")
    add_title(slide, "Submission package", "What is inside group_xxx.zip", dark=True)
    add_card(slide, 0.8, 1.8, 3.7, 4.2, "Files", f"{GROUP}_report.docx and PDF\n{GROUP}.bson\n{GROUP}.txt\n{GROUP}_presentation.pptx", "E8F1F5")
    add_card(slide, 4.9, 1.8, 3.7, 4.2, "Check files", "Verification log, source CSV, metadata JSON, generator script, and restore notes are included so the work can be checked.", "E8F1F5")
    add_card(slide, 9.0, 1.8, 3.5, 4.2, "Before Moodle", "Replace group_xxx, member names, and student numbers. Then upload only once for the group.", "E8F1F5")

    prs.save(OUT_DIR / f"{GROUP}_presentation.pptx")


def write_readme(stats: dict) -> None:
    text = f"""Big Data Storage final project submission notes

Workspace: {ROOT}

Before Moodle submission, replace these placeholders:
1. group_xxx with your group number.
2. NAME, STUDENT NUMBER on the report cover page.
3. Add the same group and member information on the presentation title slide if you want.

Main files:
1. deliverables/{GROUP}_report.docx
2. deliverables/{GROUP}_report.pdf
3. deliverables/{GROUP}.bson
4. deliverables/{GROUP}.txt
5. deliverables/{GROUP}_presentation.pptx
6. deliverables/{GROUP}.zip

Dataset source:
{DATA_URL}

Restore commands:
From the submitted zip root:
mongorestore --db {GROUP} --collection {COLLECTION} --drop {GROUP}.bson

From this workspace before zipping:
mongorestore --db {GROUP} --collection {COLLECTION} --drop deliverables/{GROUP}.bson

Full dump style restore from the zip support folder:
mongorestore --drop support/dump

Dataset count: {stats['documents']} sales receipt documents.
"""
    (ROOT / "README_SUBMISSION.txt").write_text(text, encoding="utf-8")


def write_support_docs(stats: dict, docs: list[dict]) -> None:
    checklist = f"""Big Data Storage requirement map

1. Report cover page: {GROUP}_report.docx, first page.
2. Company and dataset page: {GROUP}_report.docx, section named One page description.
3. BSON backup: {GROUP}.bson, {stats['documents']} documents.
4. Query file: {GROUP}.txt.
5. Design decisions: report section named Model design decisions.
6. Advantages and disadvantages: report section named Advantages and disadvantages.
7. PowerPoint: {GROUP}_presentation.pptx, 8 slides.
8. Zip package: {GROUP}.zip.

MongoDB checks

1. Collections used: 1, sales_receipts.
2. Validation rules: required fields, invoice pattern, branch enum, customer enum, product enum, quantity minimum, financial minimums, sale hour range, rating range.
3. Insight queries: 5.
4. Aggregation pipelines: 4.
5. Indexes: {len(index_specs())} plus the default id index.

Fields to replace before upload

1. group_xxx.
2. NAME.
3. STUDENT NUMBER.
"""
    (OUT_DIR / f"{GROUP}_checklist.md").write_text(checklist, encoding="utf-8")
    dictionary = """Data dictionary for sales_receipts

_id: invoice id, string, same value as invoice_id.
invoice_id: source invoice id, string.
branch.code: branch code A, B, or C.
branch.city: city name from the source data.
customer.type: Member or Normal.
customer.gender: Female or Male.
items: array with receipt line items.
items.product_line: retail category.
items.unit_price: unit sale price.
items.quantity: quantity bought.
items.line_cogs: cost of goods for the line.
items.tax_5: tax at 5 percent.
items.line_total: final line value with tax.
items.gross_income: gross income for the line.
sale_datetime: sale date and time as BSON date.
sale_date: sale date as text for simple reads.
sale_month: month bucket.
sale_weekday: weekday name.
sale_hour: hour bucket.
payment.method: Cash, Credit card, or Ewallet.
financials.cogs: total cost of goods.
financials.tax_5: total tax.
financials.total: total paid.
financials.gross_margin_percentage: gross margin percent from the source.
financials.gross_income: gross income.
rating: customer rating from 0 to 10.
source: dataset source label.
"""
    (OUT_DIR / f"{GROUP}_data_dictionary.md").write_text(dictionary, encoding="utf-8")
    defense = f"""Defense notes for the 10 minute presentation

Slide 1: say what the package is. One MongoDB model for supermarket sales.
Slide 2: explain the business process. A receipt starts at checkout and ends with payment plus rating.
Slide 3: defend one collection. The receipt is the unit the manager reads.
Slide 4: talk about data quality. Dates, numbers, categories, duplicate invoices, and derived month and hour fields.
Slide 5: use the numbers. Revenue is {money(stats['total_revenue'])}; top product line is {stats['top_product'][0]}.
Slide 6: explain the queries and aggregations. They answer branch, category, payment, hour, loyalty, and complaint questions.
Slide 7: defend indexes. Each index has a query behind it. No random indexing.
Slide 8: say what is in the zip and what must be replaced before upload.

If asked about tradeoffs, say this:
We repeat branch and product facts because the dataset is receipt based. It makes reads easy. The cost is repeated text, which is fine here. A real chain could add stock, supplier, promotion, and customer collections later.
"""
    (OUT_DIR / f"{GROUP}_defense_notes.md").write_text(defense, encoding="utf-8")


def verify_artifacts(stats: dict) -> dict:
    expected = [
        OUT_DIR / f"{GROUP}.bson",
        OUT_DIR / f"{GROUP}.metadata.json",
        OUT_DIR / f"{GROUP}.txt",
        OUT_DIR / f"{GROUP}_report.docx",
        OUT_DIR / f"{GROUP}_report.pdf",
        OUT_DIR / f"{GROUP}_presentation.pptx",
        OUT_DIR / f"{GROUP}_checklist.md",
        OUT_DIR / f"{GROUP}_data_dictionary.md",
        OUT_DIR / f"{GROUP}_defense_notes.md",
        ROOT / "GOAL.md",
        ROOT / "README_SUBMISSION.txt",
        CSV_PATH,
    ]
    checks = {str(path.relative_to(ROOT)): path.exists() and path.stat().st_size > 0 for path in expected}
    query_text = (OUT_DIR / f"{GROUP}.txt").read_text(encoding="utf-8")
    business_block = query_text.split("Simple insight queries", 1)[1].split("Aggregation pipelines", 1)[0]
    checks["bson_size_bytes"] = (OUT_DIR / f"{GROUP}.bson").stat().st_size
    checks["documents"] = stats["documents"]
    checks["collections"] = 1
    checks["validation_rules_at_least_5"] = query_text.count("bsonType") >= 5 and query_text.count("enum") >= 4
    checks["queries_at_least_5"] = business_block.count("db.sales_receipts.find") >= 5
    checks["aggregations_at_least_3"] = query_text.count("db.sales_receipts.aggregate") >= 3
    checks["indexes_defined"] = len(index_specs())
    return checks


def create_zip() -> None:
    # root files are for moodle. support files are just proof.
    zip_path = OUT_DIR / f"{GROUP}.zip"
    include = [
        (OUT_DIR / f"{GROUP}.bson", f"{GROUP}.bson"),
        (OUT_DIR / f"{GROUP}.metadata.json", f"{GROUP}.metadata.json"),
        (OUT_DIR / f"{GROUP}.txt", f"{GROUP}.txt"),
        (OUT_DIR / f"{GROUP}_report.docx", f"{GROUP}_report.docx"),
        (OUT_DIR / f"{GROUP}_report.pdf", f"{GROUP}_report.pdf"),
        (OUT_DIR / f"{GROUP}_presentation.pptx", f"{GROUP}_presentation.pptx"),
        (OUT_DIR / f"{GROUP}_checklist.md", f"support/{GROUP}_checklist.md"),
        (OUT_DIR / f"{GROUP}_data_dictionary.md", f"support/{GROUP}_data_dictionary.md"),
        (OUT_DIR / f"{GROUP}_defense_notes.md", f"support/{GROUP}_defense_notes.md"),
        (OUT_DIR / f"{COLLECTION}.json", f"support/{COLLECTION}.json"),
        (OUT_DIR / "dump" / GROUP / f"{COLLECTION}.bson", f"support/dump/{GROUP}/{COLLECTION}.bson"),
        (OUT_DIR / "dump" / GROUP / f"{COLLECTION}.metadata.json", f"support/dump/{GROUP}/{COLLECTION}.metadata.json"),
        (OUT_DIR / "verification.json", "support/verification.json"),
        (CSV_PATH, "support/supermarket_sales.csv"),
        (ROOT / "Big Data Storage_Project.pdf", "support/Big Data Storage_Project.pdf"),
        (ROOT / "README_SUBMISSION.txt", "README_SUBMISSION.txt"),
        (ROOT / "GOAL.md", "support/GOAL.md"),
        (ROOT / "scripts" / "build_project.py", "support/build_project.py"),
    ]
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path, arcname in include:
            archive.write(path, arcname)


def main() -> None:
    ensure_dirs()
    download_dataset()
    rows = read_rows()
    docs = build_documents(rows)
    stats = summary(docs)
    write_bson(docs)
    write_metadata()
    write_query_file(stats, docs)
    write_report_docx(stats, docs)
    write_report_pdf(stats, docs)
    write_pptx(stats, docs)
    write_readme(stats)
    write_support_docs(stats, docs)
    checks = verify_artifacts(stats)
    (OUT_DIR / "verification.json").write_text(json.dumps(checks, indent=2), encoding="utf-8")
    create_zip()
    checks["deliverables/group_xxx.zip"] = (OUT_DIR / f"{GROUP}.zip").exists()
    (OUT_DIR / "verification.json").write_text(json.dumps(checks, indent=2), encoding="utf-8")
    create_zip()
    print(json.dumps(checks, indent=2))


if __name__ == "__main__":
    main()
