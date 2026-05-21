Big Data Storage final project submission notes

Before Moodle submission, replace these placeholders:
1. group_xxx with your group number.
2. NAME, STUDENT NUMBER on the report cover page.
3. Add the same group and member information on the presentation title slide if you want.

Main files:
1. deliverables/group_xxx_report.docx
2. deliverables/group_xxx_report.pdf
3. deliverables/group_xxx.bson
4. deliverables/group_xxx.txt
5. deliverables/group_xxx_presentation.pptx
6. deliverables/group_xxx.zip

Dataset source:
https://raw.githubusercontent.com/selva86/datasets/master/supermarket_sales.csv

Restore commands:
From the submitted zip root:
mongorestore --db group_xxx --collection sales_receipts --drop group_xxx.bson

Full dump style restore from the support folder:
mongorestore --drop support/dump

Dataset count: 1000 sales receipt documents.
