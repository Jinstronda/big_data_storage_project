Big Data Storage final project submission notes

Before Moodle submission, replace group_xxx with your real group number.
The member names and student numbers are already filled in the report and presentation.

Team members:
Daniyal Ahmad, 20241831
Artemii Alekseev, 20240645
Viktoriia German, 20240650
João Panizzutti, 20241624
Ivan Romanov, 20240639

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
