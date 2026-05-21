/goal Complete the Big Data Storage final project from /Users/joaopanizzutti/Downloads/Big Data Storage_Project.pdf inside /Users/joaopanizzutti/Downloads/big_data_storage_project_done until every deliverable is generated and verified.

Goal contract:
1. Outcome: produce a Moodle ready zip for the assignment with report, BSON database backup, query text file, PowerPoint presentation, source data, and reproducible scripts.
2. Verification surface: extracted PDF requirements, generated BSON file, query file, report file, presentation file, zip contents, script logs, and file existence checks.
3. Constraints: satisfy the PDF exactly, use at most 10 MongoDB collections, include at least 5 validation rules, at least 5 business insight queries, indexes with performance discussion, and at least 3 aggregations with one or more stages. Do not claim external MongoDB server proof unless a local server actually ran.
4. Boundaries: work only inside /Users/joaopanizzutti/Downloads/big_data_storage_project_done and read the PDF from Downloads. File edits and local script runs are allowed. No commits, pushes, deploys, credential changes, external sends, destructive deletes outside the workspace, or paid services.
5. Iteration policy: after every missing or invalid artifact, fix the smallest relevant piece, rerun the generator or verification, and keep the checklist alive.
6. Blocked stop condition: only stop early for missing team member names, student numbers, group number, unavailable network data source after trying a local fallback, or a required external service. If blocked, still generate placeholders and report exactly what João must replace.

Read first:
1. /Users/joaopanizzutti/Downloads/Big Data Storage_Project.pdf.
2. The extracted requirement text in /tmp/big_data_storage_project.txt if present.
3. Existing files in this workspace.

Deliverables to create:
1. deliverables/group_xxx_report.docx, with cover page, one page company and dataset description, model design decisions, data quality, validation rules, queries, aggregations, indexes, performance discussion, pros and cons, and conclusion.
2. deliverables/group_xxx.bson, a database backup artifact containing the modeled documents. If no mongodump tool exists, create a valid BSON sequence plus a restore note.
3. deliverables/group_xxx.txt, containing all MongoDB collection validators, index creation commands, 5 insight queries, 3 aggregations, and notes for how to run them in Compass or mongosh.
4. deliverables/group_xxx_presentation.pptx, maximum 10 minute presentation.
5. deliverables/group_xxx.zip, containing every required document plus source data and scripts.

Acceptance checkboxes:
1. [ ] The PDF requirements are extracted and mapped to artifact names.
2. [ ] A commercial business process and dataset are chosen and justified.
3. [ ] The MongoDB model has at most 10 collections.
4. [ ] Imported or generated documents have correct MongoDB data types.
5. [ ] At least 5 data quality operations are documented.
6. [ ] At least 5 validation rules are present as MongoDB JSON schema validators.
7. [ ] At least 5 different queries output business insights.
8. [ ] Indexes are defined and paired with a performance discussion.
9. [ ] At least 3 aggregation pipelines have one or more stages and improve dataset information.
10. [ ] BSON backup exists and is nonempty.
11. [ ] Report exists and contains all required sections.
12. [ ] Presentation exists and fits a 10 minute defense.
13. [ ] Zip exists and includes every required artifact.

Work contract:
1. Prefer a real, public, business dataset. If access is blocked, generate a realistic synthetic retail dataset and clearly label it as a course dataset derived from the described business process.
2. Use a model that fits MongoDB naturally: nested documents, arrays, references only where they reduce duplication, indexes matching the query workload, validation rules for entity integrity.
3. Keep the outputs professor friendly. The query file should be directly usable in MongoDB Compass or mongosh.
4. Use placeholders only for unknown group information: group number, member names, student numbers. Make placeholders obvious and easy to replace.
5. Run a verification script that checks counts, artifact existence, and zip contents.

Done when:
1. Every acceptance checkbox passes or has a clear placeholder blocker.
2. The final zip exists and has the expected files.
3. The final report says what João still needs to replace before Moodle submission, if anything.

Final report:
Return only: workspace path, zip path, files created, verification result, and the placeholder fields João must replace.