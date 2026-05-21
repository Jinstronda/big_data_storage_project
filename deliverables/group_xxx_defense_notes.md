Defense notes for the 10 minute presentation

Slide 1: say what the package is. One MongoDB model for supermarket sales.
Slide 2: explain the business process. A receipt starts at checkout and ends with payment plus rating.
Slide 3: defend one collection. The receipt is the unit the manager reads.
Slide 4: talk about data quality. Dates, numbers, categories, duplicate invoices, and derived month and hour fields.
Slide 5: use the numbers. Revenue is $322,966.75; top product line is Food and beverages.
Slide 6: explain the queries and aggregations. They answer branch, category, payment, hour, loyalty, and complaint questions.
Slide 7: defend indexes. Each index has a query behind it. No random indexing.
Slide 8: say what is in the zip and what must be replaced before upload.

If asked about tradeoffs, say this:
We repeat branch and product facts because the dataset is receipt based. It makes reads easy. The cost is repeated text, which is fine here. A real chain could add stock, supplier, promotion, and customer collections later.
