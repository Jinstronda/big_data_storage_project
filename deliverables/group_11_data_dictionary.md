Data dictionary for sales_receipts

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
