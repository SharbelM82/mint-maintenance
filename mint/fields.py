"""Single source of truth for the work-order fields.

Every column of the "Work Orders" sheet in the Mintresponse Excel tracker
(A10:AK510) is described here once. Forms, the list view, the Excel/CSV
import and export all read from this registry, so the web app and the
spreadsheet always stay in the same shape.

kind:
  text, textarea  free text
  date            stored as ISO yyyy-mm-dd
  money           USD amount, >= 0 (Excel validation: decimal >= 0)
  percent         stored as a fraction (0.4 = 40%), >= -100%
  int             whole number >= 0
  choice          one of `choices` (Excel drop-down lists)
computed=True means the value is a formula in Excel and is calculated by
mint.calc; it is never stored.
"""

WO_STATUSES = ["New", "Scheduled", "In progress", "Awaiting parts",
               "Awaiting approval", "Completed", "Cancelled"]
INVOICE_STATUSES = ["Not invoiced", "Draft", "Sent", "Partially paid",
                    "Paid", "Disputed", "Void"]
PAYMENT_METHODS = ["ACH", "Check", "Credit card", "Wire", "Cash", "Other"]

OPEN_WO_STATUSES = ["New", "Scheduled", "In progress", "Awaiting parts",
                    "Awaiting approval"]


class Field:
    def __init__(self, col, key, label, kind="text", computed=False,
                 choices=None, required=False, help=""):
        self.col = col
        self.key = key
        self.label = label
        self.kind = kind
        self.computed = computed
        self.choices = choices
        self.required = required
        self.help = help


FIELDS = [
    Field("A", "wo_number", "WO #", required=True,
          help="Work order number. Required and unique."),
    Field("B", "client", "Client"),
    Field("C", "assignees", "Assignees"),
    Field("D", "store_name", "Store name"),
    Field("E", "location", "Location"),
    Field("F", "eta", "ETA", "date"),
    Field("G", "parts_status", "Parts ordered / status",
          help="Parts details or ordering status. Does not add to direct cost."),
    Field("H", "direct_cost", "Direct cost", "money",
          help="Labor, subcontractor, material, travel and other direct job costs. "
               "Include parts cost here only once."),
    Field("I", "approved_nte", "Approved NTE", "money",
          help="Client-approved maximum charge. It is not revenue."),
    Field("J", "gross_profit", "Gross profit", "money", computed=True),
    Field("K", "gross_margin", "Gross margin %", "percent", computed=True),
    Field("L", "quote", "Expected NTE / quote", "money", computed=True),
    Field("M", "job_type", "Job type / scope"),
    Field("N", "handler", "Handler"),
    Field("O", "tech", "Tech"),
    Field("P", "notes", "Notes", "textarea"),
    Field("Q", "invoice_number", "Invoice #"),
    Field("R", "wo_status", "WO status", "choice", choices=WO_STATUSES),
    Field("S", "address", "Address"),
    Field("T", "date_in", "Date in", "date"),
    Field("U", "invoice_sent_date", "Invoice sent date", "date"),
    Field("V", "invoice_status", "Invoice status", "choice", choices=INVOICE_STATUSES),
    Field("W", "payment_method", "Payment method", "choice", choices=PAYMENT_METHODS),
    Field("X", "payment_due", "Payment due", "date", computed=True),
    Field("Y", "target_markup", "Target markup %", "percent",
          help="Enter 40 for 40%. Applied to direct cost to build the quote."),
    Field("Z", "invoice_revenue", "Invoice revenue", "money",
          help="Actual invoice amount excluding sales tax, after discounts or credits. "
               "Do not enter the NTE here."),
    Field("AA", "actual_markup", "Actual markup %", "percent", computed=True),
    Field("AB", "overhead", "Allocated overhead", "money",
          help="Enter 0 when none."),
    Field("AC", "admin_fees", "Payment / admin fees", "money",
          help="Enter 0 when none."),
    Field("AD", "other_expenses", "Other expenses / tax", "money",
          help="Enter 0 when none. Include allocated taxes for profit after tax."),
    Field("AE", "net_profit", "Net profit", "money", computed=True),
    Field("AF", "net_margin", "Net margin %", "percent", computed=True),
    Field("AG", "amount_received", "Amount received", "money",
          help="Excluding sales tax. Enter 0 if unpaid."),
    Field("AH", "balance_due", "Balance due", "money", computed=True),
    Field("AI", "payment_terms", "Payment terms (days)", "int",
          help="e.g. 30 for Net 30. Payment due = invoice sent date + terms."),
    Field("AJ", "nte_check", "NTE check", computed=True),
    Field("AK", "input_status", "Financial input status", computed=True),
]

BY_KEY = {f.key: f for f in FIELDS}
BY_LABEL = {f.label.lower(): f for f in FIELDS}
INPUT_FIELDS = [f for f in FIELDS if not f.computed]
INPUT_KEYS = [f.key for f in INPUT_FIELDS]

# Form layout: (section title, field keys). Computed fields show live results.
FORM_SECTIONS = [
    ("Work order", ["wo_number", "wo_status", "client", "store_name", "location",
                    "date_in", "eta", "job_type", "address"]),
    ("Team", ["assignees", "handler", "tech"]),
    ("Parts & notes", ["parts_status", "notes"]),
    ("Cost & quote", ["direct_cost", "target_markup", "quote", "approved_nte",
                      "nte_check"]),
    ("Invoice", ["invoice_number", "invoice_status", "invoice_sent_date",
                 "payment_terms", "payment_due", "invoice_revenue",
                 "gross_profit", "gross_margin", "actual_markup"]),
    ("Expenses & net profit", ["overhead", "admin_fees", "other_expenses",
                               "net_profit", "net_margin"]),
    ("Payment", ["payment_method", "amount_received", "balance_due",
                 "input_status"]),
]

# Columns shown in the compact list view (the "All columns" view shows FIELDS).
COMPACT_KEYS = ["wo_number", "client", "store_name", "location", "job_type",
                "wo_status", "date_in", "direct_cost", "approved_nte",
                "invoice_revenue", "gross_profit", "invoice_status",
                "payment_due", "balance_due", "input_status"]

# Free-text fields that get type-ahead suggestions from existing records.
SUGGEST_KEYS = ["client", "assignees", "store_name", "location", "job_type",
                "handler", "tech"]
