
This workbook is fully formula-driven. It currently contains realistic SAMPLE data so you can see how every number is calculated. Replace the sample rows in the raw tabs with your real data (via Power Query refresh OR manual paste) and every summary, KPI card and chart on the Dashboard recalculates automatically.


1. WHY THIS CAN’T BE A LIVE, SELF-REFRESHING FILE OUT OF THE BOX
A standalone .xlsx file cannot hold live API credentials or run a background service — Stripe and Xero both require an authenticated API call (Stripe: secret key; Xero: OAuth2 token that expires and must be refreshed). The safe, standard way to get this data into Excel is Power Query, which runs locally on your machine using YOUR credentials — Anthropic/Claude never sees or stores your keys.


2. CONNECTING STRIPE VIA POWER QUERY (refreshable, no add-in needed)
•  In Excel: Data tab → Get Data → From Other Sources → Blank Query → Advanced Editor. Paste the M code below, replacing YOUR_SECRET_KEY with your Stripe secret key (Developers → API keys in the Stripe Dashboard). Use a RESTRICTED key with read-only access to Charges/Payouts.

let
    Source = Json.Document(Web.Contents("https://api.stripe.com/v1/charges",
        [Headers=[Authorization="Bearer YOUR_SECRET_KEY"]])),
    data = Source[data],
    #"Converted to Table" = Table.FromList(data, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Expanded" = Table.ExpandRecordColumn(#"Converted to Table", "Column1",
        {"created","amount","currency","status","description","customer"})
in
    #"Expanded"

•  Load the result into the Stripe_Raw tab (Close & Load To → Existing Worksheet → A4), then right-click → Refresh whenever you want updated data. For date filtering / pagination beyond 100 records, add ?limit=100&starting_after=... query parameters or loop with Power Query’s List.Generate.


3. CONNECTING XERO VIA POWER QUERY (OAuth2)
Xero requires OAuth2 (no static API key). Steps:
•  Register an app at developer.xero.com to get a Client ID/Secret and complete the OAuth consent flow once to obtain a refresh token.
•  In Power Query, use Web.Contents with an Authorization: Bearer <access_token> header, calling https://api.xero.com/api.xro/2.0/Invoices and https://api.xero.com/api.xro/2.0/Bills — the access token expires every 30 minutes, so refresh tokens must be renewed programmatically (a small script/Power Automate flow, or a connector add-in, is the practical way to keep this running unattended).

•  Simplest alternative: in Xero, go to Business → Invoices (or Bills) → Export to Excel/CSV, then paste the rows straight into Xero_Raw (keep the same column order). This is the recommended approach if you don’t want to manage OAuth token refresh yourself.


4. META (FACEBOOK ADS) INVOICES — PDF WORKFLOW
Meta invoices are PDFs, not an API feed most agencies set up for this. Recommended workflow each month:
•  Download the PDF invoice from Meta Ads Manager → Billing.
•  For each customer/ad-account line on the PDF, enter one row in Meta_Invoices: Amount Charged (the NET figure Meta shows after any coupon) and Discount/Coupon Applied (shown separately on the PDF). Column I auto-calculates Gross Amount = Charged + Discount.

•  The per-customer monthly summary table (below the detail rows on Meta_Invoices) automatically totals the Gross Amount by customer and month — this is the single number to use for client re-billing or cost analysis, and it already feeds Expense_Summary and the Dashboard.


5. BANK STATEMENT
Export a CSV from your bank portal (Wio Bank or other) and paste into Bank_Raw, columns A–F. The Dashboard’s Cash Balance KPI reads the last value in the Balance column, so keep rows in date order.


6. KEEPING CATEGORIES ACCURATE
Revenue and expense categorisation (WhatsApp / Subscription / Development) is driven entirely by the keyword table on the Assumptions tab. If you introduce a new product line or an account name that doesn’t contain one of the existing keywords, add a new keyword row there (and extend the SEARCH() formula in Stripe_Raw!J and Xero_Raw!J if you add a 4th category) — you do not need to touch any summary formulas.


7. COLOR CODING (used throughout)
Blue text = hardcoded input (paste your real data here). Black text = formula/calculation. Green text = link pulling from another sheet. Yellow fill = a key assumption cell to review before relying on the numbers.

8. UPLOADING BANK STATEMENTS, CREDIT CARDS, CUSTOMER & VENDOR INVOICES (PDF / Word / Excel)
Four dedicated tabs now accept your own source documents directly, without needing Xero or Stripe at all:
•  Bank_Raw — bank statements (PDF or Excel/CSV)
•  Credit_Card_Upload — credit card statements (PDF or Excel/CSV)
•  Customer_Invoices_Upload — customer/AR invoices (PDF or Word)
•  Vendor_Invoices_Upload — vendor/supplier bills (PDF or Word)

IMPORTANT — how the extraction actually works
A plain .xlsx file cannot parse a PDF or Word document by itself — Excel has no native formula that reads PDF/DOCX content, and Word has no Power Query connector at all. There are two reliable ways to get real data into these tabs:

•  RECOMMENDED: Upload the actual PDF/Word/Excel file to Claude in this chat and ask Claude to extract it into the matching tab. Claude reads the document properly (tables, wrapped text, multi-page statements, even messy PDF-to-text artifacts) and pastes the rows in the exact column format each tab expects — this is the same approach used to build the Meta invoice customer-wise breakdown earlier in this workspace.


•  SELF-SERVE ALTERNATIVE: For Excel/CSV bank or credit card exports, use Data → Get Data → From File → From Workbook/CSV in Power Query and load into the tab. For PDFs with clean, consistent tables (recent Excel/365 only), Data → Get Data → From File → From PDF can sometimes pull a table directly — but it is unreliable on multi-page or inconsistently formatted statements (page breaks, wrapped text, merged cells), which is exactly what real bank/Meta/vendor PDFs tend to have. Word documents have no Power Query connector at all, so customer/vendor invoices in .docx format need either the Claude-assisted route above or manual copy-paste.



After the data is in
Once rows land in Credit_Card_Upload, they roll up into Expense_Summary automatically by category (via the Assumptions keyword map). Customer_Invoices_Upload rows roll into Revenue_Summary and Receivables (aging + country split) automatically. Vendor_Invoices_Upload rows roll into Expense_Summary and the new Payables tab (AP aging by vendor) automatically. Log each processed document on the Document_Log tab for an audit trail of what has been entered, by whom, and when.

These new tabs are ADDITIVE to Xero_Raw and Bank_Raw — if an invoice or bill is already recorded in Xero_Raw, do not also enter it in Customer_Invoices_Upload / Vendor_Invoices_Upload, or it will be double-counted.

9. TRIAL BALANCE, IFRS & US GAAP FINANCIAL STATEMENTS
Trial_Balance links every sub-ledger in this workbook (Bank, Receivables, Payables, Revenue_Summary, Expense_Summary, Expenses_Upload, Revenue_Upload) into one Dr/Cr trial balance, with a live Debit = Credit balance check at the bottom.
Financial_Statements_IFRS and Financial_Statements_GAAP both build off Trial_Balance — same numbers, different presentation and terminology (e.g. "Profit for the Period" vs "Net Income", "Statement of Financial Position" vs "Balance Sheet").
IMPORTANT: "Opening Equity" on Trial_Balance is a balancing PLUG (highlighted yellow), not a real number — this sample data has no actual brought-forward trial balance. Replace it with your real opening equity once you have one, and the balance check will confirm everything still ties.
This is a demonstration template, not a substitute for your accountant: real IFRS/GAAP financials also need a fixed asset register, lease schedule, inventory, provisions, and a proper equity roll-forward — none of which exist in this sample dataset. IFRS and US GAAP are largely converged on revenue recognition (IFRS 15 / ASC 606) but differ in areas like development-cost capitalization (IAS 38 vs ASC 350-40/985) and inventory costing (LIFO permitted under US GAAP, not IFRS) — apply your actual accounting policies for a real filing.

10. CUSTOMER RECEIVABLES REPORT (Xero + Stripe, by status)
Customer_Receivables_Report shows every customer’s invoices split by status, for both sources: Xero (Paid / Awaiting Payment / Overdue / Draft / Void / Unsent) and Stripe Invoicing (Paid / Open / Draft / Void / Uncollectible). Stripe_Invoices_Upload is a separate tab from Stripe_Raw — Stripe_Raw holds completed charges, Stripe_Invoices_Upload holds the Invoicing feature’s receivables lifecycle.
Draft and Void invoices are excluded from "outstanding AR" everywhere else in the workbook (Receivables, Dashboard AR KPIs) since they are not real receivables — but they are still shown here for visibility into your full invoice pipeline.

11. 5,000-ROW CAPACITY
Every upload/raw tab (Stripe_Raw, Xero_Raw, Bank_Raw, Meta_Invoices, Credit_Card_Upload, Customer_Invoices_Upload, Vendor_Invoices_Upload, Stripe_Invoices_Upload, Expenses_Upload, Revenue_Upload, Customer_Master) is pre-formulated through row 5003 (5,000 data rows). Category/country lookup formulas are already sitting in every one of those rows — just paste your real data starting at row 4 and everything downstream (summaries, Dashboard, Trial Balance, financial statements) recalculates automatically. Blank rows stay blank; nothing errors out or shows noise until you fill them in.
Column structure: the required columns for each tab (shown in the header row) are fixed, since the formulas need to know where to find dates, amounts, and descriptions. You can safely add EXTRA columns to the right of the existing ones for your own notes without breaking anything — just don’t insert new columns in between the existing ones, since that would shift the formula references. If your export has a genuinely different column layout, paste it into a blank area and ask Claude to remap it into the tab’s expected format.
Receivables, Payables, and Customer_Receivables_Report compute their totals directly from the full 5,000-row ranges (not from a fixed sample-sized list), so aging and status totals stay accurate no matter how many rows you’ve filled in — the "detail" tables you see are illustrative samples, not the calculation source.

12. META COST RECOVERY
Meta_Cost_Recovery compares Meta ad spend (from Meta_Invoices) against two recovery sources: Stripe charges whose description contains "Meta" (edit the description text on your real Stripe charges to include this, or ask Claude to tag them during extraction), and bank transfers tagged to a customer in Bank_Raw column G. Add a customer name in that column on any bank row that represents a Meta cost recovery receipt.

13. SUBSCRIPTION RECURRING INVOICING
Subscription_Invoicing lists each subscription customer with their billing frequency and shows a 12-month grid: blank = not due yet, "⚠ Due" (amber) = should be invoiced this month but no matching Subscription invoice was found yet in Xero_Raw / Customer_Invoices_Upload / Stripe_Invoices_Upload, "✓ Invoiced" (green) = found. Add new subscription customers in the blue rows — pre-formulated through row 200.

14. CUSTOMER MASTER
Customer_Master is the single directory of customer identifiers (Xero contact ID, Stripe customer ID, Meta ad account ID) and metadata (country, industry, relationship owner, payment terms). It doesn’t feed formulas elsewhere automatically — the Assumptions country map is what drives country lookups — but it’s the reference sheet to keep both in sync when you add a new customer.
<img width="991" height="2471" alt="image" src="https://github.com/user-attachments/assets/9b540958-b2d3-489f-8021-89effefa328f" />
