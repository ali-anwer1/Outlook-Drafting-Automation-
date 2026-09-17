from playwright.sync_api import sync_playwright
import pandas as pd
from pathlib import Path
import yaml
from openpyxl import load_workbook
from datetime import datetime

# ============================================================
# INITIALIZATION AND CONFIGURATION SECTION
# ============================================================
# This section initializes the necessary libraries, loads configuration settings from YAML files, and sets up color mappings for specific values in the "Status" and "Priority" columns of the Excel sheet. It also defines a function to build an HTML table with styled cells based on these color mappings.

today = datetime.now()

# Load config and email template from YAML files
BASE_DIR = Path(__file__).resolve().parent.parent

config_path = BASE_DIR / "yaml files" / "config.yaml"
template_path = BASE_DIR / "yaml files" / "template.yaml"

with open(config_path, "r") as f:
    config = yaml.safe_load(f)

with open(template_path, "r") as f:
    template = yaml.safe_load(f)

# Define color mappings for specific values in the "Status" and "Priority" columns
STATUS_COLORS = {
    "Completed": "#C6E0B4",   
    "In Progress": "#B4C6E7",
    "Pending": "#F8CBAD",
}

PRIORITY_COLORS = {
    "Low": "#99FF99",
    "Medium": "#00B0F0",
    "Important": "#FF6600",
    "Urgent": "#FF0000",
}

# Any column not listed here just renders with a plain white background.
COLUMN_COLOR_MAPS = {
    "Status": STATUS_COLORS,
    "Priority": PRIORITY_COLORS,
}

def build_styled_html_table(df, column_color_maps=COLUMN_COLOR_MAPS, include_headers=False):
    # Builds an HTML table where each column in column_color_maps gets its cells
    # colored per that column's own value->color legend. pandas.read_excel() drops all
    # cell formatting, so colors are re-applied here rather than read from the sheet.
    # Set include_headers=False to drop the header row entirely.

    headers = df.columns.tolist()

    header_html = ""
    if include_headers:
        header_cells = "".join(
            f'<th style="background-color:#4472C4;color:#FFFFFF;padding:6px;text-align:center;'
            f'border:1px solid #999;font-family:Calibri,Arial,sans-serif;font-size:11pt;">{h}</th>'
            for h in headers
        )
        header_html = f"<tr>{header_cells}</tr>"

    rows_html = ""
    for _, row in df.iterrows():
        cells = ""
        for col in headers:
            raw_val = row[col]
            is_blank = pd.isna(raw_val) or (isinstance(raw_val, str) and raw_val.strip() == "")

            if is_blank:
                display_val = "N/A"
            elif hasattr(raw_val, "strftime"):
                # If the value is a datetime object, format it as a string
                display_val = raw_val.strftime("%d-%m-%Y")
            elif isinstance(raw_val, str):
                # If the value is a string, replace line breaks with <br> for HTML display
                display_val = raw_val.replace("\r\n", "<br>").replace("\n", "<br>")
            else:
                display_val = raw_val

            color_map = column_color_maps.get(col)
            lookup_key = "" if is_blank else str(raw_val).strip()
            # If the column has a color map, get the corresponding color for the cell value; otherwise, default to white.
            cell_color = color_map.get(lookup_key, "#FFFFFF") if color_map else "#FFFFFF"
            cells += (
                f'<td style="background-color:{cell_color};padding:6px;text-align:center;'
                f'border:1px solid #ccc;font-family:Calibri,Arial,sans-serif;'
                f'font-size:11pt;">{display_val}</td>'
            )
        rows_html += f"<tr>{cells}</tr>"

    return f'<table style="border-collapse:collapse;width:100%;">{header_html}{rows_html}</table>'

def fill_recipients_field(page, label, recipients):
    # Fills the "To" or "Cc" field with the provided list of recipients.
    field = page.get_by_label(label, exact=True)
    field.click()

    for recipient in recipients:
        page.keyboard.press("End")
        page.keyboard.insert_text(recipient)
        page.wait_for_timeout(500)
        page.keyboard.press("Enter")
        page.wait_for_timeout(500)

# ============================================================
# DATA EXTRACTION AND PROCESSING SECTION
# ============================================================
# This section loads the Excel workbook, extracts data from the "Recipients" table, and filters tasks based on today's date. It also builds an HTML table for the email body and prepares the subject and body of the email using the loaded template.

wb = load_workbook(config["excel_path"], data_only=True)
table_name = "Recipients"

# Find the worksheet and table containing the recipients
target_ws = None
target_tbl = None

for ws in wb.worksheets:
    if table_name in ws.tables:
        target_ws = ws
        target_tbl = ws.tables[table_name]
        break

if target_tbl is None:
    raise ValueError(f"Table '{table_name}' not found in workbook")

# Extract data from the table into a list of dictionaries
data_range = target_ws[target_tbl.ref]
rows = [[cell.value for cell in row] for row in data_range]
header, *data = rows

type_idx = header.index("Type")
email_idx = header.index("Email")

# Separate recipients into "To" and "Cc" lists based on the "Type" column
to_recipients = [r[email_idx] for r in data if str(r[type_idx]).strip().lower() == "to"]
cc_recipients = [r[email_idx] for r in data if str(r[type_idx]).strip().lower() == "cc"]

wb.close()

if not to_recipients:
    raise ValueError("No 'To' recipients found in the Recipients table — cannot send email without at least one 'To' recipient.")


# Select tasks for today based on "Start Date" and "Due Date" columns
tasks_df = pd.read_excel(config["excel_path"], sheet_name="Daily_Tasks", header=2)
today_tasks = tasks_df[pd.to_datetime(tasks_df["Start Date"], errors="coerce").dt.date == today.date()]
completed_tasks = tasks_df[pd.to_datetime(tasks_df["Due Date"], format="mixed", errors="coerce").dt.date == today.date()]

# Select tasks based on the "Report_Selection" sheet, which contains a list of task IDs to include in the report
selection_df = pd.read_excel(config["excel_path"], sheet_name="Report_Selection", header=1)
selected_ids = selection_df["No"].dropna().tolist()
selected_tasks = tasks_df[tasks_df["No"].isin(selected_ids)]

# Combine today's tasks, completed tasks, and selected tasks into a single DataFrame, removing duplicates and sorting by the "No" column
report_df = pd.concat([today_tasks,completed_tasks, selected_tasks]).drop_duplicates(subset="No").sort_values(by="No")

# Remove the "Add to Email" column from the report DataFrame, as it's not needed for the email content
EXCLUDED_COLUMNS = ["Add to email"]

report_df = report_df.drop(columns=EXCLUDED_COLUMNS, errors="ignore")

# Prepare the email subject and body using the loaded template, and build an HTML table for the email body
# Modify the script here if you intend to include additional dynamic content in the email body, such as a summary of tasks or other relevant information.
subject = template["subject"].format(user_name=config["user_name"],date_long=today.strftime("%d %B %Y"))

body = template["body"].format(day_name=today.strftime("%A"),date_short=today.strftime("%d/%m/%Y"))

html_table = build_styled_html_table(report_df)

body_html = body.replace("\n", "<br>")
final_html_body = (
    f'<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;">{body_html}</div>'
    f"<br>{html_table}"
)

# ============================================================
# EMAIL COMPOSITION AND SENDING SECTION
# ============================================================
# This section uses Playwright to automate the process of composing and sending an email in Outlook. It launches the Edge browser, loads the saved session state, navigates to the Outlook mail page, fills in the "To" and "Cc" fields with the extracted recipients, sets the subject and body of the email, and applies the user's signature.
with sync_playwright() as p:
    try:
        browser = p.chromium.launch(
            channel="msedge",
            headless=False
        )

        context = browser.new_context(
            storage_state="outlook_session.json"
        )

        page = context.new_page()

        page.goto("https://outlook.office.com/mail")

        # Click the "New" button to start composing a new email
        page.get_by_role("button", name="New", exact=True).click()

        # Fill in the "To" and "Cc" fields with the extracted recipients
        fill_recipients_field(page, "To", to_recipients)
        fill_recipients_field(page, "Cc", cc_recipients)

        # Set the subject of the email
        page.get_by_placeholder("Add a subject").fill(subject)

        # Set the body of the email using a custom paste method to preserve HTML formatting
        message_body = page.get_by_label("Message body")
        message_body.click()
        message_body.evaluate(
            """
            (el, [html, plainFallback]) => {
                const dt = new DataTransfer();
                dt.setData('text/html', html);
                dt.setData('text/plain', plainFallback);
                const event = new ClipboardEvent('paste', {
                    clipboardData: dt,
                    bubbles: true,
                    cancelable: true
                });
                el.dispatchEvent(event);
            }
            """,
            [final_html_body, body],
        )

        # Ensure the body is properly set before applying the signature
        # page.keyboard.press("Enter")  

        # Apply the user's signature by clicking the "Signature" button and selecting the appropriate signature from the menu
        # page.get_by_label("Signature").click()
        # page.get_by_role("menuitem", name=config["signature"]).click()

        # Wait for the user to review the email in the browser and close the window when user closes the browser window
        print("Review the email in the browser. Close the window when you're done.")
        try:
            page.wait_for_event("close", timeout=0)

        finally:
            print("Browser window closed. Exiting script.")
            browser.close()
   
    finally:
        pass
