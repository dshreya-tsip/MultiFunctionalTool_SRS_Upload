import os
import docx
import openpyxl
import requests

# Step 1: Extract text from SRS.docx
def extract_srs_text(doc_path):
    doc = docx.Document(doc_path)
    return "\n".join([para.text for para in doc.paragraphs if para.text.strip()])

# Step 2: Send prompt to Claude API
def get_testcases_from_claude(srs_text):
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        raise ValueError("Missing Anthropic API key. Set ANTHROPIC_API_KEY environment variable.")

    prompt = (
        "Read the uploaded Software Requirements Specification (SRS.docx) and generate both "
        "functional and non-functional test cases. Populate the results into the provided "
        "TestCases_Template.xlsx document. Functional test cases should cover all described features, "
        "while non-functional test cases should address performance, usability, and compatibility. "
        "Return the test cases in markdown table format with columns: "
        "`Test Case ID`, `Preconditions`, `Test Condition`, `Steps with description`, `Expected Result`, `Actual Result`, `Remarks`.\n\n"
        "SRS Content:\n" + srs_text
    )

    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }

    payload = {
        "model": "claude-3-7-sonnet-20250219",
        "max_tokens": 1000,
        "temperature": 0.3,
        "messages": [
            {
                "role": "user",
                "content": prompt
            }
        ]
    }

    response = requests.post("https://api.anthropic.com/v1/messages", json=payload, headers=headers)
    response.raise_for_status()
    return response.json()["content"]

# Step 3: Parse markdown table into structured test cases
def parse_markdown_table(md_text):
    if isinstance(md_text, list):
        md_text = "\n".join(md_text)
    lines = [line for line in md_text.splitlines() if "|" in line]

    headers = [h.strip() for h in lines[0].split("|")[1:-1]]
    test_cases = []

    for line in lines[2:]:  # skip header and separator
        values = [v.strip() for v in line.split("|")[1:-1]]
        test_case = dict(zip(headers, values))
        test_cases.append(test_case)

    return test_cases

# Step 4: Fill test cases into Excel template
def fill_excel_template(test_cases, template_path, output_path):
    wb = openpyxl.load_workbook(template_path)
    ws = wb["Testcases"]
    start_row = 6

    for i, tc in enumerate(test_cases):
        row = start_row + i
        ws.cell(row=row, column=2, value=tc.get("Test Case ID"))
        ws.cell(row=row, column=3, value=tc.get("Preconditions"))
        ws.cell(row=row, column=4, value=tc.get("Test Condition"))
        ws.cell(row=row, column=5, value=tc.get("Steps with description"))
        ws.cell(row=row, column=6, value=tc.get("Expected Result"))
        ws.cell(row=row, column=7, value=tc.get("Actual Result"))
        ws.cell(row=row, column=8, value=tc.get("Remarks"))

    wb.save(output_path)

# Main execution
srs_text = extract_srs_text("SRS.docx")
md_testcases = get_testcases_from_claude(srs_text)
test_cases = parse_markdown_table(md_testcases)
fill_excel_template(test_cases, "TestCases_Template.xlsx", "Generated_TestCases.xlsx")
