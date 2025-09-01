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
    "and once completed, continue numbering with non-functional test cases (performance, usability, "
    "and compatibility) without adding any new section headers or titles. "
    "All test cases must be placed in a single continuous markdown table with sequential numbering "
    "for `Test Case ID` (e.g., TC001, TC002, ...). "
    "Before the table, fill in the sheet header fields with suitable values: "
    "`Component` (choose the most appropriate component/module name based on the SRS), "
    "`MFP` (use 'Any' if not specified), `Build` (leave blank), `Date` (leave blank), "
    "and `Target` (leave blank). "
    "Then return the test cases in markdown table format with columns: "
    "`Test Case ID`, `Preconditions`, `Test Condition`, `Steps with description`, "
    "`Expected Result`, `Actual Result`, `Remarks`.\n\n"
    "SRS Content:\n" + srs_text
)



    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }

    payload = {
        "model": "claude-3-7-sonnet-20250219",
        "max_tokens": 1500,
        "temperature": 0.3,
        "messages": [
            {"role": "user", "content": prompt}
        ]
    }

    response = requests.post("https://api.anthropic.com/v1/messages", json=payload, headers=headers)
    response.raise_for_status()

    result = response.json()

    # ✅ Extract only text blocks from Claude response
    md_text = "\n".join(
        block["text"] for block in result["content"] if block["type"] == "text"
    )

    # Debug: print raw output (can be removed later)
    print("\n--- Claude Response (Markdown Table) ---\n")
    print(md_text[:1000])  # show first 1000 chars for safety
    print("\n---------------------------------------\n")

    return md_text

# Step 3: Parse markdown table into structured test cases
def parse_markdown_table(md_text):
    if isinstance(md_text, list):
        if all(isinstance(item, dict) for item in md_text):
            return md_text
        else:
            md_text = "\n".join(str(item) for item in md_text)
    elif not isinstance(md_text, str):
        raise TypeError(f"Expected md_text to be a string or list of dicts, got {type(md_text)}")

    lines = [line for line in md_text.splitlines() if "|" in line]
    if len(lines) < 3:
        raise ValueError("Markdown table format is invalid or incomplete.")

    headers = [h.strip() for h in lines[0].split("|")[1:-1]]
    test_cases = []

    for line in lines[2:]:  # skip header + separator
        values = [v.strip() for v in line.split("|")[1:-1]]
        if len(values) == len(headers):
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
if __name__ == "__main__":
    srs_text = extract_srs_text("SRS.docx")
    md_testcases = get_testcases_from_claude(srs_text)
    test_cases = parse_markdown_table(md_testcases)
    fill_excel_template(test_cases, "TestCases_Template.xlsx", "Generated_TestCases.xlsx")
    print("✅ Test cases generated successfully: Generated_TestCases.xlsx")
