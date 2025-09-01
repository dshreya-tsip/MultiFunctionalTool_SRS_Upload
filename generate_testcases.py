import os
import docx
import openpyxl
import requests

# Step 1: Extract text from SRS.docx
def extract_srs_text(doc_path):
    doc = docx.Document(doc_path)
    return "\n".join([para.text for para in doc.paragraphs if para.text])

# Step 2: Generate test cases using Claude API
def generate_testcases(srs_text, component_name):
    CLAUDE_API_KEY = os.getenv("CLAUDE_API_KEY")
    if not CLAUDE_API_KEY:
        raise ValueError("Claude API key not set. Use: export CLAUDE_API_KEY='your_key_here'")

    prompt = f"""
You are a test engineer. Based on the following Software Requirement Specification (SRS), 
generate detailed functional and negative test cases. Ensure each test case has:

1. Test Case ID (format: nf001, nf002, … for negative cases; f001, f002, … for functional cases)
2. Description
3. Preconditions
4. Test Steps
5. Expected Result
6. Actual Result (keep blank)
7. Status (keep blank)

SRS:
{srs_text}
"""

    headers = {
        "x-api-key": CLAUDE_API_KEY,
        "Content-Type": "application/json",
    }

    data = {
        "model": "claude-3-sonnet-20240229",
        "max_tokens": 1000,
        "temperature": 0,
        "messages": [
            {"role": "user", "content": prompt}
        ],
    }

    response = requests.post("https://api.anthropic.com/v1/messages", headers=headers, json=data)
    response.raise_for_status()
    return response.json()["content"][0]["text"]

# Step 3: Save test cases into Excel (with header section preserved)
def save_to_excel(testcases_text, output_path, component_name):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Test Cases"

    # --- Header Section ---
    ws["A1"] = "Build:"
    ws["C1"] = "Date:"
    ws["E1"] = "Target:"
    ws["G1"] = "Component Name:"
    ws["H1"] = component_name  # Insert Component Name here
    ws["A2"] = "MFP Details:"

    # --- Table Header ---
    headers = [
        "Test Case ID", "Description", "Preconditions",
        "Test Steps", "Expected Result", "Actual Result", "Status"
    ]
    ws.append(headers)

    # --- Parse test cases ---
    for line in testcases_text.split("\n"):
        if line.strip():
            parts = [p.strip() for p in line.split("|")]
            if len(parts) >= 7:
                ws.append(parts[:7])

    wb.save(output_path)

# Main function
if __name__ == "__main__":
    srs_path = "SRS.docx"
    output_path = "Generated_TestCases.xlsx"
    component_name = "Authentication Module"  # You can change this dynamically

    print("Extracting SRS...")
    srs_text = extract_srs_text(srs_path)

    print("Generating test cases from Claude...")
    testcases_text = generate_testcases(srs_text, component_name)

    print("Saving to Excel...")
    save_to_excel(testcases_text, output_path, component_name)

    print(f"✅ Test cases saved to {output_path}")
