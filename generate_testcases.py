import docx
import openpyxl

# Step 1: Extract text from SRS.docx
def extract_srs_text(doc_path):
    doc = docx.Document(doc_path)
    return "\n".join([para.text for para in doc.paragraphs if para.text.strip()])

# Step 2: Generate test cases using Claude-style prompt logic
def generate_test_cases(srs_text):
    # Claude-style prompt embedded in function
    prompt = (
        "Read the uploaded Software Requirements Specification (SRS.docx) and generate both "
        "functional and non-functional test cases. Populate the results into the provided "
        "TestCases_Template.xlsx document. Functional test cases should cover all described features, "
        "while non-functional test cases should address performance, usability, and compatibility. "
        "Return the completed Excel file for review and integration.\n\n"
        "SRS Content:\n" + srs_text
    )
    print("Prompt to Claude:\n", prompt)  # Optional: log the prompt

    # Simulated parsing of SRS to extract functional features
    functional_features = []
    current_title = ""
    for line in srs_text.splitlines():
        line = line.strip()
        if line.startswith("4.") and len(line.split()) > 1:
            current_title = line
        elif current_title and line:
            functional_features.append((current_title, line))
            current_title = ""

    # Non-functional features
    non_functional_features = [
        ("Performance", "Measure time taken to complete all diagnostic operations", "Operations complete within acceptable time limits"),
        ("Usability", "Verify GUI layout and ease of use for all features", "User interface is intuitive and easy to navigate"),
        ("Compatibility", "Run tool on both desktop and laptop environments", "Tool functions correctly on all supported platforms")
    ]

    # Generate test cases
    test_cases = []
    for i, (title, description) in enumerate(functional_features, start=1):
        test_cases.append({
            "Test Case ID": f"TC_FUNC_{i:03}",
            "Preconditions": "Tool installed and accessible",
            "Test Condition": title,
            "Steps with description": f"Verify functionality: {description}",
            "Expected Result": f"{title} works as expected",
            "Actual Result": "Not Executed",
            "Remarks": ""
        })

    for j, (condition, steps, expected) in enumerate(non_functional_features, start=1):
        test_cases.append({
            "Test Case ID": f"TC_NONFUNC_{j:03}",
            "Preconditions": "Tool installed and running",
            "Test Condition": condition,
            "Steps with description": steps,
            "Expected Result": expected,
            "Actual Result": "Not Executed",
            "Remarks": ""
        })

    return test_cases

# Step 3: Fill test cases into Excel template
def fill_excel_template(test_cases, template_path, output_path):
    wb = openpyxl.load_workbook(template_path)
    ws = wb["Testcases"]
    start_row = 6

    for i, tc in enumerate(test_cases):
        row = start_row + i
        ws.cell(row=row, column=2, value=tc["Test Case ID"])
        ws.cell(row=row, column=3, value=tc["Preconditions"])
        ws.cell(row=row, column=4, value=tc["Test Condition"])
        ws.cell(row=row, column=5, value=tc["Steps with description"])
        ws.cell(row=row, column=6, value=tc["Expected Result"])
        ws.cell(row=row, column=7, value=tc["Actual Result"])
        ws.cell(row=row, column=8, value=tc["Remarks"])

    wb.save(output_path)

# Main execution
srs_text = extract_srs_text("SRS.docx")
test_cases = generate_test_cases(srs_text)
fill_excel_template(test_cases, "TestCases_Template.xlsx", "Generated_TestCases.xlsx")
