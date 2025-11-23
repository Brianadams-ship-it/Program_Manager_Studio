# Optimum Upgrade – Facilitator Script

This script provides talking points and guidance for trainers delivering the
"Getting Started with the Optimum Upgrade Automation Suite" session.

## 1. Welcome and Context

- Introduce yourself and the purpose of the session.
- Explain that many teams struggle with either:
  - Heavy tools that nobody likes using, or
  - Unstructured spreadsheets and documents.

Suggested line:
> "Optimum Upgrade is meant to give us structure and automation while still using tools we already know: Excel and Word."

## 2. Tour of the Package

1. Show the extracted Optimum Upgrade folder.
2. Highlight these key directories:
   - `pages/` – HTML documentation and product pages
   - `automation/` – automation module and installer script
   - `developer_sdk/` – docs for extending the automation
   - `templates/` – starter toolkits and Word-like templates
   - `sample_project/` – FalconEye example program
   - `training_demo/` – training materials for this session
3. Open `Product_Catalog_OptimumUpgrade.pdf` and briefly explain each product.

## 3. Automation Concepts

- Use a whiteboard or slide:
  - Requirements → TestCases → VerificationSummary
- Emphasize that Requirements are the primary source of truth.
- Explain that the automation helps:
  - Generate test cases
  - Build verification summaries
  - Keep project metadata aligned through ProjectInfo

## 4. Live Automation Demo

Walk through the steps slowly:

1. Open `templates/Automation_Base_Toolkit.xlsx` or the macro-enabled version.
2. Confirm that the automation module is present (or run the installer beforehand).
3. Show the macros available in Excel.
4. Run `EnsureProjectInfoSheet` and fill in a demo project (e.g., "FalconEye Demo").
5. Enter a few realistic requirements in the Requirements sheet.
6. Run `GenerateTestCasesFromRequirements` and show the new rows in the TestCases sheet.
7. Run `BuildVerificationSummary` and walk through the summary contents.
8. Optionally run `ExportVerificationSummaryCSV` to show how data can feed reports.

Narrate what you're doing and why each step matters in a real program.

## 5. FalconEye Sample Project

- Open the FalconEye sample project files.
- Show how:
  - Requirements relate to test cases
  - Risk registers and other artifacts complement the automation flow
- Explain that FalconEye is a learning and reference model, not a real system.

## 6. Hands-On Lab

- Direct participants to open `training_demo/Training_Exercise_Requirements.xlsx`.
- Ask them to:
  - Add 5–10 simple requirements
  - Run `GenerateTestCasesFromRequirements`
  - Run `BuildVerificationSummary`
- Encourage questions, circulate, and help participants who get stuck.

## 7. Wrap-Up and Next Steps

- Ask for one or two quick takeaways from the group.
- Suggest a pilot project where they can apply the toolkit.
- Mention where documentation and internal support will live (e.g., a wiki or shared drive).
