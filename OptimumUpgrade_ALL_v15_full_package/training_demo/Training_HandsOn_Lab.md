# Optimum Upgrade – Hands-On Lab Guide

## Objective

In this lab you will:
- Enter a small set of requirements into the training workbook
- Use automation to generate test cases
- Build a verification summary
- Discuss how this scales to real projects

## Before You Start

You should have:
- Windows and Excel installed
- Macros enabled in Excel
- The training workbook available:
  - `training_demo/Training_Exercise_Requirements.xlsx`
  - Or a macro-enabled variant provided by the trainer

## Step 1 – Open the Workbook

1. Open Excel.
2. Use **File → Open** and navigate to the training demo folder.
3. Open `Training_Exercise_Requirements.xlsx`.
4. Go to the `Requirements` sheet.

## Step 2 – Review Existing Requirements

You will see example requirements such as:
- REQ-001 – "The system shall measure wind speed with an accuracy of ±1 m/s."
- REQ-002 – "The system shall operate from -40°C to +55°C ambient temperature."

Notice the columns:
- Req ID
- Requirement Text
- Source
- Level
- Verification Method
- Test Case ID
- Status

## Step 3 – Add Your Own Requirements

Add at least 5–10 new requirements. Make them simple and realistic for your domain:
- Performance (e.g., accuracy, speed)
- Environmental (e.g., temperature, vibration)
- Interfaces (e.g., communication protocol)

Set the Verification Method to something that includes "Test" or "Demonstration" for items
you expect to verify with testing.

## Step 4 – Run GenerateTestCasesFromRequirements

1. Open the Excel Macros dialog.
2. Select `GenerateTestCasesFromRequirements`.
3. Click **Run**.
4. Switch to the `TestCases` sheet and review the rows that were created:

You should see:
- New `TC-###` IDs
- Titles and descriptions derived from the requirements
- References back to the originating requirement IDs

## Step 5 – Run BuildVerificationSummary

1. Run the macro `BuildVerificationSummary`.
2. Open the `VerificationSummary` sheet.
3. Confirm that each requirement with a test has an entry showing:
   - Req ID
   - Requirement text
   - Verification method
   - Test Case ID
   - Status

## Step 6 – Reflect and Discuss

Consider the following questions:
- How would this approach help on a project with 100+ requirements?
- How could your team ensure that all requirements have a verification method and test?
- Which existing documents (e.g., Test Plan, Test Report) could use data exported from this sheet?

## Optional – Export CSV

If available in your environment:
1. Run `ExportVerificationSummaryCSV`.
2. Save the CSV file.
3. Open it in a text editor or Excel and see how it could feed into a test report.
