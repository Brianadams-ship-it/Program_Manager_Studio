# VBA Automation API Reference

## Public Macros

### GenerateTestCasesFromRequirements()

Scans the `Requirements` sheet and, for requirements whose Verification Method includes "test" or "demonstration" and have no Test Case ID assigned, it:

- Generates the next `TC-###` ID.
- Writes a row into the `TestCases` sheet with a default title, description, and placeholders.

### BuildVerificationSummary()

Creates or refreshes the `VerificationSummary` sheet and copies:

- Req ID
- Requirement text
- Verification Method
- Test Case ID
- Status

from the `Requirements` sheet.

### ExportVerificationSummaryCSV()

Exports the `VerificationSummary` sheet to a user-selected CSV file, suitable for importing into Word or other reporting tools.

### EnsureProjectInfoSheet()

Creates a `ProjectInfo` sheet if it does not exist, with standard fields like:

- ProjectName
- ProjectAcronym
- Customer
- ContractNumber
- ProgramID
- PMName
- SELead
- StartDate
- EndDate

### UpdateSheetHeadersFromProjectInfo()

Reads ProjectInfo values and updates sheet headers (cells A1 and A2) across the workbook, excluding the ProjectInfo sheet itself.

## Helper Functions

### NextTestCaseID(ws As Worksheet) As String

Finds the highest existing `TC-###` in the specified worksheet and returns the next sequential ID.

### GetProjectField(wsInfo As Worksheet, fieldName As String) As String

Looks up a row in the `ProjectInfo` sheet where Column A matches `fieldName` and returns the corresponding value from Column B.
