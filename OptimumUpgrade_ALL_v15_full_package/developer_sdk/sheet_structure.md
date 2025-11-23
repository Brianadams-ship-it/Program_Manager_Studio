# Sheet Structure Specifications

To use the automation reliably, keep these sheet structures.

## Requirements Sheet

Expected columns (minimally):

- Column A: Req ID
- Column B: Requirement Text
- Column E: Verification Method
- Column F: Test Case ID
- Column G: Status

Other columns may exist, but the macros assume these positions.

## TestCases Sheet

Suggested columns:

- Column A: Test Case ID
- Column B: Title
- Column C: Related Req ID(s)
- Column D: Description
- Column E: Type
- Column F: Configuration
- Column G: Steps
- Column H: Expected Result
- Column I: Status

## VerificationSummary Sheet

Created by the automation:

- Column A: Req ID
- Column B: Requirement
- Column C: Verification Method
- Column D: Test Case ID
- Column E: Status
