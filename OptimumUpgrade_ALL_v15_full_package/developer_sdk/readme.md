# Optimum Upgrade – Developer SDK

This SDK allows developers to extend the Optimum Upgrade automation engine.

## Included

- `vba_api_reference.md` – Public macros and helper functions.
- `sheet_structure.md` – Expected sheet layouts for Requirements, TestCases, and VerificationSummary.
- This readme – Overview and extension guidance.

## Extension Philosophy

The automation module is designed to be:

- **Simple:** Plain VBA modules, no external dependencies.
- **Predictable:** Fixed sheet and column conventions.
- **Extensible:** You can add your own macros alongside the core ones.

Typical extension examples:

- Generating additional worksheets from Requirements.
- Building summary dashboards.
- Auto-formatting or exporting data for reports.
