# XF1 Actuals - Functional Specs

This file is our working spec document. We will update it feature by feature.

## Project Goal
Build an Excel web add-in for financial modeling workflows, starting with actuals consolidation.

## Spec Status
- Overall status: In progress
- Current focus: Functionality 1 finalization
- Coding status: Not started (by design)

## Functionality 1: Consolidate from multiple files
- Name: Consolidate from multiple files
- Problem it solves: User needs to combine data from multiple Excel source files into one consolidated output file.
- Primary user: Financial model user in Excel
- Trigger: User clicks section button `Consolidate from multiple files` in the task pane.

### User Story
As a user, I want to configure multiple source files with mapping ranges and then export one consolidated dataset, so that I can quickly prepare actuals data for modeling and reporting.

### Confirmed Decisions (2026-02-19)
- Output columns are fixed as: `Account, Entity, Department, Date, Value, SourceFile, SourceSheet`.
- Output file type for v1 is `.xlsx`.
- Source file type for v1 is `.xlsx` only.
- User can add multiple file blocks and remove previously added blocks.
- Maximum files allowed in one run: `12`.
- If output file already exists, behavior is `overwrite`.
- Constants are plain text inputs (no quotes) for `Entity` and `Department`.
- `Date` constant uses calendar picker.
- Blank `Entity` and `Department` are kept blank (no default label).
- If `Date`, `Entity`, or `Department` is constant, it repeats for all output rows from that source.
- If more than one of `Entity`, `Department`, `Date` is set to multi-column, run is blocked with validation error.
- Saved configuration should include full source file paths.
- Config usage mode for v1 is personal mode: each user keeps their own source files and config in their own OneDrive.
- Rows such as `Total for ...` are included for now (no filtering in v1).
- If `Value` is blank or non-numeric text, that output row is skipped.
- If `Account` is blank, that output row is skipped.
- For multi-column `Date`, do not convert Excel serials to ISO format.
- Apply trim (`left` and `right` spaces) on `Account`, `Entity`, and `Department`.

### UI Flow (Draft)
1. User clicks `Consolidate from multiple files`.
2. Section expands and shows `File 1` block.
3. User selects `Source file` using browse button.
4. Add-in reads workbook and populates `Source Sheet` dropdown.
5. User enters mappings: `Account`, `Entity`, `Department`, `Date`, `Value`.
6. User clicks `+` to add `File 2` to `File 12`.
7. User may remove any file block before run.
8. User chooses output location/name and runs consolidation.

### Input Rules (Draft)
- Base rule depends on `Value` range shape.
- If `Value` is single column:
- `Account` must be single column with same row count.
- `Entity` can be constant, blank, or single-column range with same row count.
- `Department` can be constant, blank, or single-column range with same row count.
- `Date` can be constant or single-column range with same row count.
- If `Value` has multiple columns:
- `Account` must be single column with same row count as `Value`.
- `Entity` can be constant, blank, single-column range (same row count), or multi-column range (same column count as `Value`).
- `Department` can be constant, blank, single-column range (same row count), or multi-column range (same column count as `Value`).
- `Date` can be constant, single-column range (same row count), or multi-column range (same column count as `Value`).
- Only one of `Entity`, `Department`, `Date` may be multi-column.
- The multi-column field is unpivoted.
- Multi-column field can be any one of `Entity`, `Department`, or `Date`.
- If one field is multi-column, the other two dimension fields must be constant, blank, or single-column (not multi-column).

### Unpivot Rule (Draft)
- Example: `Department` range `D5:G5` with 4 departments.
- Each `Value` column maps to one department label from that range.
- Consolidated output unpivots so each record has one `Value` and one department value.
- Multi-column dimension shape is `1 row x N columns` for v1.

### Worked Example (Confirmed Pattern)
- Source sheet layout:
- `Account` is a single column (example: column `B`).
- `Value` is multi-column (example: columns `C:K`, same row span as `Account`).
- `Department` is the multi-column dimension in one header row (example: `C2:K2`).
- `Entity` is a single constant value for the source.
- `Date` is a single constant date for the source.
- Consolidation behavior:
- For each account row and each value column, create one output row.
- Department value comes from the matching header cell in `C2:K2`.
- Entity and Date constants are repeated on every generated row.

### Additional Valid Patterns (Confirmed)
- Multi-column `Entity`:
- `Value` is multi-column, `Entity` is `1 x N`, and `Department` + `Date` are constants or blanks.
- Multi-column `Date`:
- `Value` is multi-column, `Date` is `1 x N`, and `Entity` + `Department` are constants or blanks.

### Output (Draft)
- Consolidated dataset columns:
- `Account`
- `Entity`
- `Department`
- `Date`
- `Value`
- `SourceFile`
- `SourceSheet`
- Output file is saved to user-chosen location with user-chosen name.

### Validation Rules (Draft)
- Required: source file selected.
- Required: source sheet selected.
- Required: `Account` and `Value` mappings provided.
- Required: mapping dimensions compatible with selected `Value` shape.
- Required: only one multi-column field among `Entity`, `Department`, `Date`.
- Required: file count between `1` and `12`.
- Required: source file extension is `.xlsx`.

### Data Cleanup Rules (Confirmed)
- Trim leading/trailing spaces for `Account`, `Entity`, and `Department`.
- Skip row when trimmed `Account` is blank.
- Skip row when `Value` is blank.
- Skip row when `Value` is non-numeric text.

### Edge Cases (Draft)
- Invalid range text.
- Range outside worksheet bounds.
- Source file missing or moved.
- Source sheet missing or renamed.
- Blank rows in mapped ranges.
- Non-numeric `Value` cells.

### Acceptance Criteria (Draft)
- [ ] User can configure 1 to 12 source files and run consolidation.
- [ ] User can add and remove file blocks.
- [ ] Validation prevents invalid ranges and shape mismatches.
- [ ] Validation prevents more than one multi-column field among `Entity`, `Department`, `Date`.
- [ ] Consolidated output contains fixed schema and correct row-level values.
- [ ] Existing output file is overwritten when same output name/path is used.

### Requested Extension: Save/Load Mapping Config
- Request: User wants to save setup so recurring processes can be rerun quickly.
- Proposal:
- `Save Config` button exports mappings to JSON file (for example `my-setup.xf1config.json`).
- `Load Config` button imports that file and rehydrates all file blocks and mappings.
- Config stores full source file paths to support recurring process reuse.
- Operating mode (v1): personal mode. Each colleague maintains their own config and source files in their own OneDrive.
- Note: Because Office add-ins run in a webview sandbox, file access must be user-driven (import/export), not silent background read/write to arbitrary folders.
- Note: If a saved local path is unavailable on a machine, the add-in should prompt user to reselect that file path for that source block.

### Open Questions for Review
- None currently. Spec is ready for implementation after user approval.

### Technical Notes
- Initial implementation file target: `src/taskpane/taskpane.ts`
- UI updates expected in: `src/taskpane/taskpane.html`
- Styling updates expected in: `src/taskpane/taskpane.css`

## Change Log
- 2026-02-19: Created initial spec template for collaborative feature definition.
- 2026-02-19: Added Functionality 1 draft for "Consolidate from multiple files" based on user requirements.
- 2026-02-19: Revised rules and recorded confirmed decisions from user clarifications.
- 2026-02-19: Finalized core decisions for multi-column dimension behavior, output format, config paths, and overwrite policy.
- 2026-02-19: Finalized row-skip behavior and trim cleanup rules.
- 2026-02-19: Confirmed v1 config operating model as personal OneDrive mode per user.
