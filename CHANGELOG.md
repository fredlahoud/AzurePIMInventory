# Changelog

All notable changes to this project are documented here.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased] — `dev`

### Fixed
- **Ancestor-scope filter used prefix matching.** `-like "$scope*"` treated a
  management group such as `mg-prod-legacy` (or any id sharing a prefix) as inside
  `mg-prod`, so assignments from outside the hierarchy could leak into the report.
  Filtering is now exact, case-insensitive and O(n) via a hash set.
- **Future-dated assignments were labelled `Active`** when using
  `-IncludeFutureAssignments`. They now get the status `Pending`.
- **A `Retry-After` header in HTTP-date form crashed the run** (`[int]` cast outside
  the retry `try`). Both RFC 9110 forms are now parsed, with exponential backoff as
  the fallback.
- **An unparseable `endDateTime` was reported as `Permanent`.** It is now reported
  as `Unknown` with a warning, never as a clean row.
- **CSV had no BOM on PowerShell 7**, so Excel garbled non-ASCII display names.
  The CSV is now UTF-8 with BOM on every version.
- `Get-PimInstance` read `-IncludeFutureAssignments` from script scope implicitly;
  it is now an explicit parameter.
- Dates are normalised explicitly to UTC whether ARM JSON yields strings
  (PowerShell 5.1) or `DateTime` objects (PowerShell 7).

### Security
- CSV text columns starting with `=`, `+`, `-` or `@` are prefixed with `'` to
  prevent spreadsheet formula injection (CWE-1236) from user-controlled display names.
  The HTML "export visible rows" button applies the same protection.

### Added
- `PrivilegedRole` column and `-PrivilegedRoleName` parameter.
- **Standing privileged access** card on the dashboard and in the console summary.
- `-PassThru` to emit report rows to the pipeline.
- `-FailOnScopeError` so pipelines and runbooks fail on an incomplete inventory.
- `-ReferenceDateUtc` (offline mode) for reproducible sample output and tests.
- Dashboard: privileged-only filter, visible row count, export of filtered rows to
  CSV, correct sorting of blank expiry/days values, dark-mode support, `Pending`
  and `Unknown` statuses.
- Pester 5 test suite (61 tests) covering paging, retry, hierarchy, status logic,
  scope filtering, CSV/HTML output and an end-to-end offline run.
- GitHub Actions CI: PSScriptAnalyzer, Pester on Windows PowerShell 5.1,
  PowerShell 7 (Windows and Linux), and a sample-report artifact.
- Dependabot for GitHub Actions, PR template, `.gitignore`, analyzer settings.

### Changed
- CSV column order is now fixed explicitly; `PrivilegedRole` is inserted after
  `RoleType`. Consumers that read columns by name are unaffected.
- Header timestamp is shown in UTC and reflects the evaluation instant.
- `SampleOutput/` regenerated with the current column set.
- README documents every column, status, parameter and the branch workflow.
