# Changelog
All notable changes to **OpenExcelLite** will be documented in this file.

The format follows **[Keep a Changelog](https://keepachangelog.com/en/1.1.0/)**  
and adheres to **[Semantic Versioning](https://semver.org/spec/v2.0.0.html)**.

---

## [Unreleased]
### Added
- (placeholder for future improvements)

---

## [1.2.0] - 2025-11-23
### Added
- Full **hyperlink support** for both in-memory and streaming modes.
- New model: `HyperlinkCell` and helper `XL.Hyper(url, displayText)`.
- Proper `<hyperlinks>` element generation with external relationships.
- Unit tests for in-memory and streaming hyperlink scenarios.
- New examples in README and Program.cs showing hyperlink usage.

### Improved
- Consistency in hyperlink schema generation (ECMA-376 compliant).
- README updated with hyperlink feature documentation.
- Streaming writer now tracks hyperlinks without memory explosion.

### Fixed
- Ensured hyperlinks do not trigger Excel “Repaired Records” dialogs.
- Fixed edge case: hyperlinks combined with empty rows and tables.

---

## [1.1.1] - 2025-11-22
### Changed
- Metadata-only update for NuGet.
- Improved SEO-optimized package tags.
- Enhanced package description for discoverability.
- Updated README & CHANGELOG formatting.
- No functional code changes.

---

## [1.1.0] - 2025-11-21
### Added
- `AddEmptyRows(int)` for in-memory writer.
- `WriteEmptyRows(int)` for streaming writer.

### Improved
- Header row detection logic.
- AutoFilter and table range calculation.
- Column-count consistency validation.
- AutoFit column logic for datasets with leading empty rows.

### Fixed
- Eliminated Excel “Repaired Records: Table…” warnings.
- Corrected handling of blank rows before header in both modes.
- Streaming empty row writer produced invalid XML — now fixed.

---

## [1.0.0] - 2025-11-18
### Added
- Initial release of **OpenExcelLite**.
- In-memory and streaming Excel creators.
- AutoFit columns, AutoFilter, table creation.
- Date style handling.
- Schema-safe OpenXML generation.

---

## Links
- **Released versions:** https://github.com/livedcode/OpenExcelLite/releases
