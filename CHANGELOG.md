# CHANGELOG

## v1.1.0 — Improved Row Handling & Streaming Enhancements

### ✨ Added
- `WorksheetBuilder.AddEmptyRows(int count)` — safely inserts schema-valid blank rows.
- `StreamingWorksheetWriter.WriteEmptyRows(int count)` — streaming blank-row support.

### 🛠 Improved
- Header detection now uses the first non-empty row.
- `_headerColumnCount` and `_headerRowIndex` tracked properly.
- AutoFilter range now uses actual header row.
- Table ranges fixed to avoid Excel repair warnings.

### 🧪 Tests
- `InMemory_WithEmptyRowsBeforeHeader_ShouldBeSchemaValid`
- `Streaming_EmptyRows_ShouldBeSchemaValid`

### 🪲 Fixed
- Excel “Repaired Records: Table…” warnings.
- Column-count mismatch when blank rows preceded header.
- Streaming blank rows previously generated invalid XML.
