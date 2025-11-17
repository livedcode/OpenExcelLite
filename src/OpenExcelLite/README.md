# OpenExcelLite

A lightweight, schema-safe Excel (XLSX) generator for .NET using the official OpenXML SDK.  
Supports both in-memory and streaming Excel creation — designed for fast, dependency-free exports.

---

## ✨ Features

- In-memory Excel builder  
- Streaming XLSX writer for large datasets (100k–1M rows)  
- Table creation with styling  
- AutoFilter support  
- AutoFit column widths (approx algorithm)  
- Date formatting (OADate + style index)  
- Automatic header validation and deduplication  
- Boolean, numeric, string inference  
- **NEW in v1.1.0**: Blank row support (in-memory + streaming)

---

## 🚀 New in v1.1.0 — Blank Row Enhancements

### ✔ AddEmptyRows() — In-Memory Builder

```csharp
s.AddEmptyRows(3);
s.AddRow("Id", "Name");
s.AddRow(1, "Alex");
```

### ✔ Streaming: WriteEmptyRows()

```csharp
writer.WriteEmptyRows(5);
writer.WriteRow("Id", "Name");
```

### ✔ Improved Stability

- Table ranges compute correct header-row offset  
- AutoFilter respects actual header row  
- Eliminates Excel “Repaired Records” warnings  
- Fully compliant with ECMA-376 schema

---

## 📄 Example (In-Memory)

```csharp
var bytes = new WorkbookBuilder()
    .AddSheet("Demo", s =>
    {
        s.AddEmptyRows(2);
        s.AddRow("Id", "Name", "Active");
        s.AddRow(1, "Alex", true);
        s.AddRow(2, "Brian", false);
        s.AddTable("Employees");
    })
    .Build();

File.WriteAllBytes("demo.xlsx", bytes);
```

---

## 📄 Example (Streaming)

```csharp
var bytes = StreamingWorkbookBuilder.Build("Demo", writer =>
{
    writer.WriteEmptyRows(4);
    writer.WriteRow("Id", "Name");
    writer.WriteRow(1, "Alex");
});
```

---

## 📜 License

MIT License (included in package)
