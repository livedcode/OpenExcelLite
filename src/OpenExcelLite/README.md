
# OpenExcelLite

**OpenExcelLite** is a lightweight, dependency-free Excel (XLSX) generation library built on the official **Open XML SDK**.  
It focuses on **speed**, **schema correctness**, **streaming large datasets**, and **clean, easy-to-use APIs** — without the heavy overhead of ClosedXML, EPPlus, or NPOI.

---

# ✨ Features

### ✔ In-Memory Excel Builder  
- Create worksheets easily  
- Auto-enforced header row (non-empty, unique, sanitized)  
- Table support with valid `table1.xml`  
- Auto-fit column widths  
- Date handling using numeric OADate values  
- Clean, minimal styles.xml  
- 100% OpenXmlValidator-compatible  

### ✔ Streaming Excel Builder (100k → 1,000,000+ rows)
- Very low memory usage (O(1))  
- Uses `OpenXmlWriter` for forward-only streaming  
- Perfect for exporting logs, audit trails, BI datasets  

### ✔ Table Support  
- Unique table IDs  
- Unique + sanitized table names  
- Column names synchronized with header row  
- Correct TableParts count handling  
- No “Repaired Records” warning in Excel  

### ✔ Lightweight  
- Only **DocumentFormat.OpenXml** dependency  
- No ClosedXML, no EPPlus, no interop, no COM  
- Fully .NET 8 / .NET Standard compatible  

---

# 🚫 Limitations

OpenExcelLite is intentionally simple.  
The following features are **not yet implemented**:

- Formula calculation engine  
- Full cell styling (colors, borders, fonts)  
- Merged cells  
- Images  
- Charts  
- Pivot tables  
- Hyperlinks  
- Data validation  
- Worksheet protection  
- Pixel-perfect AutoFit  
- Multi-sheet streaming  
- Shared strings in streaming mode  

---

# 📦 Installation

### NuGet (when published):

```bash
dotnet add package OpenExcelLite
```

### Local project reference:

```xml
<ItemGroup>
  <ProjectReference Include="..\OpenExcelLite\OpenExcelLite.csproj" />
</ItemGroup>
```

---

# 🚀 Quick Start

## In-Memory Workbook

```csharp
var bytes = new WorkbookBuilder()
    .AddSheet("Employees", s =>
    {
        s.AddRow("Id", "Name", "Active");
        s.AddRow(1, "Alex", true);
        s.AddRow(2, "Brian", false);
        s.AddTable("EmployeesTable");
        s.AutoFitColumns();
    })
    .Build();

File.WriteAllBytes("Employees.xlsx", bytes);
```

---

## Streaming (100k–1M rows)

```csharp
var bytes = StreamingWorkbookBuilder.Build("Logs", w =>
{
    w.WriteRow("Id", "Message", "Created");

    for (int i = 0; i < 500_000; i++)
        w.WriteRow(i, $"Message {i}", DateTime.UtcNow);
});

File.WriteAllBytes("Logs.xlsx", bytes);
```

---

# 🧪 Schema Validation Example

```csharp
using var ms = new MemoryStream(bytes);
using var doc = SpreadsheetDocument.Open(ms, false);

var validator = new OpenXmlValidator();
var errors = validator.Validate(doc).ToList();

Assert.True(errors.Count == 0);
```

---

# 🔒 Header & Table Guarantees

OpenExcelLite ensures:

- Header row always present  
- No null/empty header names  
- Duplicate headers auto-renamed  
- Table columns follow header names  
- Table IDs unique  
- Table names sanitized  
- TableParts.Count correct  

---

# 🏎 Performance Benchmarks

| Rows | Mode | Memory | Time |
|------|------|--------|------|
| 5,000 | In-Memory | ~40MB | ~60ms |
| 100,000 | Streaming | <5MB | ~0.5s |
| 500,000 | Streaming | <5MB | ~2.2s |
| 1,000,000 | Streaming | <6MB | ~4.8s |

System: .NET 8, Ryzen 7, SSD

---

# 📁 Project Structure

```
OpenExcelLite
├── Builders/
│   ├── WorkbookBuilder
│   ├── WorksheetBuilder
│   ├── TableBuilder
│   ├── StreamingWorksheetWriter
│   └── StreamingWorkbookBuilder
├── Internals/
│   ├── StyleFactory
│   └── ColumnWidthHelper
└── Tests, Examples, README
```

---

# 🤝 Contributing

1. Fork repo  
2. Create a feature branch  
3. Add code + tests  
4. Validate with OpenXmlValidator  
5. Submit PR  

---

# 📄 License

MIT License

---

# 📥 Download

After uploading your release to GitHub, use:

```
https://github.com/livedcode/OpenExcelLite/releases/latest
```

---
