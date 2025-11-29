# OpenExcelLite

[![NuGet Version](https://img.shields.io/nuget/v/OpenExcelLite.svg?label=NuGet&color=3982CE)](https://www.nuget.org/packages/OpenExcelLite)
[![NuGet Downloads](https://img.shields.io/nuget/dt/OpenExcelLite.svg?color=blue)](https://www.nuget.org/packages/OpenExcelLite)
![License](https://img.shields.io/badge/License-MIT-green.svg)

OpenExcelLite is a lightweight, schema-safe Excel (XLSX) generator for .NET using the official OpenXML SDK.  
Supports **in-memory**, **streaming**, **multi-sheet**, **hyperlinks**, and **empty row** generation — with zero dependencies.

---

# ✨ Features

- In-memory Excel builder  
- Streaming writer for 100k–1M+ rows  
- Multi-sheet (in-memory + streaming)  
- Hyperlinks (in-memory + streaming)  
- Blank row support  
- AutoFilter & table creation  
- AutoFit columns (approx)  
- Header validation & range correction  
- Fast, dependency-free, ECMA-376 valid  

---

# 🚀 New in v1.3.0  
### Multi-sheet streaming + hybrid mode

```csharp
var bytes = StreamingWorkbookBuilder.Build(wb =>
{
    wb.AddSheet("Users", s => {
        s.WriteRow("Id", "Name");
        s.WriteRow(1, "Alex");
    });

    wb.AddSheet("Logs", s => {
        s.WriteRow("Timestamp", "Message");
        s.WriteRow(DateTime.Now, "Started");
    });
});

---

# 🚀 New in v1.2.0 — Hyperlink Support

OpenExcelLite now supports clickable **Excel hyperlinks** with:

- Custom display text  
- Full ECMA-376 compliant `<hyperlinks>` + relationship parts  
- Works in both in-memory and streaming modes  
- No Excel repair warnings  
- Fully schema-valid output  

### ✔ Create a hyperlink

```csharp
s.AddRow("Name", "Website");
s.AddRow("Google", XL.Hyper("https://google.com", "Visit Google"));
```

### ✔ Streaming hyperlinks

```csharp
var bytes = StreamingWorkbookBuilder.Build("Links", w =>
{
    w.WriteRow("Name", "Website");
    w.WriteRow("GitHub", XL.Hyper("https://github.com/livedcode/OpenExcelLite"));
});
```

---

# 🚀 New in v1.1.0 — Blank Row Enhancements

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

---

# 📄 Example (In-Memory)

```csharp
var bytes = new WorkbookBuilder()
    .AddSheet("Demo", s =>
    {
        s.AddEmptyRows(2);
        s.AddRow("Id", "Name", "Active");
        s.AddRow(1, "Alex", true);
        s.AddRow(2, "Brian", false);
        s.AddTable("Employees");
        s.AutoFitColumns();
    })
    .Build();

File.WriteAllBytes("demo.xlsx", bytes);
```

---

# 📄 Example (In-Memory Hyperlinks)

```csharp
var bytes = new WorkbookBuilder()
    .AddSheet("Links", s =>
    {
        s.AddRow("Name", "Website");
        s.AddRow("Google", XL.Hyper("https://google.com", "Visit Google"));
        s.AddRow("GitHub", XL.Hyper("https://github.com/livedcode/OpenExcelLite"));
    })
    .Build();

File.WriteAllBytes("hyperlinks.xlsx", bytes);
```

---

# 📄 Example (Streaming)

```csharp
var bytes = StreamingWorkbookBuilder.Build("Demo", writer =>
{
    writer.WriteEmptyRows(4);
    writer.WriteRow("Id", "Name");
    writer.WriteRow(1, "Alex");
});
```

---

# 📄 Example (Streaming Hyperlinks)

```csharp
var bytes = StreamingWorkbookBuilder.Build("Links", writer =>
{
    writer.WriteRow("Name", "Website");
    writer.WriteRow("Google", XL.Hyper("https://google.com", "Visit"));
    writer.WriteRow("GitHub", XL.Hyper("https://github.com/livedcode/OpenExcelLite"));
});

File.WriteAllBytes("streaming_links.xlsx", bytes);
```

---

# 📌 Hyperlink Behavior

- Display text stored in the cell  
- URL stored in hyperlink relationship (`.rels`)  
- Excel renders as a standard clickable hyperlink  
- Works in both in-memory & streaming modes  
- Fully valid when checked with OpenXML Validator  
- No Excel “Repaired Records” alerts  

---

# 📜 License

MIT License (included in package)
