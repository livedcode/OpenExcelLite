using OpenExcelLite.Builders;
using OpenExcelLite.Models;

Console.WriteLine("Generating Excel example files...");

// ---------------------------
// Call all demos here
// ---------------------------
GenerateInMemory();
GenerateInMemoryWithEmptyRows();
GenerateInMemoryWithAfterHeaderEmptyRows();
GenerateInMemoryHyperlinks();
GenerateInMemoryHyperlinksWithEmptyRows();
GenerateInMemoryMultiSheet();
GenerateInMemoryMultiSheetHyperlinks();
GenerateInMemoryMultiSheetWithEmptyRows();
GenerateInMemoryTenSheets();

GenerateStreamingSingleSheet();
GenerateStreamingEntryRows();
GenerateStreamingEntryRowsBetweenRows();
GenerateStreamingHyperlinks();
GenerateStreamingHyperlinksWithEmptyRows();
GenerateStreamingMultiSheet();
GenerateStreamingMultiSheetWithHyperlinks();

// Hybrid (Streaming + InMemory)
GenerateHybridWorkbook();

Console.WriteLine("Done.");




// ============================================================
// 1) In-Memory Excel Demo
// ============================================================
static void GenerateInMemory()
{
    var bytes = new WorkbookBuilder()
        .AddSheet("Employees", s =>
        {
            s.AddRow("Id", "Name", "JoinDate", "Salary", "Active");
            s.AddRow(1, "Alex", DateTime.Today, 5000.5m, true);
            s.AddRow(2, "Brian", DateTime.Today.AddDays(-3), 6500.75m, true);
            s.AddRow(3, "Cindy", DateTime.Today.AddDays(-10), 7200m, false);

            s.AddTable("EmployeesTable");
            s.AutoFitColumns();
        })
        .Build();

    File.WriteAllBytes("InMemory.xlsx", bytes);
}




// ============================================================
// 2) In-Memory Empty Rows (Before Header)
// ============================================================
static void GenerateInMemoryWithEmptyRows()
{
    var bytes = new WorkbookBuilder()
        .AddSheet("Employees", s =>
        {
            s.AddEmptyRows(2);
            s.AddRow("Id", "Name", "JoinDate", "Salary", "Active");
            s.AddRow(1, "Alex", DateTime.Today, 5000.5m, true);
            s.AddRow(2, "Brian", DateTime.Today.AddDays(-3), 6500.75m, true);
            s.AddRow(3, "Cindy", DateTime.Today.AddDays(-10), 7200m, false);

            s.AddTable("EmployeesTable");
            s.AutoFitColumns();
        })
        .Build();

    File.WriteAllBytes("InMemoryEmptyRows.xlsx", bytes);
}




// ============================================================
// 3) In-Memory Empty Rows (After Header)
// ============================================================
static void GenerateInMemoryWithAfterHeaderEmptyRows()
{
    var bytes = new WorkbookBuilder()
        .AddSheet("Employees", s =>
        {
            s.AddRow("Id", "Name", "JoinDate", "Salary", "Active");
            s.AddEmptyRows(2);
            s.AddRow(1, "Alex", DateTime.Today, 5000.5m, true);
            s.AddRow(2, "Brian", DateTime.Today.AddDays(-3), 6500.75m, true);
            s.AddRow(3, "Cindy", DateTime.Today.AddDays(-10), 7200m, false);

            s.AddTable("EmployeesTable");
            s.AutoFitColumns();
        })
        .Build();

    File.WriteAllBytes("InMemoryEmptyRowsAF.xlsx", bytes);
}




// ============================================================
// 4) In-Memory Hyperlinks
// ============================================================
static void GenerateInMemoryHyperlinks()
{
    var bytes = new WorkbookBuilder()
        .AddSheet("Links", s =>
        {
            s.AddRow("Name", "Website");
            s.AddRow("Google", XL.Hyper("https://google.com", "Visit Google"));
            s.AddRow("Repo", XL.Hyper("https://github.com/livedcode/OpenExcelLite"));
        })
        .Build();

    File.WriteAllBytes("InMemoryHyperlinks.xlsx", bytes);
}




// ============================================================
// 5) In-Memory Hyperlinks + Empty Rows
// ============================================================
static void GenerateInMemoryHyperlinksWithEmptyRows()
{
    var bytes = new WorkbookBuilder()
        .AddSheet("Links", s =>
        {
            s.AddEmptyRows(2);
            s.AddRow("Name", "Website");
            s.AddRow("Google", XL.Hyper("https://google.com", "Visit Google"));
            s.AddRow("Repo", XL.Hyper("https://github.com/livedcode/OpenExcelLite"));
        })
        .Build();

    File.WriteAllBytes("InMemoryHyperlinksEmptyRows.xlsx", bytes);
}




// ============================================================
// 6) In-Memory Multi-Sheet
// ============================================================
static void GenerateInMemoryMultiSheet()
{
    var bytes = new WorkbookBuilder()
        .AddSheet("Employees", s =>
        {
            s.AddRow("Id", "Name");
            s.AddRow(1, "Alex");
            s.AddRow(2, "Brian");
        })
        .AddSheet("Departments", s =>
        {
            s.AddRow("DeptId", "Department");
            s.AddRow(10, "Finance");
            s.AddRow(20, "IT");
        })
        .AddSheet("Summary", s =>
        {
            s.AddRow("Generated", DateTime.Now);
        })
        .Build();

    File.WriteAllBytes("InMemoryMultiSheet.xlsx", bytes);
}




// ============================================================
// 7) In-Memory Multi-Sheet Hyperlinks
// ============================================================
static void GenerateInMemoryMultiSheetHyperlinks()
{
    var bytes = new WorkbookBuilder()
        .AddSheet("Links1", s =>
        {
            s.AddRow("Name", "Website");
            s.AddRow("Google", XL.Hyper("https://google.com", "Visit Google"));
        })
        .AddSheet("Links2", s =>
        {
            s.AddRow("API", "URL");
            s.AddRow("Users", XL.Hyper("https://yourapi.com/users", "User API"));
        })
        .AddSheet("Links3", s =>
        {
            s.AddRow("Doc", "URL");
            s.AddRow("README", XL.Hyper("https://github.com/livedcode/OpenExcelLite/blob/main/README.md"));
        })
        .Build();

    File.WriteAllBytes("InMemoryMultiSheetHyperlinks.xlsx", bytes);
}




// ============================================================
// 8) In-Memory Multi-Sheet with Empty Rows
// ============================================================
static void GenerateInMemoryMultiSheetWithEmptyRows()
{
    var bytes = new WorkbookBuilder()
        .AddSheet("A", s =>
        {
            s.AddEmptyRows(3);
            s.AddRow("Id", "Value");
            s.AddRow(1, "AAA");
        })
        .AddSheet("B", s =>
        {
            s.AddRow("Key", "Result");
            s.AddEmptyRows(2);
            s.AddRow("X", 111);
        })
        .AddSheet("C", s =>
        {
            s.AddEmptyRows(5);
            s.AddRow("Title", "Data");
            s.AddRow("Demo", 999);
        })
        .Build();

    File.WriteAllBytes("InMemoryMultiSheetEmptyRows.xlsx", bytes);
}




// ============================================================
// 9) In-Memory 10 Sheets
// ============================================================
static void GenerateInMemoryTenSheets()
{
    var builder = new WorkbookBuilder();

    for (int i = 1; i <= 10; i++)
    {
        builder = builder.AddSheet($"Sheet_{i}", s =>
        {
            s.AddRow("Row", "Value");
            for (int r = 1; r <= 5; r++)
                s.AddRow(r, $"Data {r} in Sheet {i}");
        });
    }

    var bytes = builder.Build();
    File.WriteAllBytes("InMemoryTenSheets.xlsx", bytes);
}




// ============================================================
// 10) Streaming (Single Sheet)
// ============================================================
static void GenerateStreamingSingleSheet()
{
    var bytes = StreamingWorkbookBuilder.Build(wb =>
    {
        wb.AddSheet("BigData", w =>
        {
            w.WriteRow("Id", "Name", "Created");
            for (int i = 1; i <= 50000; i++)
                w.WriteRow(i, "Row " + i, DateTime.UtcNow.AddMinutes(-i));
        });
    });

    File.WriteAllBytes("Streaming.xlsx", bytes);
}




// ============================================================
// 11) Streaming Empty Rows (Before Header)
// ============================================================
static void GenerateStreamingEntryRows()
{
    var bytes = StreamingWorkbookBuilder.Build(wb =>
    {
        wb.AddSheet("BigData", w =>
        {
            w.WriteEmptyRows(2);
            w.WriteRow("Id", "Name", "Created");

            for (int i = 1; i <= 50000; i++)
                w.WriteRow(i, "Row " + i, DateTime.UtcNow.AddMinutes(-i));
        });
    });

    File.WriteAllBytes("StreamingEmptyRows.xlsx", bytes);
}




// ============================================================
// 12) Streaming Empty Rows (Between Rows)
// ============================================================
static void GenerateStreamingEntryRowsBetweenRows()
{
    var bytes = StreamingWorkbookBuilder.Build(wb =>
    {
        wb.AddSheet("BigData", w =>
        {
            w.WriteRow("Id", "Name", "Created");

            for (int i = 1; i <= 50000; i++)
            {
                w.WriteRow(i, "Row " + i, DateTime.UtcNow.AddMinutes(-i));
                w.WriteEmptyRows(1);
            }
        });
    });

    File.WriteAllBytes("StreamingEmptyRowsAF.xlsx", bytes);
}




// ============================================================
// 13) Streaming Hyperlinks
// ============================================================
static void GenerateStreamingHyperlinks()
{
    var bytes = StreamingWorkbookBuilder.Build(wb =>
    {
        wb.AddSheet("Links", w =>
        {
            w.WriteRow("Name", "Website");
            w.WriteRow("Google", XL.Hyper("https://google.com", "Visit"));
            w.WriteRow("Repo", XL.Hyper("https://github.com/livedcode/OpenExcelLite"));
        });
    });

    File.WriteAllBytes("StreamingHyperlinks.xlsx", bytes);
}




// ============================================================
// 14) Streaming Hyperlinks + Empty Rows
// ============================================================
static void GenerateStreamingHyperlinksWithEmptyRows()
{
    var bytes = StreamingWorkbookBuilder.Build(wb =>
    {
        wb.AddSheet("Links", w =>
        {
            w.WriteEmptyRows(3);
            w.WriteRow("Name", "Website");
            w.WriteRow("Google", XL.Hyper("https://google.com", "Visit"));
            w.WriteRow("Docs", XL.Hyper("https://github.com/livedcode/OpenExcelLite", "View Docs"));
        });
    });

    File.WriteAllBytes("StreamingHyperlinksEmptyRows.xlsx", bytes);
}




// ============================================================
// 15) Streaming Multi-Sheet
// ============================================================
static void GenerateStreamingMultiSheet()
{
    var bytes = StreamingWorkbookBuilder.Build(wb =>
    {
        wb.AddSheet("Users", s =>
        {
            s.WriteRow("Id", "Name");
            s.WriteRow(1, "Alex");
            s.WriteRow(2, "Brian");
        });

        wb.AddSheet("Logs", s =>
        {
            s.WriteRow("Timestamp", "Message");
            s.WriteRow(DateTime.Now, "Started");
        });
    });

    File.WriteAllBytes("StreamingMultiSheet.xlsx", bytes);
}




// ============================================================
// 16) Streaming Multi-Sheet with Hyperlinks
// ============================================================
static void GenerateStreamingMultiSheetWithHyperlinks()
{
    var bytes = StreamingWorkbookBuilder.Build(wb =>
    {
        wb.AddSheet("Links1", s =>
        {
            s.WriteRow("Name", "Website");
            s.WriteRow("Google", XL.Hyper("https://google.com", "Visit"));
        });

        wb.AddSheet("Links2", s =>
        {
            s.WriteRow("Item", "URL");
            s.WriteRow("Repo", XL.Hyper("https://github.com/livedcode/OpenExcelLite"));
        });
    });

    File.WriteAllBytes("StreamingMultiSheetHyperlinks.xlsx", bytes);
}



// ============================================================
// 17) Hybrid: Combine Streaming + InMemory
// NOTE: You must implement WorkbookMerger.Merge() separately.
// ============================================================
static void GenerateHybridWorkbook()
{
    // 1) Streaming workbook
    byte[] bigData = StreamingWorkbookBuilder.Build(wb =>
    {
        wb.AddSheet("BigSheet", s =>
        {
            s.WriteRow("Id", "Value");
            for (int i = 1; i <= 50000; i++)
                s.WriteRow(i, "Row " + i);
        });
    });

    // 2) In-memory workbook
    byte[] smallData = new WorkbookBuilder()
        .AddSheet("Summary", s =>
        {
            s.AddRow("GeneratedOn", DateTime.Now);
            s.AddRow("Version", "1.3.0");
        })
        .AddSheet("Info", s =>
        {
            s.AddRow("Key", "Value");
            s.AddRow("Author", "livedcode");
        })
        .Build();

    // 3) Merge into one workbook (placeholder)
    // TODO: Implement WorkbookMerger.Merge(stream, inMemory)
    byte[] merged = bigData; // temporary fallback

    File.WriteAllBytes("HybridWorkbook.xlsx", merged);
}
