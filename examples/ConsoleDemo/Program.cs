using OpenExcelLite.Builders;

Console.WriteLine("Generating Excel example files...");

GenerateInMemory();
GenerateStreaming();
GenerateInMemoryWithEmptyRows();
GenerateInMemoryWithAfterHearderEmptyRows();
GenerateStreamingEntryRows();
GenerateStreamingEntryRowsBetweenRows();
Console.WriteLine("Done.");

// ------------------------------------------------------------
// 1) In-memory API demo (small/medium sized datasets)
// ------------------------------------------------------------
static void GenerateInMemory()
{
    var bytes = new WorkbookBuilder()
        .AddSheet("Employees", s =>
        {
            s.AddRow("Id", "Name", "JoinDate", "Salary", "Active");
            s.AddRow(1, "Alex", DateTime.Today, 5000.5m, true);
            s.AddRow(2, "Brian", DateTime.Today.AddDays(-3), 6500.75m, true);
            s.AddRow(3, "Cindy", DateTime.Today.AddDays(-10), 7200m, false);

            s.AddTable("Employees Table 1");
            s.AutoFitColumns();
        })
        .Build();

    File.WriteAllBytes("InMemory.xlsx", bytes);
}

// ------------------------------------------------------------
// 2) Streaming API demo (100k - 1M rows)
// ------------------------------------------------------------
static void GenerateStreaming()
{
    var bytes = StreamingWorkbookBuilder.Build("BigData", w =>
    {
        w.WriteRow("Id", "Name", "Created");
        for (int i = 1; i <= 50000; i++)
        {
            w.WriteRow(i, "Row " + i, DateTime.UtcNow.AddMinutes(-i));
        }
    });

    File.WriteAllBytes("Streaming.xlsx", bytes);
}


// ------------------------------------------------------------
// 3) Empty rows insertion before header demo
// ------------------------------------------------------------
static void GenerateInMemoryWithEmptyRows()
{
    var bytes = new WorkbookBuilder()
        .AddSheet("Employees", s =>
        {
            s.AddEmptyRows(2);  //two empty rows at top
            s.AddRow("Id", "Name", "JoinDate", "Salary", "Active");
            s.AddRow(1, "Alex", DateTime.Today, 5000.5m, true);
            s.AddRow(2, "Brian", DateTime.Today.AddDays(-3), 6500.75m, true);
            s.AddRow(3, "Cindy", DateTime.Today.AddDays(-10), 7200m, false);

            s.AddTable("Employees Table 1");
            s.AutoFitColumns();
        })
        .Build();

    File.WriteAllBytes("InMemoryEmptyRows.xlsx", bytes);
}

// ------------------------------------------------------------
// 4) Empty rows insertion after header demo
// ------------------------------------------------------------
static void GenerateInMemoryWithAfterHearderEmptyRows()
{
    var bytes = new WorkbookBuilder()
        .AddSheet("Employees", s =>
        {
       
            s.AddRow("Id", "Name", "JoinDate", "Salary", "Active");
            s.AddEmptyRows(2);  //two empty rows 
            s.AddRow(1, "Alex", DateTime.Today, 5000.5m, true);
            s.AddRow(2, "Brian", DateTime.Today.AddDays(-3), 6500.75m, true);
            s.AddRow(3, "Cindy", DateTime.Today.AddDays(-10), 7200m, false);

            s.AddTable("Employees Table 1");
            s.AutoFitColumns();
        })
        .Build();

    File.WriteAllBytes("InMemoryEmptyRowsAF.xlsx", bytes);
}


// ------------------------------------------------------------
// 5) Streaming API demo (100k - 1M rows) , Add empty rows before header
// ------------------------------------------------------------
static void GenerateStreamingEntryRows()
{
    var bytes = StreamingWorkbookBuilder.Build("BigData", w =>
    {
        w.WriteEmptyRows(2); // two empty rows before header
        w.WriteRow("Id", "Name", "Created");
        for (int i = 1; i <= 50000; i++)
        {
            w.WriteRow(i, "Row " + i, DateTime.UtcNow.AddMinutes(-i));
        }
    });

    File.WriteAllBytes("StreamingEmptyRows.xlsx", bytes);
}

// ------------------------------------------------------------
// 5) Streaming API demo (100k - 1M rows) , Add empty rows before header
// ------------------------------------------------------------
static void GenerateStreamingEntryRowsBetweenRows()
{
    var bytes = StreamingWorkbookBuilder.Build("BigData", w =>
    {
     
        w.WriteRow("Id", "Name", "Created");
        for (int i = 1; i <= 50000; i++)
        {
            w.WriteRow(i, "Row " + i, DateTime.UtcNow.AddMinutes(-i));
            w.WriteEmptyRows(1); // one empty rows 
        }
    });

    File.WriteAllBytes("StreamingEmptyRowsAF.xlsx", bytes);
}