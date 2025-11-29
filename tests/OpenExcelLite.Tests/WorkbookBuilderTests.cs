using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OpenExcelLite.Builders;
using OpenExcelLite.Models;

namespace OpenExcelLite.Tests;

public class WorkbookBuilderTests
{
    // ============================================================
    // Helpers
    // ============================================================
    private static void AssertSchemaValid(byte[] bytes)
    {
        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 0);

        using var ms = new MemoryStream(bytes);
        using var doc = SpreadsheetDocument.Open(ms, false);

        var validator = new OpenXmlValidator(FileFormatVersions.Office2016);
        var errors = validator.Validate(doc).ToList();

        Assert.True(
            errors.Count == 0,
            "OpenXML validation errors:\n" +
            string.Join(Environment.NewLine,
                errors.Select(e => $"{e.Path.XPath}: {e.Description}"))
        );
    }

    // ============================================================
    // 1) IN-MEMORY BASIC TESTS
    // ============================================================
    [Fact]
    public void InMemory_Workbook_ShouldBeSchemaValid()
    {
        var bytes = new WorkbookBuilder()
            .AddSheet("Test", s =>
            {
                s.AddRow("Id", "Name");
                s.AddRow(1, "Alex");
                s.AddRow(2, "Brian");
            })
            .Build();

        AssertSchemaValid(bytes);
    }

    [Fact]
    public void InMemory_Table_ShouldBeSchemaValid()
    {
        var bytes = new WorkbookBuilder()
            .AddSheet("Employees", s =>
            {
                s.AddRow("Id", "Name", "Active");
                s.AddRow(1, "Alex", true);
                s.AddRow(2, "Brian", false);
                s.AddTable("EmployeesTable");
            })
            .Build();

        AssertSchemaValid(bytes);
    }

    [Fact]
    public void InMemory_WithEmptyRowsBeforeHeader_ShouldBeSchemaValid()
    {
        var bytes = new WorkbookBuilder()
            .AddSheet("WithGaps", s =>
            {
                s.AddEmptyRows(2);
                s.AddRow("Id", "Name");
                s.AddRow(1, "Alex");
                s.AddRow(2, "Brian");
                s.AddTable("GapTable");
            })
            .Build();

        AssertSchemaValid(bytes);
    }

    [Fact]
    public void InMemory_Hyperlinks_ShouldBeSchemaValid()
    {
        var bytes = new WorkbookBuilder()
            .AddSheet("Links", s =>
            {
                s.AddRow("Id", "Website");
                s.AddRow(1, XL.Hyper("https://google.com", "Google"));
                s.AddRow(2, XL.Hyper("https://github.com/livedcode/OpenExcelLite", "Repo"));
            })
            .Build();

        AssertSchemaValid(bytes);

        using var ms = new MemoryStream(bytes);
        using var doc = SpreadsheetDocument.Open(ms, false);

        var links = doc.WorkbookPart.WorksheetParts
            .SelectMany(ws => ws.HyperlinkRelationships)
            .ToList();

        Assert.Equal(2, links.Count);
    }


    // ============================================================
    // 2) STREAMING - MULTI-SHEET TESTS
    // ============================================================

    [Fact]
    public void Streaming_MultiSheet_ShouldBeSchemaValid()
    {
        var bytes = StreamingWorkbookBuilder.Build(wb =>
        {
            wb.AddSheet("Users", s =>
            {
                s.WriteRow("Id", "Name");
                s.WriteRow(1, "Alex");
            });

            wb.AddSheet("Logs", s =>
            {
                s.WriteRow("Timestamp", "Message");
                s.WriteRow(DateTime.Now, "Started");
            });
        });

        AssertSchemaValid(bytes);

        using var ms = new MemoryStream(bytes);
        using var doc = SpreadsheetDocument.Open(ms, false);

        Assert.Equal(2, doc.WorkbookPart.WorksheetParts.Count());
    }

    [Fact]
    public void Streaming_MultiSheetWithEmptyRows_ShouldBeSchemaValid()
    {
        var bytes = StreamingWorkbookBuilder.Build(wb =>
        {
            wb.AddSheet("Sheet1", s =>
            {
                s.WriteEmptyRows(3);
                s.WriteRow("A", "B");
                s.WriteRow(1, 2);
            });

            wb.AddSheet("Sheet2", s =>
            {
                s.WriteRow("X", "Y");
                s.WriteEmptyRows(2);
                s.WriteRow(5, 6);
            });
        });

        AssertSchemaValid(bytes);
    }

    // ============================================================
    // 3) STREAMING - HYPERLINK TESTS
    // ============================================================

    [Fact]
    public void Streaming_MultiSheet_Hyperlinks_ShouldBeSchemaValid()
    {
        var bytes = StreamingWorkbookBuilder.Build(wb =>
        {
            wb.AddSheet("Links1", s =>
            {
                s.WriteRow("Id", "Link");
                s.WriteRow(1, XL.Hyper("https://google.com", "Google"));
            });

            wb.AddSheet("Links2", s =>
            {
                s.WriteRow("Key", "Url");
                s.WriteRow("Repo", XL.Hyper("https://github.com/livedcode/OpenExcelLite"));
            });
        });

        AssertSchemaValid(bytes);

        using var ms = new MemoryStream(bytes);
        using var doc = SpreadsheetDocument.Open(ms, false);

        var allLinks = doc.WorkbookPart.WorksheetParts
            .SelectMany(ws => ws.HyperlinkRelationships)
            .ToList();

        Assert.Equal(2, allLinks.Count);

        Assert.Contains(allLinks, l => l.Uri.AbsoluteUri.Contains("google"));
        Assert.Contains(allLinks, l => l.Uri.AbsoluteUri.Contains("github"));
    }

    [Fact]
    public void Streaming_Hyperlinks_WithEmptyRows_ShouldBeSchemaValid()
    {
        var bytes = StreamingWorkbookBuilder.Build(wb =>
        {
            wb.AddSheet("Links", s =>
            {
                s.WriteEmptyRows(4);
                s.WriteRow("Id", "Website");
                s.WriteRow(1, XL.Hyper("https://google.com", "Google"));
            });
        });

        AssertSchemaValid(bytes);
    }
}
