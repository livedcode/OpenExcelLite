using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OpenExcelLite.Builders;
using OpenExcelLite.Models;

namespace OpenExcelLite.Tests;

public class WorkbookBuilderTests
{
    // -----------------------------------------------------------
    // BASIC TEST
    // -----------------------------------------------------------
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

    // -----------------------------------------------------------
    // TABLE TEST
    // -----------------------------------------------------------
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

    // -----------------------------------------------------------
    // EMPTY ROWS BEFORE HEADER
    // -----------------------------------------------------------
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

    // -----------------------------------------------------------
    // NEW: IN-MEMORY HYPERLINKS
    // -----------------------------------------------------------
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

        // Additional: verify hyperlink relationships exist
        using var ms = new MemoryStream(bytes);
        using var doc = SpreadsheetDocument.Open(ms, false);

        var links = doc.WorkbookPart.WorksheetParts
            .SelectMany(ws => ws.HyperlinkRelationships)
            .ToList();

        Assert.True(links.Count == 2, "Expected 2 hyperlink relationships.");
    }

    // -----------------------------------------------------------
    // STREAMING EMPTY ROWS
    // -----------------------------------------------------------
    [Fact]
    public void Streaming_EmptyRows_ShouldBeSchemaValid()
    {
        var bytes = StreamingWorkbookBuilder.Build("StreamSheet", writer =>
        {
            writer.WriteEmptyRows(5);
            writer.WriteRow("Id", "Name");
            writer.WriteRow(1, "Alex");
            writer.WriteRow(2, "Brian");
        });

        AssertSchemaValid(bytes);
    }

    // -----------------------------------------------------------
    // NEW: STREAMING HYPERLINK TEST
    // -----------------------------------------------------------
    [Fact]
    public void Streaming_Hyperlinks_ShouldBeSchemaValid()
    {
        var bytes = StreamingWorkbookBuilder.Build("LinkStream", writer =>
        {
            writer.WriteRow("Id", "Website");
            writer.WriteRow(1, XL.Hyper("https://google.com", "Google"));
            writer.WriteRow(2, XL.Hyper("https://github.com", "GitHub"));
        });

        AssertSchemaValid(bytes);

        // Verify hyperlink relationships exist
        using var ms = new MemoryStream(bytes);
        using var doc = SpreadsheetDocument.Open(ms, false);

        var links = doc.WorkbookPart.WorksheetParts
            .SelectMany(ws => ws.HyperlinkRelationships)
            .ToList();

        Assert.Equal(2, links.Count);
        Assert.Contains(links, l => l.Uri.AbsoluteUri.Contains("google"));
        Assert.Contains(links, l => l.Uri.AbsoluteUri.Contains("github"));
    }

    // -----------------------------------------------------------
    // VALIDATION HELPER
    // -----------------------------------------------------------
    private static void AssertSchemaValid(byte[] bytes)
    {
        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 0);

        using var ms = new MemoryStream(bytes);
        using var doc = SpreadsheetDocument.Open(ms, false);

        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();

        Assert.True(errors.Count == 0,
            "Validation errors:\n" +
            string.Join(Environment.NewLine,
                errors.Select(e => $"{e.Path.XPath}: {e.Description}")));
    }
}