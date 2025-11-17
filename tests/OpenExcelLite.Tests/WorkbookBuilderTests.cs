using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OpenExcelLite.Builders;

namespace OpenExcelLite.Tests;

public class WorkbookBuilderTests
{
    [Fact]
    public void Streaming_EmptyRows_ShouldBeSchemaValid()
    {
        // Arrange + Act
        var bytes = StreamingWorkbookBuilder.Build("StreamSheet", writer =>
        {
            writer.WriteEmptyRows(5);        // NEW: write 5 empty rows
            writer.WriteRow("Id", "Name");   // header row at row 6
            writer.WriteRow(1, "Alex");      // data row 7
            writer.WriteRow(2, "Brian");     // data row 8
        });

        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 0);

        // Validate OpenXML Schema
        using var ms = new MemoryStream(bytes);
        using var doc = SpreadsheetDocument.Open(ms, false);

        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();

        Assert.True(errors.Count == 0,
            "Schema validation errors:\n" +
            string.Join(Environment.NewLine,
                errors.Select(e => $"{e.Path.XPath}: {e.Description}")));
    }

    [Fact]
    public void InMemory_WithEmptyRowsBeforeHeader_ShouldBeSchemaValid() // NEW
    {
        var bytes = new WorkbookBuilder()
            .AddSheet("WithGaps", s =>
            {
                s.AddEmptyRows(2);             // NEW: two blank rows
                s.AddRow("Id", "Name");        // header on row 3
                s.AddRow(1, "Alex");
                s.AddRow(2, "Brian");
                s.AddTable("GapTable");        // table should start at row 3
            })
            .Build();

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

 

  
    [Fact]
    public void InMemory_Table_ShouldBeSchemaValid()
    {
        var bytes = new WorkbookBuilder()
            .AddSheet("Employees", s =>
            {
                s.AddRow("Id", "Name", "Active");
                s.AddRow(1, "Alex", true);
                s.AddRow(2, "Brian", false);
                s.AddTable("Employees Table");
            })
            .Build();

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

    [Fact]
    public void Streaming_Workbook_ShouldBeSchemaValid()
    {
        var bytes = StreamingWorkbookBuilder.Build("BigSheet", writer =>
        {
            writer.WriteRow("Id", "Value");
            for (int i = 1; i <= 2000; i++)
                writer.WriteRow(i, "Row " + i);
        });

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