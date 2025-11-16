
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelLite.Internals;
using System.Globalization;

namespace OpenExcelLite.Builders;
/// <summary>
/// Entry point for building an Excel workbook in memory using a fluent API.
/// </summary>
public sealed class WorkbookBuilder
{
    private readonly List<WorksheetBuilder> _worksheets = new();

    public WorkbookBuilder AddSheet(string sheetName, Action<WorksheetBuilder> configure)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new ArgumentException("Sheet name cannot be empty.", nameof(sheetName));
        if (configure == null)
            throw new ArgumentNullException(nameof(configure));

        var builder = new WorksheetBuilder(sheetName);
        configure(builder);
        _worksheets.Add(builder);

        return this;
    }

    public byte[] Build()
    {
        if (_worksheets.Count == 0)
            throw new InvalidOperationException("Workbook must contain at least one sheet.");

        using var ms = new MemoryStream();

        using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            uint dateStyleIndex = StyleFactory.EnsureDefaultStyles(workbookPart);

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            uint sheetId = 1;

            foreach (var wsBuilder in _worksheets)
            {
                var sheetPart = wsBuilder.Build(workbookPart, dateStyleIndex);

                sheets.Append(new Sheet
                {
                    Id = workbookPart.GetIdOfPart(sheetPart),
                    SheetId = sheetId++,
                    Name = wsBuilder.SheetName
                });
            }

            workbookPart.Workbook.Save();
        }

        if (ms.Length == 0)
            throw new InvalidOperationException("Workbook build failed: resulting stream is empty.");

        return ms.ToArray();
    }
}