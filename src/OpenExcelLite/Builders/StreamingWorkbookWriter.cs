using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelLite.Internals;
using System;

namespace OpenExcelLite.Builders;

public sealed class StreamingWorkbookWriter : IDisposable
{
    private readonly SpreadsheetDocument _doc;
    private readonly WorkbookPart _workbookPart;
    private readonly Sheets _sheets;
    private uint _sheetIdCounter = 1;
    private bool _disposed;

    internal StreamingWorkbookWriter(SpreadsheetDocument doc)
    {
        _doc = doc;

        _workbookPart = doc.AddWorkbookPart();
        _workbookPart.Workbook = new Workbook();

        _sheets = new Sheets();
        _workbookPart.Workbook.Append(_sheets);
    }

    /// <summary>
    /// Adds a new sheet to the workbook.
    /// </summary>
    public void AddSheet(string sheetName, Action<StreamingWorksheetWriter> configure)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new ArgumentException("Sheet name cannot be empty.", nameof(sheetName));
        if (configure == null)
            throw new ArgumentNullException(nameof(configure));

        var sheetPart = _workbookPart.AddNewPart<WorksheetPart>();
        uint dateStyle = StyleFactory.EnsureDefaultStyles(_workbookPart);

        using (var writer = new StreamingWorksheetWriter(sheetPart, dateStyle))
        {
            configure(writer);
        }

        var sheet = new Sheet
        {
            Id = _workbookPart.GetIdOfPart(sheetPart),
            SheetId = _sheetIdCounter++,
            Name = sheetName
        };

        _sheets.Append(sheet);
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _workbookPart.Workbook.Save();
        _disposed = true;
    }
}
