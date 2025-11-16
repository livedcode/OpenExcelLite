using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelLite.Internals;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelLite.Builders;
/// <summary>
/// Helper for building workbooks using streaming writers, suitable for huge datasets.
/// </summary>
public static class StreamingWorkbookBuilder
{
    public static byte[] Build(string sheetName, Action<StreamingWorksheetWriter> configure)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new ArgumentException("Sheet name cannot be empty.", nameof(sheetName));
        if (configure == null)
            throw new ArgumentNullException(nameof(configure));

        using var ms = new MemoryStream();

        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            uint dateStyleIndex = StyleFactory.EnsureDefaultStyles(workbookPart);

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            using (var writer = new StreamingWorksheetWriter(worksheetPart, dateStyleIndex))
            {
                configure(writer);
            }

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            sheets.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1U,
                Name = sheetName
            });

            workbookPart.Workbook.Save();
        }

        if (ms.Length == 0)
            throw new InvalidOperationException("Streaming build failed: resulting stream is empty.");

        return ms.ToArray();
    }
}