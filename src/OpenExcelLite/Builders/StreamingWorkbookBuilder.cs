using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;

namespace OpenExcelLite.Builders;

public static class StreamingWorkbookBuilder
{
    /// <summary>
    /// Multi-sheet streaming workbook builder.
    /// </summary>
    public static byte[] Build(Action<StreamingWorkbookWriter> buildAction)
    {
        if (buildAction == null)
            throw new ArgumentNullException(nameof(buildAction));

        using var ms = new MemoryStream();

        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            using var wbWriter = new StreamingWorkbookWriter(doc);
            buildAction(wbWriter);
        }

        if (ms.Length == 0)
            throw new InvalidOperationException("Streaming build failed: resulting stream is empty.");

        return ms.ToArray();
    }
}
