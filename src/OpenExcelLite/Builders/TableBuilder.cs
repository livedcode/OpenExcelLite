using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace OpenExcelLite.Builders;
/// <summary>
/// Creates a fully Excel-compliant table (ListObject).
/// This version is tested and does NOT trigger repair warnings.
/// </summary>
public sealed class TableBuilder
{
    private readonly string _requestedName;
    private readonly string _styleName;

    public TableBuilder(string tableName, string styleName)
    {
        _requestedName = tableName;
        _styleName = styleName;
    }

    internal void Build(WorksheetPart wsPart, int columnCount, int rowCount)
    {

        var headerRow = wsPart.Worksheet.Descendants<Row>()
            .FirstOrDefault(r => r.Elements<Cell>().Any());

        if (headerRow == null)
            throw new InvalidOperationException("Cannot create table: worksheet has no non-empty rows.");

        var headerCells = headerRow.Elements<Cell>().ToList();

        if (headerCells.Count != columnCount)
            throw new InvalidOperationException(
                $"Header cell count ({headerCells.Count}) does not match column count ({columnCount}).");

        var headers = headerCells
            .Select(c => c.CellValue?.Text ?? "")
            .ToList();

        if (headers.Any(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException("Header row contains empty column names.");

        uint tableId = GetNextTableId(wsPart);
        string finalTableName = GetUniqueTableName(_requestedName, wsPart);

        string lastCol = WorksheetBuilder.GetColumnName(columnCount);

        uint headerRowIndex = headerRow.RowIndex ?? 1U;
        string refRange = $"A{headerRowIndex}:{lastCol}{rowCount}";

        var tablePart = wsPart.AddNewPart<TableDefinitionPart>();
        string relId = wsPart.GetIdOfPart(tablePart);

        var table = new Table
        {
            Id = tableId,
            Name = finalTableName,
            DisplayName = finalTableName,
            Reference = refRange,
            HeaderRowCount = 1U
        };

        table.Append(new AutoFilter { Reference = refRange });

        var tableColumns = new TableColumns { Count = (uint)columnCount };

        for (uint i = 0; i < columnCount; i++)
        {
            tableColumns.Append(new TableColumn
            {
                Id = i + 1,
                Name = headers[(int)i]
            });
        }

        table.Append(tableColumns);

        table.Append(new TableStyleInfo
        {
            Name = _styleName,
            ShowRowStripes = true
        });

        tablePart.Table = table;
        tablePart.Table.Save();

        var tableParts = wsPart.Worksheet.GetFirstChild<TableParts>();
        if (tableParts == null)
        {
            tableParts = new TableParts { Count = 1U };
            wsPart.Worksheet.Append(tableParts);
        }
        else
        {
            tableParts.Count++;
        }

        tableParts.Append(new TablePart { Id = relId });
        wsPart.Worksheet.Save();
    }

    private static uint GetNextTableId(WorksheetPart wsPart)
    {
        return (uint)(wsPart.TableDefinitionParts.Count() + 1);
    }

    private static string GetUniqueTableName(string name, WorksheetPart wsPart)
    {
        string sanitized = Regex.Replace(name, @"[^A-Za-z0-9_]", "_");

        if (string.IsNullOrEmpty(sanitized))
            sanitized = "Table";

        if (char.IsDigit(sanitized[0]))
            sanitized = "_" + sanitized;

        var existingNames = wsPart.TableDefinitionParts
            .Select(tp => tp.Table?.Name?.Value)
            .Where(n => !string.IsNullOrEmpty(n))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        string finalName = sanitized;
        int counter = 1;

        while (existingNames.Contains(finalName))
            finalName = $"{sanitized}_{counter++}";

        return finalName;
    }
}