using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelLite.Internals;
using OpenExcelLite.Models;

namespace OpenExcelLite.Builders;

public sealed class WorksheetBuilder
{
    private readonly List<List<object?>> _sheetRows = new();

    private bool _hasHeaderRow;
    private int _headerColumnCount;
    private uint _headerRowIndex;

    private bool _enableSheetFilter;
    private bool _enableAutoFit;
    private TableBuilder? _sheetTableBuilder;

    public string SheetName { get; }

    internal WorksheetBuilder(string sheetName)
    {
        SheetName = sheetName;
    }

    public WorksheetBuilder AddEmptyRows(int count)
    {
        if (count <= 0)
            return this;

        for (int i = 0; i < count; i++)
            _sheetRows.Add(new List<object?>());

        return this;
    }

    public WorksheetBuilder AddRow(params object?[] values)
    {
        if (values == null || values.Length == 0)
            throw new ArgumentException("Row must contain values.");

        if (!_hasHeaderRow)
        {
            AddHeaderRow(values);
        }
        else
        {
            AddDataRow(values);
        }
        return this;
    }

    private void AddHeaderRow(object?[] values)
    {
        var headerNames = values.Select(v => v?.ToString()?.Trim() ?? "").ToList();

        if (headerNames.Any(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException("Header row contains empty column names.");

        for (int i = 0; i < headerNames.Count; i++)
        {
            string baseName = headerNames[i];
            int suffix = 1;

            while (headerNames
                   .Take(i)
                   .Contains(headerNames[i], StringComparer.OrdinalIgnoreCase))
            {
                headerNames[i] = $"{baseName}_{suffix++}";
            }
        }

        _sheetRows.Add(headerNames.Cast<object?>().ToList());

        _hasHeaderRow = true;
        _headerColumnCount = headerNames.Count;
        _headerRowIndex = (uint)_sheetRows.Count;
    }

    private void AddDataRow(object?[] values)
    {
        if (values.Length != _headerColumnCount)
            throw new InvalidOperationException(
                $"Data row has {values.Length} cells but header has {_headerColumnCount}.");

        _sheetRows.Add(values.ToList());
    }

    public WorksheetBuilder ApplyAutoFilter()
    {
        _enableSheetFilter = true;
        return this;
    }

    public WorksheetBuilder AutoFitColumns()
    {
        _enableAutoFit = true;
        return this;
    }

    public WorksheetBuilder AddTable(string tableName, string styleName = "TableStyleMedium2")
    {
        _sheetTableBuilder = new TableBuilder(tableName, styleName);
        return this;
    }

    internal WorksheetPart Build(WorkbookPart workbookPart, uint dateStyleIndex)
    {
        if (_sheetRows.Count == 0)
            throw new InvalidOperationException("Worksheet must have at least one row.");
        if (!_hasHeaderRow)
            throw new InvalidOperationException("Worksheet must have a header row.");

        int columnCount = _headerColumnCount;
        int rowCount = _sheetRows.Count;

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        var widthHelper = new ColumnWidthHelper();

        uint rowIndex = 1;

        foreach (var rowValues in _sheetRows)
        {
            var row = new Row { RowIndex = rowIndex };

            bool isEmpty = rowValues.Count == 0 || rowValues.All(v => v == null);

            if (isEmpty)
            {
                sheetData.Append(row);
                rowIndex++;
                continue;
            }

            for (int c = 0; c < columnCount; c++)
            {
                var value = rowValues[c];
                var cell = CreateCell(value, c + 1, rowIndex, dateStyleIndex);

                widthHelper.Track(c + 1, FormatDisplay(value));
                row.Append(cell);
            }

            sheetData.Append(row);
            rowIndex++;
        }

        var worksheet = new Worksheet();

        if (_enableAutoFit)
            worksheet.Append(widthHelper.BuildColumns());

        worksheet.Append(sheetData);

        if (_enableSheetFilter && _sheetTableBuilder == null)
        {
            string lastCol = GetColumnName(columnCount);
            worksheet.Append(new AutoFilter
            {
                Reference = $"A{_headerRowIndex}:{lastCol}{rowCount}"
            });
        }

        worksheetPart.Worksheet = worksheet;
        worksheetPart.Worksheet.Save();

        _sheetTableBuilder?.Build(worksheetPart, columnCount, rowCount);

        AppendHyperlinks(worksheetPart);

        return worksheetPart;
    }

    private void AppendHyperlinks(WorksheetPart worksheetPart)
    {
        var hyperlinkCells = _sheetRows
            .SelectMany((row, rIdx) => row.Select((val, cIdx) => new { val, rIdx, cIdx }))
            .Where(x => x.val is HyperlinkCell)
            .ToList();

        if (hyperlinkCells.Count == 0)
            return;

        var hyperlinks = worksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();
        if (hyperlinks == null)
        {
            hyperlinks = new Hyperlinks();
            worksheetPart.Worksheet.Append(hyperlinks);
        }

        foreach (var item in hyperlinkCells)
        {
            var link = (HyperlinkCell)item.val!;
            uint rowIndex = (uint)(item.rIdx + 1);
            int colIndex = item.cIdx + 1;

            string cellRef = GetColumnName(colIndex) + rowIndex;

            string relId = worksheetPart.AddHyperlinkRelationship(
                new Uri(link.Url, UriKind.Absolute),
                true
            ).Id;

            hyperlinks.Append(new Hyperlink
            {
                Reference = cellRef,
                Id = relId,
                Display = link.Display
            });
        }

        worksheetPart.Worksheet.Save();
    }

    private static Cell CreateCell(object? value, int columnIndex, uint rowIndex, uint dateStyleIndex)
    {
        string cellRef = GetColumnName(columnIndex) + rowIndex;

        if (value == null)
        {
            return new Cell
            {
                CellReference = cellRef,
                DataType = CellValues.String,
                CellValue = new("")
            };
        }

        if (value is HyperlinkCell link)
        {
            return new Cell
            {
                CellReference = cellRef,
                DataType = CellValues.String,
                CellValue = new(link.Display)
            };
        }

        return value switch
        {
            string s => new Cell { CellReference = cellRef, DataType = CellValues.String, CellValue = new(s) },
            bool b => new Cell { CellReference = cellRef, DataType = CellValues.Boolean, CellValue = new(b ? "1" : "0") },
            DateTime dt => new Cell { CellReference = cellRef, StyleIndex = dateStyleIndex, CellValue = new(dt.ToOADate().ToString(CultureInfo.InvariantCulture)) },
            int or long or float or double or decimal => new Cell { CellReference = cellRef, DataType = CellValues.Number, CellValue = new(Convert.ToString(value, CultureInfo.InvariantCulture)) },
            _ => new Cell { CellReference = cellRef, DataType = CellValues.String, CellValue = new(value.ToString()) }
        };
    }

    internal static string GetColumnName(int index)
    {
        string name = "";
        while (index > 0)
        {
            index--;
            name = (char)('A' + index % 26) + name;
            index /= 26;
        }
        return name;
    }

    private static string FormatDisplay(object? value)
        => value switch
        {
            null => "",
            HyperlinkCell link => link.Display,
            DateTime dt => dt.ToString("yyyy-MM-dd HH:mm:ss"),
            _ => Convert.ToString(value) ?? ""
        };
}
