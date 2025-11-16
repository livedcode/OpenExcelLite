using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelLite.Internals;

namespace OpenExcelLite.Builders;
/// <summary>
/// Fluent builder for an Excel worksheet (in-memory).
/// Fully compliant with Excel table rules.
/// </summary>
public sealed class WorksheetBuilder
{
    private readonly List<List<object?>> _sheetRows = new();
    private bool _enableSheetFilter;
    private bool _enableAutoFit;
    private TableBuilder? _sheetTableBuilder;

    public string SheetName { get; }

    internal WorksheetBuilder(string sheetName)
    {
        SheetName = sheetName;
    }

    public WorksheetBuilder AddRow(params object?[] values)
    {
        if (values == null || values.Length == 0)
            throw new ArgumentException("Row must contain at least one value.", nameof(values));

        // FIRST ROW = HEADER ROW
        if (_sheetRows.Count == 0)
        {
            var headerNames = values.Select(v => v?.ToString()?.Trim() ?? "").ToList();

            // A. Must not contain empty header names
            if (headerNames.Any(string.IsNullOrWhiteSpace))
                throw new InvalidOperationException("Header row cannot contain empty or null column names.");

            // B. Deduplicate header names (case-insensitive)
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

            // C. Assign back sanitized header names as strings
            for (int i = 0; i < values.Length; i++)
                values[i] = headerNames[i];
        }
        else
        {
            // Data rows must match header width
            if (values.Length != _sheetRows[0].Count)
                throw new InvalidOperationException(
                    $"Row has {values.Length} cells but header has {_sheetRows[0].Count}.");
        }

        _sheetRows.Add(values.ToList());
        return this;
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

        int columnCount = _sheetRows[0].Count;
        int rowCount = _sheetRows.Count;

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        var widthHelper = new ColumnWidthHelper();

        uint rowIndex = 1;

        foreach (var rowValues in _sheetRows)
        {
            var row = new Row { RowIndex = rowIndex };

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
        {
            worksheet.Append(widthHelper.BuildColumns());
        }

        worksheet.Append(sheetData);

        if (_enableSheetFilter && _sheetTableBuilder == null)
        {
            string lastCol = GetColumnName(columnCount);
            worksheet.Append(new AutoFilter
            {
                Reference = $"A1:{lastCol}{rowCount}"
            });
        }

        worksheetPart.Worksheet = worksheet;
        worksheetPart.Worksheet.Save();

        _sheetTableBuilder?.Build(worksheetPart, columnCount, rowCount);

        return worksheetPart;
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

        return value switch
        {
            string s => new Cell
            {
                CellReference = cellRef,
                DataType = CellValues.String,
                CellValue = new(s)
            },
            bool b => new Cell
            {
                CellReference = cellRef,
                DataType = CellValues.Boolean,
                CellValue = new(b ? "1" : "0")
            },
            DateTime dt => new Cell
            {
                CellReference = cellRef,
                StyleIndex = dateStyleIndex,
                CellValue = new(dt.ToOADate().ToString(CultureInfo.InvariantCulture))
            },
            int or long or float or double or decimal => new Cell
            {
                CellReference = cellRef,
                DataType = CellValues.Number,
                CellValue = new(Convert.ToString(value, CultureInfo.InvariantCulture))
            },
            _ => new Cell
            {
                CellReference = cellRef,
                DataType = CellValues.String,
                CellValue = new(value.ToString())
            }
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
            DateTime dt => dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
            _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""
        };
}