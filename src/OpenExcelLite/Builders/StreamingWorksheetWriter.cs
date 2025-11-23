using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelLite.Models;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace OpenExcelLite.Builders;

/// <summary>
/// Ultra-fast writer for huge Excel worksheets (100k–1M rows),
/// streaming rows directly to XML using OpenXmlWriter.
/// Fully supports:
/// - Hyperlinks (external URLs)
/// - Empty rows
/// - Date styles
/// - Header column consistency checking
/// </summary>
public sealed class StreamingWorksheetWriter : IDisposable
{
    private readonly WorksheetPart _worksheetPart;
    private readonly OpenXmlWriter _writer;
    private readonly uint _dateStyleIndex;

    private uint _currentRowIndex = 1;
    private int _headerColumnCount = 0;
    private bool _headerWritten = false;

    // NEW: track hyperlinks so we can write <hyperlinks> after SheetData
    private readonly List<(uint RowIndex, int ColIndex, HyperlinkCell Link)> _hyperlinks = new();

    public StreamingWorksheetWriter(WorksheetPart worksheetPart, uint dateStyleIndex)
    {
        _worksheetPart = worksheetPart ?? throw new ArgumentNullException(nameof(worksheetPart));
        _dateStyleIndex = dateStyleIndex;

        _writer = OpenXmlWriter.Create(worksheetPart);

        // <worksheet>
        _writer.WriteStartElement(new Worksheet());

        // <sheetData>
        _writer.WriteStartElement(new SheetData());
    }

    // ============================================================
    // WriteEmptyRows() for streaming mode
    // ============================================================
    public void WriteEmptyRows(int count)
    {
        if (count <= 0)
            return;

        for (int i = 0; i < count; i++)
        {
            _writer.WriteStartElement(new Row { RowIndex = _currentRowIndex });
            _writer.WriteEndElement(); // </row>
            _currentRowIndex++;
        }
    }

    // ============================================================
    // WriteRow()
    // ============================================================
    public void WriteRow(params object?[] values)
    {
        if (values == null || values.Length == 0)
            throw new ArgumentException("Row must contain at least one value.");

        // Header validation
        if (!_headerWritten)
        {
            _headerColumnCount = values.Length;
            _headerWritten = true;
        }
        else if (values.Length != _headerColumnCount)
        {
            throw new InvalidOperationException(
                $"Data row has {values.Length} cells but header has {_headerColumnCount}.");
        }

        _writer.WriteStartElement(new Row { RowIndex = _currentRowIndex });

        for (int col = 0; col < values.Length; col++)
        {
            WriteCell(values[col], col + 1, _currentRowIndex);
        }

        _writer.WriteEndElement(); // </row>
        _currentRowIndex++;
    }

    // ============================================================
    // WriteCell()
    // ============================================================
    private void WriteCell(object? value, int columnIndex, uint rowIndex)
    {
        string cellRef = GetColumnName(columnIndex) + rowIndex;

        // --------------------------
        // Handle null
        // --------------------------
        if (value == null)
        {
            _writer.WriteElement(new Cell
            {
                CellReference = cellRef,
                DataType = CellValues.String,
                CellValue = new CellValue("")
            });
            return;
        }

        // --------------------------
        // NEW: Hyperlink cell
        // --------------------------
        if (value is HyperlinkCell link)
        {
            // Write display text cell
            _writer.WriteElement(new Cell
            {
                CellReference = cellRef,
                DataType = CellValues.String,
                CellValue = new CellValue(link.Display)
            });

            // Record hyperlink for later <hyperlinks> output
            _hyperlinks.Add((rowIndex, columnIndex, link));
            return;
        }

        // --------------------------
        // Normal primitive types
        // --------------------------
        switch (value)
        {
            case string s:
                _writer.WriteElement(new Cell
                {
                    CellReference = cellRef,
                    DataType = CellValues.String,
                    CellValue = new CellValue(s)
                });
                break;

            case bool b:
                _writer.WriteElement(new Cell
                {
                    CellReference = cellRef,
                    DataType = CellValues.Boolean,
                    CellValue = new CellValue(b ? "1" : "0")
                });
                break;

            case DateTime dt:
                _writer.WriteElement(new Cell
                {
                    CellReference = cellRef,
                    StyleIndex = _dateStyleIndex,
                    CellValue = new CellValue(dt.ToOADate().ToString(CultureInfo.InvariantCulture))
                });
                break;

            case int or long or float or double or decimal:
                _writer.WriteElement(new Cell
                {
                    CellReference = cellRef,
                    DataType = CellValues.Number,
                    CellValue = new CellValue(Convert.ToString(value, CultureInfo.InvariantCulture))
                });
                break;

            default:
                _writer.WriteElement(new Cell
                {
                    CellReference = cellRef,
                    DataType = CellValues.String,
                    CellValue = new CellValue(value.ToString())
                });
                break;
        }
    }

    // ============================================================
    // Helpers
    // ============================================================
    private static string GetColumnName(int index)
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

    // ============================================================
    // Dispose()
    // ============================================================
    public void Dispose()
    {
        // Close <sheetData>
        _writer.WriteEndElement();

        // ========================================================
        // Write <hyperlinks> AFTER SheetData but BEFORE Worksheet ends
        // ========================================================
        if (_hyperlinks.Count > 0)
        {
            _writer.WriteStartElement(new Hyperlinks());

            foreach (var (rowIndex, colIndex, link) in _hyperlinks)
            {
                string cellRef = GetColumnName(colIndex) + rowIndex;

                // Relationship created here
                string relId = _worksheetPart.AddHyperlinkRelationship(
                    new Uri(link.Url, UriKind.Absolute),
                    true
                ).Id;

                _writer.WriteElement(new Hyperlink
                {
                    Reference = cellRef,
                    Id = relId,
                    Display = link.Display
                });
            }

            _writer.WriteEndElement(); // </hyperlinks>
        }

        // Close </worksheet>
        _writer.WriteEndElement();

        _writer.Dispose();
    }
}
