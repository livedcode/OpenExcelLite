using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelLite.Builders;
/// <summary>
/// Ultra-fast writer for huge Excel worksheets (100k–1M rows),
/// streaming rows directly to XML using OpenXmlWriter.
/// </summary>
public sealed class StreamingWorksheetWriter : IDisposable
{
    private readonly OpenXmlWriter _writer;
    private readonly uint _dateStyleIndex;
    private uint _currentRowIndex = 1;

    internal StreamingWorksheetWriter(WorksheetPart worksheetPart, uint dateStyleIndex)
    {
        _dateStyleIndex = dateStyleIndex;
        _writer = OpenXmlWriter.Create(worksheetPart);

        // <worksheet>
        _writer.WriteStartElement(new Worksheet());
        // <sheetData>
        _writer.WriteStartElement(new SheetData());
    }

    public void WriteRow(params object?[] values)
    {
        if (values == null || values.Length == 0)
            throw new ArgumentException("Row must contain at least one value.", nameof(values));

        _writer.WriteStartElement(new Row { RowIndex = _currentRowIndex });

        for (int c = 0; c < values.Length; c++)
        {
            string cellRef = WorksheetBuilder.GetColumnName(c + 1) + _currentRowIndex;
            WriteCell(cellRef, values[c]);
        }

        _writer.WriteEndElement(); // </row>
        _currentRowIndex++;
    }

    private void WriteCell(string cellRef, object? value)
    {
        if (value == null)
        {
            _writer.WriteElement(new Cell
            {
                CellReference = cellRef,
                DataType = CellValues.String,
                CellValue = new("")
            });
            return;
        }

        switch (value)
        {
            case string s:
                _writer.WriteElement(new Cell
                {
                    CellReference = cellRef,
                    DataType = CellValues.String,
                    CellValue = new(s)
                });
                break;

            case bool b:
                _writer.WriteElement(new Cell
                {
                    CellReference = cellRef,
                    DataType = CellValues.Boolean,
                    CellValue = new(b ? "1" : "0")
                });
                break;

            case DateTime dt:
                _writer.WriteElement(new Cell
                {
                    CellReference = cellRef,
                    StyleIndex = _dateStyleIndex,
                    CellValue = new(dt.ToOADate().ToString(CultureInfo.InvariantCulture))
                });
                break;

            case int or long or float or double or decimal:
                _writer.WriteElement(new Cell
                {
                    CellReference = cellRef,
                    DataType = CellValues.Number,
                    CellValue = new(Convert.ToString(value, CultureInfo.InvariantCulture))
                });
                break;

            default:
                _writer.WriteElement(new Cell
                {
                    CellReference = cellRef,
                    DataType = CellValues.String,
                    CellValue = new(value.ToString())
                });
                break;
        }
    }

    public void Dispose()
    {
        // close </sheetData>
        _writer.WriteEndElement();
        // close </worksheet>
        _writer.WriteEndElement();
        _writer.Dispose();
    }
}