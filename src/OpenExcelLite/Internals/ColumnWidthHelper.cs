using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenExcelLite.Internals;

/// <summary>
/// Tracks maximum text length per column and produces basic auto-fit column widths.
/// </summary>
internal sealed class ColumnWidthHelper
{
    private readonly Dictionary<int, int> _maxLengths = new();

    public void Track(int columnIndex, string displayText)
    {
        if (string.IsNullOrEmpty(displayText))
            return;

        int len = displayText.Length;

        if (_maxLengths.TryGetValue(columnIndex, out var existing))
        {
            if (len > existing)
                _maxLengths[columnIndex] = len;
        }
        else
        {
            _maxLengths[columnIndex] = len;
        }
    }

    public Columns BuildColumns()
    {
        var cols = new Columns();

        foreach (var entry in _maxLengths)
        {
            cols.Append(new Column
            {
                Min = (uint)entry.Key,
                Max = (uint)entry.Key,
                Width = entry.Value + 2,
                CustomWidth = true
            });
        }

        return cols;
    }
}