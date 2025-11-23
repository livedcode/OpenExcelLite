using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelLite.Models;
/// <summary>
/// Represents a hyperlink cell: a URL and optional display text.
/// </summary>
public sealed class HyperlinkCell
{
    public string Url { get; }
    public string Display { get; }

    public HyperlinkCell(string url, string? display = null)
    {
        Url = url ?? throw new ArgumentNullException(nameof(url));
        Display = string.IsNullOrWhiteSpace(display) ? url : display;
    }
}

