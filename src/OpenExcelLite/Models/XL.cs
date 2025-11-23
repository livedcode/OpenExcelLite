using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelLite.Models;
/// <summary>
/// Simple helper static class so you can write XL.Hyper("url", "text").
/// </summary>
public static class XL
{
    public static HyperlinkCell Hyper(string url, string? text = null)
        => new HyperlinkCell(url, text);
}