using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelLite.Internals;
/// <summary>
/// Creates basic styles (e.g., date style) for the workbook.
/// </summary>
internal static class StyleFactory
{
    /// <summary>
    /// Create a minimal, ECMA-376-compliant styles.xml and return
    /// the style index used for dates.
    /// </summary>
    public static uint EnsureDefaultStyles(WorkbookPart workbookPart)
    {
        if (workbookPart.WorkbookStylesPart != null)
            return 1; // we assume index 1 is still our date style

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        var stylesheet = new Stylesheet();

        // 1. Custom date format: yyyy-mm-dd (id 164)
        var numFmts = new NumberingFormats();
        numFmts.Append(new NumberingFormat
        {
            NumberFormatId = 164,
            FormatCode = "yyyy-mm-dd"
        });
        numFmts.Count = (uint)numFmts.ChildElements.Count;

        // 2. Fonts (at least one)
        var fonts = new Fonts();
        fonts.Append(new Font());
        fonts.Count = (uint)fonts.ChildElements.Count;

        // 3. Fills (must include none + gray125)
        var fills = new Fills();
        fills.Append(new Fill(new PatternFill { PatternType = PatternValues.None }));
        fills.Append(new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));
        fills.Count = (uint)fills.ChildElements.Count;

        // 4. Borders (at least one)
        var borders = new Borders();
        borders.Append(new Border());
        borders.Count = (uint)borders.ChildElements.Count;

        // 5. cellStyleXfs (style definitions for cell styles)
        var cellStyleFormats = new CellStyleFormats();
        cellStyleFormats.Append(new CellFormat
        {
            NumberFormatId = 0,
            FontId = 0,
            FillId = 0,
            BorderId = 0
        });
        cellStyleFormats.Count = 1;

        // 6. cellXfs (actual cell formats)
        var cellFormats = new CellFormats();

        // index 0 = default
        cellFormats.Append(new CellFormat
        {
            NumberFormatId = 0,
            FontId = 0,
            FillId = 0,
            BorderId = 0,
            ApplyNumberFormat = false
        });

        // index 1 = date style
        cellFormats.Append(new CellFormat
        {
            NumberFormatId = 164,
            FontId = 0,
            FillId = 0,
            BorderId = 0,
            ApplyNumberFormat = true
        });

        cellFormats.Count = (uint)cellFormats.ChildElements.Count;

        // 7. cellStyles (Normal)
        var cellStyles = new CellStyles();
        cellStyles.Append(new CellStyle
        {
            Name = "Normal",
            FormatId = 0,
            BuiltinId = 0
        });
        cellStyles.Count = 1;

        // Assemble stylesheet
        stylesheet.Append(numFmts);
        stylesheet.Append(fonts);
        stylesheet.Append(fills);
        stylesheet.Append(borders);
        stylesheet.Append(cellStyleFormats);
        stylesheet.Append(cellFormats);
        stylesheet.Append(cellStyles);

        stylesPart.Stylesheet = stylesheet;
        stylesPart.Stylesheet.Save();

        // date style index
        return 1;
    }
}