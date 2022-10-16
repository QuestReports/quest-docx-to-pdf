using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QuestPDF.Helpers;
using Picture = DocumentFormat.OpenXml.Drawing.Pictures.Picture;

namespace QuestReports.Converters.DocXToPdf.Extensions;

public static class OXmlElementsExtensions
{
    public static float GetParagraphRightIndentation(this Paragraph paragraph)
        => paragraph.ParagraphProperties?.Indentation?.Right?.Value.GetPxFromDxa() ?? 0;

    public static float GetParagraphLeftIndentation(this Paragraph paragraph)
        => paragraph.ParagraphProperties?.Indentation?.Left?.Value.GetPxFromDxa()
           ?? paragraph.ParagraphProperties?.Indentation?.FirstLine?.Value.GetPxFromDxa()
           ?? 0;

    public static EnumValue<TextDirectionValues>? GetParagraphTextDirection(this Paragraph paragraph)
        => paragraph.ParagraphProperties?.TextDirection?.Val;

    public static TableRow[] GetRows(this Table table)
        => table.ChildElements.OfType<TableRow>().ToArray();

    public static TableCell[] GetCells(this TableRow row)
        => row.ChildElements.OfType<TableCell>().ToArray();

    public static uint GetHorizontalMergeFromGridSpan(this TableCellProperties? properties)
        => properties?.GridSpan?.Val?.Value is null or 0
            ? 1 // if there is no merge use 1 to represent it (1 cell takes 1 column)
            : (uint)properties.GridSpan.Val.Value;

    public static bool IsStartOfVerticalMerge(this TableCellProperties? properties)
        => properties?.VerticalMerge?.Val?.Value == MergedCellValues.Restart;

    public static TableVerticalAlignmentValues GetVerticalCellAlignmentOrDefault(this TableCellProperties? properties)
        => properties?.TableCellVerticalAlignment?.Val?.Value ?? TableVerticalAlignmentValues.Center;

    public static TableBorders? GetTableBorders(this Table table)
        => table.ChildElements?.OfType<TableProperties>()?.FirstOrDefault()?.TableBorders;

    public static string GetCellFillHexColorOrDefault(this TableCellProperties? cellProperties)
        => cellProperties?.Shading?.Fill?.GetHexColor() ?? Colors.White;

    public static uint? GetTableRowHeight(this TableRow row)
        => row.TableRowProperties?
            .OfType<TableRowHeight>()
            .FirstOrDefault()?
            .Val?
            .Value;

    public static string? GetTableRowHeightByInnerContent(this TableRow row)
        => row.Descendants<RunProperties>()
            .MaxBy(x => x.FontSize?.Val?.Value?.GetPxFromPt())?
            .FontSize?
            .Val?
            .Value;

    public static TableGrid? GetTableGrid(this Table table)
        => table.ChildElements.OfType<TableGrid>().FirstOrDefault();

    public static GridColumn[] GetGridColumns(this TableGrid grid)
        => grid.ChildElements.OfType<GridColumn>().ToArray();

    public static T? Get<T>(this Run run)
        => run.ChildElements.OfType<T>().FirstOrDefault();

    public static Picture? GetPicture(this Drawing drawing)
        => drawing.Inline?.Graphic?.GraphicData?
            .Descendants<Picture>().FirstOrDefault();

    public static string? GetBlip(this Picture picture)
        => picture.BlipFill?.Blip?.Embed?.Value;

    public static Extent? GetExtent(this Drawing drawing)
        => drawing.Descendants<Extent>().FirstOrDefault();

    public static OpenXmlPart? GetMainPartById(this Document document, string partId)
        => document.MainDocumentPart.GetPartById(partId);

    public static OpenXmlPart? GetHeaderPartById(this Document document, string partId)
        => document.MainDocumentPart.HeaderParts?.FirstOrDefault()?.GetPartById(partId);

    public static OpenXmlPart? GetFooterPartById(this Document document, string partId)
        => document.MainDocumentPart.FooterParts?.FirstOrDefault()?.GetPartById(partId);

    public static float? GetFontSizeInParagraph(this Paragraph paragraph)
        => paragraph.ParagraphProperties?
            .ChildElements?
            .OfType<ParagraphMarkRunProperties>()?
            .FirstOrDefault()?
            .ChildElements
            .OfType<FontSize>()?
            .FirstOrDefault()?
            .Val?
            .Value?
            .GetPxFromPt();

    public static bool ContainsTable(this TableCell cell)
        => cell.ChildElements.OfType<Table>().Any();

    public static bool ContainsPageBreak(this Paragraph paragraph)
        => paragraph.Descendants<Break>().Any(x => x.Type?.Value == BreakValues.Page);

    public static IEnumerable<HeaderPart> GetHeadersParts(this WordprocessingDocument document)
        => document.MainDocumentPart.HeaderParts;

    public static IEnumerable<FooterPart> GetFootersParts(this WordprocessingDocument document)
        => document.MainDocumentPart.FooterParts;

    public static Document GetDocument(this WordprocessingDocument document)
        => document.MainDocumentPart.Document;

    public static OpenXmlElementList GetDocumentBodyChildElements(this WordprocessingDocument document)
        => document.MainDocumentPart.Document.Body.ChildElements;

    public static float GetParagraphSpacingOrDefault(this Paragraph paragraph)
        => paragraph.Descendants<RunProperties>()
            .FirstOrDefault()?
            .Spacing?
            .Val.GetPxFromDxa() ?? 0;

    public static JustificationValues GetParagraphAlignmentOrDefault(this Paragraph paragraph)
        => paragraph.ParagraphProperties?.Justification?.Val?.Value ?? JustificationValues.Left;
}
