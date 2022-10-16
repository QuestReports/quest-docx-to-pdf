using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QuestReports.Converters.DocXToPdf.Extensions;
using QuestPDF.Fluent;

namespace QuestReports.Converters.DocXToPdf.CommonAssets;

internal class QuestDocXPageProperties
{
    public string BackgroundColor { get; }

    public QuestDocXMargins Margins { get; }

    public QuestPDF.Helpers.PageSize PageSize { get; }

    public QuestDocXPageProperties(WordprocessingDocument document, DocXQuestConversionOptions options)
    {
        BackgroundColor = document
                              .MainDocumentPart
                              .Document
                              .DocumentBackground?
                              .Background?
                              .Fillcolor
                              .GetHexColor()
                          ?? options.DefaultBackgroundColor;

        var sectionProps = document
            .MainDocumentPart
            .Document
            .Body
            .ChildElements
            .OfType<SectionProperties>()
            .FirstOrDefault();
        var margins = sectionProps?.ChildElements.OfType<PageMargin>().FirstOrDefault();
        Margins = new QuestDocXMargins
        {
            Bottom = margins?.Bottom?.GetPxFromDxa() ?? 0f,
            Top = margins?.Top?.GetPxFromDxa() ?? 0f,
            Left = margins?.Left?.GetPxFromDxa() ?? 0f,
            Right = margins?.Right?.GetPxFromDxa()?? 0f
        };

        var size = sectionProps?.ChildElements.OfType<PageSize>().FirstOrDefault();
        PageSize = new QuestPDF.Helpers.PageSize(
            size?.Width.GetPxFromDxa() ?? 0,
            size?.Height.GetPxFromDxa() ?? 0);
    }

    public PageDescriptor WrapPageStyle(PageDescriptor descriptor)
    {
        descriptor.PageColor(BackgroundColor);
        descriptor.Size(PageSize);
        descriptor.MarginLeft(Margins.Left);
        descriptor.MarginRight(Margins.Right);
        return descriptor;
    }
}
