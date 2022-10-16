using JetBrains.Annotations;
using QuestPDF.Drawing;
using QuestPDF.Helpers;

namespace QuestReports.Converters.DocXToPdf;

[PublicAPI]
public class DocXQuestConversionOptions
{
    public DocumentMetadata QuestDocumentMetadata { get; init; } = new();

    public float DefaultCellHorizontalMargins { get; init; } = 4;

    public float DefaultCellVerticalMargins { get; init; }

    public string DefaultFontColor { get; init; } = Colors.Black;

    public string DefaultBackgroundColor { get; init; } = Colors.White;

    public string DefaultTextHighlightColor { get; init; } = Colors.Transparent;

    public string DefaultFont { get; init; } = Fonts.Arial;

    public float DefaultFontSize { get; init; } = 8;

    public float DefaultBreakFontSize { get; init; } = 4;
}
