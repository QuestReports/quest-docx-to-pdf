using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using QuestReports.Converters.DocXToPdf.Extensions;
using QuestPDF.Fluent;
using QuestPDF.Helpers;

namespace QuestReports.Converters.DocXToPdf.CommonAssets;

public class QuestDocXTextStyling
{
    public string Font { get; }

    public float Size { get; }

    public string Color { get; }

    public bool Bold { get; }

    public bool Italics { get; }

    public bool StrikeThrough { get; }

    public QuestDocXLineType Underline { get; }

    public string Background { get; }

    public QuestDocXTextStyling(RunProperties? runProperties, DocXQuestConversionOptions options)
    {
        Font = runProperties?.RunFonts?.Ascii?.Value?.CheckAndReturnComparableFont(options) ?? options.DefaultFont;
        Size = ParseFontSize(runProperties?.FontSize?.Val, options).GetPxFromPt();
        Background = runProperties?.Highlight?.Val?.Value.GetHexColor() ?? options.DefaultTextHighlightColor;
        Color = runProperties?.Color?.Val?.GetHexColor() ?? Colors.Black;
        Bold = runProperties?.Bold is not null;
        Italics = runProperties?.Italic is not null;
        StrikeThrough = runProperties?.Strike is not null;
        Underline = runProperties?.Underline?.Val.GetLineTypeFromLineValue() ?? QuestDocXLineType.None;
    }

    public TextSpanDescriptor WrapTextStyle(TextSpanDescriptor descriptor)
    {
        if (Bold)
            descriptor.Bold();
        if (Italics)
            descriptor.Italic();
        if (StrikeThrough)
            descriptor.Strikethrough();
        if (Underline != QuestDocXLineType.None)
            descriptor.Underline();
        descriptor
            .FontColor(Color)
            .BackgroundColor(Background)
            .FontFamily(Font)
            .FontSize(Size);
        return descriptor;
    }

    private static float ParseFontSize(StringValue? value, DocXQuestConversionOptions options)
    {
        if (value is null || !value.HasValue)
            return options.DefaultFontSize;
        return float.TryParse(value, out var result) ? result : options.DefaultFontSize;
    }
}
