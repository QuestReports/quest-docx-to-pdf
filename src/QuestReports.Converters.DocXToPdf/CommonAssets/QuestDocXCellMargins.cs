using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace QuestReports.Converters.DocXToPdf.CommonAssets;

public class QuestDocXCellMargins
{
    public float Top { get; }

    public float Bottom { get; }

    public float Right { get; } = DocXFormatConstants.DefaultStrokeScale;

    internal float Left { get; } = DocXFormatConstants.DefaultStrokeScale;

    public QuestDocXCellMargins(
        TableCellMargin? margin,
        DocXQuestConversionOptions options,
        bool isInheritedFromTable = false)
    {
        if (isInheritedFromTable)
            return;

        if (margin is null)
        {
            Top = options.DefaultCellVerticalMargins;
            Bottom = options.DefaultCellVerticalMargins;
            Right = options.DefaultCellHorizontalMargins;
            Left = options.DefaultCellHorizontalMargins;
            return;
        }

        Top = Parse(margin.TopMargin.Width) / DocXFormatConstants.DxaScale;
        Bottom = Parse(margin.BottomMargin.Width) / DocXFormatConstants.DxaScale;
        Right = Parse(margin.RightMargin.Width) / DocXFormatConstants.DxaScale;
        Left = Parse(margin.LeftMargin.Width) / DocXFormatConstants.DxaScale;
    }

    private static float Parse(StringValue? value)
        => value?.HasValue == true ? float.Parse(value.Value) : 0;
}
