using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using QuestReports.Converters.DocXToPdf.CommonAssets;

namespace QuestReports.Converters.DocXToPdf.Extensions;

public static class QuestExtensions
{
    public static IContainer VerticallyAlign(this IContainer container, TableVerticalAlignmentValues values)
    {
        return values switch
        {
            TableVerticalAlignmentValues.Bottom => container.AlignBottom(),
            TableVerticalAlignmentValues.Center => container.AlignMiddle(),
            TableVerticalAlignmentValues.Top => container.AlignTop(),
            _ => throw new ArgumentOutOfRangeException(nameof(values), values, null)
        };
    }

    public static void HorizontallyAlign(this TextDescriptor descriptor, JustificationValues values)
    {
        switch (values)
        {
            case JustificationValues.Left:
                descriptor.AlignLeft();
                break;
            case JustificationValues.Center:
                descriptor.AlignCenter();
                break;
            case JustificationValues.Right:
                descriptor.AlignRight();
                break;
            default:
                descriptor.AlignLeft();
                break;
        }
    }

    public static IContainer ResolveTextDirection(this IContainer container, EnumValue<TextDirectionValues>? values)
    {
        if (values is null)
            return container;
        if (!values.HasValue)
            return container;
        container
            .TranslateX(50)
            .TranslateY(50);
        switch (values.Value)
        {
            case TextDirectionValues.LefToRightTopToBottom:
                container.RotateRight();
                break;
            case TextDirectionValues.LeftToRightTopToBottom2010:
                container.RotateRight();
                break;
            case TextDirectionValues.TopToBottomRightToLeft:
                container.RotateRight();
                break;
            case TextDirectionValues.TopToBottomRightToLeft2010:
                container.RotateRight();
                break;
            case TextDirectionValues.BottomToTopLeftToRight:
                container.RotateLeft();
                break;
            case TextDirectionValues.BottomToTopLeftToRight2010:
                container.RotateLeft();
                break;
            case TextDirectionValues.LefttoRightTopToBottomRotated:
                container.RotateRight();
                break;
            case TextDirectionValues.LeftToRightTopToBottomRotated2010:
                container.RotateRight();
                break;
            case TextDirectionValues.TopToBottomRightToLeftRotated:
                container.RotateRight();
                break;
            case TextDirectionValues.TopToBottomRightToLeftRotated2010:
                container.RotateRight();
                break;
            case TextDirectionValues.TopToBottomLeftToRightRotated:
                container.RotateRight();
                break;
            case TextDirectionValues.TopToBottomLeftToRightRotated2010:
                container.RotateRight();
                break;
        }

        return container.TranslateX(-50)
            .TranslateY(-50);
    }

    public static IContainer ApplyCellBorders(this IContainer container, QuestDocXCellBorders borders)
    {
        if (borders.Bottom.QuestDocXLineType != QuestDocXLineType.None)
            container = container.BorderBottom(borders.Bottom.Thickness)
                .BorderColor(borders.Bottom.ColorHex);
        if (borders.Top.QuestDocXLineType != QuestDocXLineType.None)
            container = container.BorderTop(borders.Top.Thickness)
                .BorderColor(borders.Top.ColorHex);
        if (borders.Left.QuestDocXLineType != QuestDocXLineType.None)
            container = container.BorderLeft(borders.Left.Thickness)
                .BorderColor(borders.Left.ColorHex);
        if (borders.Right.QuestDocXLineType != QuestDocXLineType.None)
            container = container.BorderRight(borders.Right.Thickness)
                .BorderColor(borders.Right.ColorHex);
        return container;
    }

    public static IContainer ApplyPaddings(this IContainer container, QuestDocXCellMargins cellMargins)
        => container
            .PaddingBottom(cellMargins.Bottom)
            .PaddingTop(cellMargins.Top)
            .PaddingLeft(cellMargins.Left)
            .PaddingRight(cellMargins.Right);

    public static string CheckAndReturnComparableFont(this string? font, DocXQuestConversionOptions options)
    {
        if (font is null)
            return options.DefaultFont;
        var fontToUse = DocXFormatConstants.Fonts.TryGetValue(font, out var value) ? value : options.DefaultFont;
        return fontToUse;
    }
}
