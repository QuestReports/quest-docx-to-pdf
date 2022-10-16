using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using QuestPDF.Helpers;
using QuestReports.Converters.DocXToPdf.CommonAssets;
using static QuestReports.Converters.DocXToPdf.DocXFormatConstants;

namespace QuestReports.Converters.DocXToPdf.Extensions;

internal static class OXmlStylingExtensions
{
    public static string? GetHexColor(this StringValue value)
        => value.HasValue
            ? value.Value.CheckAutoReturnDefault()
            : null;

    public static string GetHexColor(this HighlightColorValues values)
        => values switch
        {
            HighlightColorValues.Black => Colors.Black,
            HighlightColorValues.Blue => Colors.Blue.Medium,
            HighlightColorValues.Cyan => Colors.Cyan.Medium,
            HighlightColorValues.Green => Colors.Green.Medium,
            HighlightColorValues.Magenta => Colors.Pink.Medium,
            HighlightColorValues.Red => Colors.Red.Medium,
            HighlightColorValues.Yellow => Colors.Yellow.Medium,
            HighlightColorValues.White => Colors.White,
            HighlightColorValues.DarkBlue => Colors.Blue.Darken2,
            HighlightColorValues.DarkCyan => Colors.Cyan.Darken2,
            HighlightColorValues.DarkGreen => Colors.Green.Darken2,
            HighlightColorValues.DarkMagenta => Colors.Pink.Darken2,
            HighlightColorValues.DarkRed => Colors.Red.Darken2,
            HighlightColorValues.DarkYellow => Colors.Yellow.Darken2,
            HighlightColorValues.DarkGray => Colors.Grey.Darken2,
            HighlightColorValues.LightGray => Colors.Grey.Lighten2,
            HighlightColorValues.None => Colors.Transparent,
            _ => Colors.Transparent
        };

    public static float GetPxFromPt(this string? value)
    {
        if (float.TryParse(value, out var result))
            return result / PtFontScale;
        return 0f;
    }

    public static float GetPxFromPt<T>(this T? value) where T : struct, IConvertible
    {
        float converted;
        if (value is null)
            converted = 0;
        else
            converted = Convert.ToSingle(value);
        return converted / PtFontScale;
    }

    public static float GetPxFromPt<T>(this T value) where T : struct, IConvertible
    {
        var converted = Convert.ToSingle(value);
        return converted / PtFontScale;
    }

    public static float GetPxFromDxa<T>(this T? value) where T : struct, IConvertible
    {
        float converted;
        if (value is null)
            converted = 0;
        else
            converted = Convert.ToSingle(value);
        return converted / DxaScale;
    }

    public static float GetPxFromDxa(this string value)
    {
        if (float.TryParse(value,out var result))
            return result / DxaScale;
        return 0f;
    }

    public static float GetPxFromDxa(this UInt32Value? value)
    {
        if (value is null || !value.HasValue)
            return 0f;
        if (float.TryParse(value, out var result))
            return result / DxaScale;
        return 0f;
    }

    public static float GetPxFromDxa(this Int32Value? value)
    {
        if (value is null || !value.HasValue)
            return 0f;
        if (float.TryParse(value, out var result))
            return result / DxaScale;
        return 0f;
    }

    public static float GetPxFromEmu<T>(this T value) where T : struct, IConvertible
    {
        var converted = Convert.ToSingle(value);
        return converted / EmuScale;
    }

    public static QuestDocXLineType GetLineTypeFromLineValue(this EnumValue<BorderValues> borderValues)
    {
        if (!borderValues.HasValue)
            return QuestDocXLineType.Plain;
        return borderValues.Value switch
        {
            BorderValues.Dashed => QuestDocXLineType.Dashed,
            BorderValues.Dotted => QuestDocXLineType.Dotted,
            _ => QuestDocXLineType.Plain
        };
    }

    public static QuestDocXLineType GetLineTypeFromLineValue(this EnumValue<UnderlineValues>? borderValues)
    {
        if (borderValues is null || !borderValues.HasValue)
            return QuestDocXLineType.None;
        return borderValues.Value switch
        {
            UnderlineValues.Dash => QuestDocXLineType.Dashed,
            UnderlineValues.Single => QuestDocXLineType.Plain,
            UnderlineValues.Dotted => QuestDocXLineType.Dotted,
            _ => QuestDocXLineType.None
        };
    }

    private static string CheckAutoReturnDefault(this string value)
        => value == "auto" ? Colors.Black : value;
}
