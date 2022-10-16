using DocumentFormat.OpenXml.Wordprocessing;
using QuestReports.Converters.DocXToPdf.Extensions;
using QuestPDF.Helpers;

namespace QuestReports.Converters.DocXToPdf.CommonAssets;

public class QuestDocXBorder
{
    public float Thickness { get; }

    public string ColorHex { get; }

    public QuestDocXLineType QuestDocXLineType { get; }

    public QuestDocXBorder(BorderType? border, bool isExclude)
    {
        switch (border)
        {
            case null when !isExclude:
                Thickness = 0f;
                ColorHex = Colors.Transparent;
                QuestDocXLineType = QuestDocXLineType.None;
                return;
            case null when isExclude:
                Thickness = DocXFormatConstants.DefaultStrokeScale;
                ColorHex = Colors.Black;
                QuestDocXLineType = QuestDocXLineType.Plain;
                return;
        }

        if (border!.Val.Value == BorderValues.Nil)
        {
            Thickness = 0f;
            ColorHex = Colors.Transparent;
            QuestDocXLineType = QuestDocXLineType.None;
            return;
        }

        Thickness = border.Size?.GetPxFromDxa() ?? DocXFormatConstants.DefaultStrokeScale;
        ColorHex = border.Color?.GetHexColor() ?? Colors.Black;
        QuestDocXLineType = border.Val?.GetLineTypeFromLineValue() ?? QuestDocXLineType.Plain;
    }


}
