using DocumentFormat.OpenXml.Wordprocessing;

namespace QuestReports.Converters.DocXToPdf.CommonAssets;

public class QuestDocXCellBorders
{
    public QuestDocXBorder Top { get; }

    public QuestDocXBorder Bottom { get; }

    public QuestDocXBorder Left { get; }

    public QuestDocXBorder Right { get; }

    public QuestDocXCellBorders(TableCellBorders? borders, TableBorders? tableBorders)
    {
        var isExclude = borders?.Descendants<BorderType>().Any(x => x.Val.Value == BorderValues.Nil) ?? borders is null;
        Top = new QuestDocXBorder(borders?.TopBorder ?? tableBorders?.TopBorder, isExclude);
        Bottom = new QuestDocXBorder(borders?.BottomBorder ?? tableBorders?.BottomBorder, isExclude);
        Left = new QuestDocXBorder(borders?.LeftBorder ?? tableBorders?.LeftBorder, isExclude);
        Right = new QuestDocXBorder(borders?.RightBorder ?? tableBorders?.RightBorder, isExclude);
    }
}
