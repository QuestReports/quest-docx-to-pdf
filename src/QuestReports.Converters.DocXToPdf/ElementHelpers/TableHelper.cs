using DocumentFormat.OpenXml.Wordprocessing;
using JetBrains.Annotations;
using QuestReports.Converters.DocXToPdf.Extensions;

namespace QuestReports.Converters.DocXToPdf.ElementHelpers;

[UsedImplicitly]
public class TableHelper
{
    public uint ParseVerticalMerge(IReadOnlyList<TableRow> tableRows, int row, int column)
    {
        uint counter = 1;
        uint increment = 1;
        while (true)
        {
            if (tableRows.Count <= row + increment)
                break;
            var cells = tableRows[(int)(row + increment)].GetCells();
            if (column >= cells.Length)
                break;
            if (cells[column].TableCellProperties?.VerticalMerge == null)
                break;

            increment++;
            counter++;
        }

        return counter;
    }

    public float[] CalculateTableColumns(Table table)
    {
        var grid = table.GetTableGrid();
        var columns = grid!.GetGridColumns();
        return columns.Select(col => float.Parse(col.Width)).ToArray();
    }

    public float CalculateTableRowHeight(TableRow tableRow)
    {
        var height = 0f;
        var fontHeight = 0f;
        var rawHeight = tableRow.GetTableRowHeight();
        var rawMaxFont = tableRow.GetTableRowHeightByInnerContent();

        if (rawMaxFont is not null)
            fontHeight = rawMaxFont.GetPxFromPt();
        if (rawHeight is not null)
            height = rawHeight.GetPxFromDxa();
        return fontHeight > height ? fontHeight : height;
    }
}
