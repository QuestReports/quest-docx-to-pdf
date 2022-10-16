using DocumentFormat.OpenXml.Wordprocessing;
using JetBrains.Annotations;
using QuestReports.Converters.DocXToPdf.Extensions;
using QuestPDF.Fluent;
using QuestReports.Converters.DocXToPdf.CommonAssets;
using QuestReports.Converters.DocXToPdf.ElementHelpers;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;

namespace QuestReports.Converters.DocXToPdf.DescriptorResolvers;

[UsedImplicitly]
public class TableDescriptorResolver
{
    private readonly TableHelper _tableHelper;

    public TableDescriptorResolver(TableHelper tableHelper)
        => _tableHelper = tableHelper;

    public TableDescriptor Resolve(
        TableDescriptor descriptor,
        Table table,
        Document document,
        DocXQuestConversionOptions options)
    {
        var grid = SetupTableColumns(table, descriptor);
        var rows = table.GetRows();
        var tableMap = new uint[rows.Length, grid];
        uint rowCounter = 1; // initial row number in PDF table (always starts with 1)
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            uint buffer = 0;
            var cellIndex = 0;
            uint columnCounter = 1; // initial column number in PDF table (always starts with 1)
            var height = _tableHelper.CalculateTableRowHeight(rows[rowIndex]);
            var cells = rows[rowIndex].GetCells();
            for (var virtualCellIndex = 0; virtualCellIndex < grid; virtualCellIndex++)
            {
                if (tableMap[rowIndex, virtualCellIndex] > 0)
                {
                    if (tableMap[rowIndex, virtualCellIndex] > 1)
                    {
                        buffer = tableMap[rowIndex, virtualCellIndex];
                        cellIndex++;
                    }
                    else
                    {
                        cellIndex++;
                        columnCounter++;
                    }

                    continue;
                }

                if (buffer > 0)
                {
                    buffer--;
                    columnCounter++;
                    continue;
                }

                var cellProps = cells[cellIndex].TableCellProperties;
                var horizontalMerge = cellProps.GetHorizontalMergeFromGridSpan();
                buffer = horizontalMerge - 1;
                var verticalMerge = cellProps.IsStartOfVerticalMerge()
                    ? _tableHelper.ParseVerticalMerge(rows, rowIndex, virtualCellIndex)
                    : 1; // if there is no merge use 1 to represent it (1 cell takes 1 row)

                for (var rowComp = 1; rowComp < verticalMerge; rowComp++)
                    tableMap[rowIndex + rowComp, virtualCellIndex] = horizontalMerge;

                var inheritance = cells[cellIndex].ContainsTable();
                var cellAlignment = cellProps.GetVerticalCellAlignmentOrDefault();
                var cellColor = cellProps.GetCellFillHexColorOrDefault();
                var cellMargins = new QuestDocXCellMargins(cellProps?.TableCellMargin, options, inheritance);
                var cellBorders = new QuestDocXCellBorders(cellProps?.TableCellBorders, table.GetTableBorders());

                var index = cellIndex;
                descriptor
                    .Cell()
                    .Row(rowCounter)
                    .Column(columnCounter)
                    .ColumnSpan(horizontalMerge)
                    .RowSpan(verticalMerge)
                    .ApplyCellBorders(cellBorders)
                    .Background(cellColor)
                    .MinHeight(height)
                    .VerticallyAlign(cellAlignment)
                    .ApplyPaddings(cellMargins)
                    .ShowEntire()
                    .Column(
                        columnDescriptor => ResolversProvider.ColumnDescriptorResolver.Resolve(
                            columnDescriptor,
                            document,
                            cells[index].ChildElements,
                            options));

                columnCounter++;
                cellIndex++;
            }

            rowCounter++;
        }

        return descriptor;
    }

    private int SetupTableColumns(Table table, TableDescriptor descriptor)
    {
        var columns = _tableHelper.CalculateTableColumns(table);
        descriptor.ColumnsDefinition(
            tableColumnsDescriptor =>
            {
                foreach (var column in columns)
                    tableColumnsDescriptor.RelativeColumn(column);
            });
        return columns.Length;
    }
}
