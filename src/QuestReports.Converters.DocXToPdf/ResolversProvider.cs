using QuestReports.Converters.DocXToPdf.DescriptorResolvers;
using QuestReports.Converters.DocXToPdf.ElementHelpers;

namespace QuestReports.Converters.DocXToPdf;

public static class ResolversProvider
{
    public static TableDescriptorResolver TableDescriptorResolver { get; } = new(new TableHelper());

    public static ColumnDescriptorResolver ColumnDescriptorResolver { get; } = new();

    public static TextDescriptorResolver TextDescriptorResolver { get; } = new(new ParagraphHelper());
}
