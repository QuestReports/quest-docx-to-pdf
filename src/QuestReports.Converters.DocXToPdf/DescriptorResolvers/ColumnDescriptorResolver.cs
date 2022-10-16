using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using QuestReports.Converters.DocXToPdf.Extensions;
using QuestPDF.Fluent;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;

namespace QuestReports.Converters.DocXToPdf.DescriptorResolvers;

public class ColumnDescriptorResolver
{
    public void Resolve(
        ColumnDescriptor descriptor,
        Document document,
        OpenXmlElementList elements,
        DocXQuestConversionOptions options)
    {
        foreach (var element in elements)
            switch (element)
            {
                case Paragraph paragraph:
                    ResolveParagraph(descriptor, document, paragraph, options);
                    break;
                case Table table:
                    ResolveTable(descriptor, document, table, options);
                    break;
            }
    }

    private static void ResolveParagraph(
        ColumnDescriptor descriptor,
        Document document,
        Paragraph paragraph,
        DocXQuestConversionOptions options)
    {
        descriptor
            .Item()
            .ExtendHorizontal()
            .ResolveTextDirection(paragraph.GetParagraphTextDirection())
            .Text(
                textDescriptor =>
                {
                    var left = paragraph.GetParagraphLeftIndentation();
                    textDescriptor
                        .Element()
                        .MinWidth(left);

                    ResolversProvider.TextDescriptorResolver.Resolve(textDescriptor, paragraph, document, options);

                    var right = paragraph.GetParagraphRightIndentation();
                    textDescriptor
                        .Element()
                        .MinWidth(right);
                });
        if (paragraph.ContainsPageBreak())
            descriptor
                .Item()
                .PageBreak();
    }

    private static void ResolveTable(
        ColumnDescriptor descriptor,
        Document document,
        Table table,
        DocXQuestConversionOptions options)
    {
        descriptor
            .Item()
            .ExtendHorizontal()
            .Table(
                tableDescriptor
                    => ResolversProvider
                        .TableDescriptorResolver
                        .Resolve(tableDescriptor, table, document, options));
    }
}
