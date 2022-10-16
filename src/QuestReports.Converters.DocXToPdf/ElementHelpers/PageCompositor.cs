using DocumentFormat.OpenXml.Packaging;
using JetBrains.Annotations;
using QuestReports.Converters.DocXToPdf.Extensions;
using QuestPDF.Fluent;
using QuestReports.Converters.DocXToPdf.CommonAssets;

namespace QuestReports.Converters.DocXToPdf.ElementHelpers;

[UsedImplicitly]
public class PageCompositor
{
    public QuestDocXMargins SetupPageStyle(
        PageDescriptor page,
        WordprocessingDocument document,
        DocXQuestConversionOptions options)
    {
        var pageProps = new QuestDocXPageProperties(document, options);
        pageProps.WrapPageStyle(page);
        return pageProps.Margins;
    }

    public ColumnDescriptor ComposeHeader(
        ColumnDescriptor column,
        float margin,
        WordprocessingDocument document,
        DocXQuestConversionOptions options)
    {
        if (!document.GetHeadersParts().Any())
        {
            column.Item().Height(margin);
            return column;
        }

        column
            .Item()
            .MinHeight(margin)
            .AlignMiddle()
            .Column(
            columnDescriptor =>
            {
                foreach (var headerPart in document.GetHeadersParts())
                    ResolversProvider.ColumnDescriptorResolver.Resolve(
                        columnDescriptor,
                        document.GetDocument(),
                        headerPart.Header.ChildElements,
                        options);
            });

        return column;
    }

    public ColumnDescriptor ComposeBody(
        ColumnDescriptor column,
        WordprocessingDocument document,
        DocXQuestConversionOptions options)
    {
        column
            .Item()
            .Column(
            columnDescriptor => ResolversProvider.ColumnDescriptorResolver.Resolve(
                columnDescriptor,
                document.GetDocument(),
                document.GetDocumentBodyChildElements(),
                options));

        return column;
    }

    public ColumnDescriptor ComposeFooter(
        ColumnDescriptor column,
        float margin,
        WordprocessingDocument document,
        DocXQuestConversionOptions options)
    {
        if (!document.GetFootersParts().Any())
        {
            column
                .Item()
                .Height(margin);
            return column;
        }

        column
            .Item()
            .MinHeight(margin)
            .Column(
                columnDescriptor =>
                {
                    foreach (var footerPart in document.GetFootersParts())
                        ResolversProvider.ColumnDescriptorResolver.Resolve(
                            columnDescriptor,
                            document.GetDocument(),
                            footerPart.Footer.ChildElements,
                            options);
                });
        return column;
    }
}
