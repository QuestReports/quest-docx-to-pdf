using DocumentFormat.OpenXml.Packaging;
using QuestPDF.Drawing;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using QuestReports.Converters.DocXToPdf.ElementHelpers;

namespace QuestReports.Converters.DocXToPdf;

public class DocXQuestDocument : IDocument
{
    private readonly WordprocessingDocument[] _documents;
    private readonly DocXQuestConversionOptions _options;
    private readonly PageCompositor _pageCompositor;

    public DocXQuestDocument(
        WordprocessingDocument[] documents,
        DocXQuestConversionOptions options,
        PageCompositor pageCompositor)
    {
        _documents = documents;
        _pageCompositor = pageCompositor;
        _options = options;
    }

    public DocumentMetadata GetMetadata() => _options.QuestDocumentMetadata;

    public void Compose(IDocumentContainer container)
    {
        foreach (var document in _documents)
            container.Page(page => Compose(page, document));
    }

    private void Compose(PageDescriptor page, WordprocessingDocument document)
    {
        var margins = _pageCompositor.SetupPageStyle(page, document, _options);
        page.Header().Column(columnDescriptor => _pageCompositor.ComposeHeader(columnDescriptor, margins.Top, document, _options));
        page.Content().Column(columnDescriptor => _pageCompositor.ComposeBody(columnDescriptor, document, _options));
        page.Footer().Column(columnDescriptor => _pageCompositor.ComposeFooter(columnDescriptor, margins.Bottom, document, _options));
    }
}
