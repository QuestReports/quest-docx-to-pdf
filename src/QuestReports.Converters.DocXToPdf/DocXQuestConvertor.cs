using DocumentFormat.OpenXml.Packaging;
using QuestPDF;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using QuestReports.Converters.DocXToPdf.Abstractions;
using QuestReports.Converters.DocXToPdf.ElementHelpers;

namespace QuestReports.Converters.DocXToPdf;

public class DocXQuestConvertor : IQuestConvertor
{
    private readonly PageCompositor _pageCompositor;

    public DocXQuestConvertor(PageCompositor pageCompositor)
    {
        _pageCompositor = pageCompositor;
        Settings.DocumentLayoutExceptionThreshold = DocXFormatConstants.DefaultMaxPagesNum;
    }

    public byte[] ConvertToPdf(WordprocessingDocument[] documents, Action<DocXQuestConversionOptions> options)
    {
        var newOptions = new DocXQuestConversionOptions();
        options(newOptions);
        return Convert(documents, newOptions).GeneratePdf();
    }

    public byte[] ConvertToPdf(WordprocessingDocument[] documents)
        => Convert(documents, new DocXQuestConversionOptions()).GeneratePdf();

    public IEnumerable<byte[]> ConvertToImages(
        WordprocessingDocument[] documents,
        Action<DocXQuestConversionOptions> options)
    {
        var newOptions = new DocXQuestConversionOptions();
        options(newOptions);
        return Convert(documents, newOptions).GenerateImages();
    }

    public IEnumerable<byte[]> ConvertToImages(WordprocessingDocument[] document)
        => Convert(document, new DocXQuestConversionOptions()).GenerateImages();

    private IDocument Convert(WordprocessingDocument[] documents, DocXQuestConversionOptions options)
        => new DocXQuestDocument(documents, options, _pageCompositor);
}
