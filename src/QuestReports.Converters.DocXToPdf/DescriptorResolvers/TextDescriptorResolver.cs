using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using QuestReports.Converters.DocXToPdf.Extensions;
using QuestPDF.Fluent;
using QuestReports.Converters.DocXToPdf.CommonAssets;
using QuestReports.Converters.DocXToPdf.ElementHelpers;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;

namespace QuestReports.Converters.DocXToPdf.DescriptorResolvers;

public class TextDescriptorResolver
{
    private readonly ParagraphHelper _paragraphHelper;

    public TextDescriptorResolver(ParagraphHelper paragraphHelper)
        => _paragraphHelper = paragraphHelper;

    public void Resolve(
        TextDescriptor descriptor,
        Paragraph paragraph,
        Document document,
        DocXQuestConversionOptions options)
    {
        var prevWasCode = false;
        var alignment = paragraph.GetParagraphAlignmentOrDefault();
        var spacing = paragraph.GetParagraphSpacingOrDefault();
        descriptor.HorizontallyAlign(alignment);
        descriptor.ParagraphSpacing(spacing);
        Resolve(descriptor, paragraph, document, options, paragraph.ChildElements, ref prevWasCode);
    }

    private void Resolve(
        TextDescriptor descriptor,
        Paragraph paragraph,
        Document document,
        DocXQuestConversionOptions options,
        OpenXmlElementList elements,
        ref bool prevWasCode)
    {
        if (!elements.OfType<Run>().Any())
        {
            var fontSize = paragraph.GetFontSizeInParagraph() ?? options.DefaultBreakFontSize;
            descriptor
                .Element()
                .Text(string.Empty)
                .FontSize(fontSize);
        }

        foreach (var element in elements)
            if (element is Run run)
                ResolveRun(descriptor, document, run, ref prevWasCode, options);
    }

    private void ResolveRun(
        TextDescriptor descriptor,
        Document document,
        Run run,
        ref bool prevWasCode,
        DocXQuestConversionOptions options)
    {
        var runProperties = run.RunProperties;
        var text = run.Get<Text>();
        var drawing = run.Get<Drawing>();
        var @break = run.Get<Break>();
        var paging = run.Get<FieldCode>();
        if (paging is not null)
        {
            var textStyle = new QuestDocXTextStyling(runProperties, options);
            _paragraphHelper.ResolvePaging(descriptor, paging, textStyle);
            prevWasCode = true;
            return;
        }

        if (text is not null && !prevWasCode)
        {
            var textStyle = new QuestDocXTextStyling(runProperties, options);
            _paragraphHelper.ResolveText(descriptor, text, textStyle);
        }
        else if (drawing is not null)
        {
            _paragraphHelper.ResolveDrawing(document, drawing, descriptor.Element());
        }
        else if (@break is not null)
        {
            _paragraphHelper.ResolveBreak(descriptor, @break);
        }

        if (text is null)
            return;
        prevWasCode = false;
    }
}
