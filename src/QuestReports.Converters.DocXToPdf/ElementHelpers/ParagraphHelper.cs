using DocumentFormat.OpenXml.Wordprocessing;
using QuestReports.Converters.DocXToPdf.Extensions;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using QuestReports.Converters.DocXToPdf.CommonAssets;

namespace QuestReports.Converters.DocXToPdf.ElementHelpers;

public class ParagraphHelper
{
    public void ResolveText(TextDescriptor descriptor, Text text, QuestDocXTextStyling questDocXTextStyling)
    {
        var replacement = text.Text == string.Empty ? " " : text.Text;
        questDocXTextStyling.WrapTextStyle(
            descriptor
                .Span(replacement)
                .WrapAnywhere());
    }

    public void ResolvePaging(
        TextDescriptor descriptor,
        FieldCode code,
        QuestDocXTextStyling questDocXTextStyling)
    {
        if (code.InnerText.Contains(DocXFormatConstants.PageNumberPlaceholder))
            questDocXTextStyling.WrapTextStyle(
                descriptor.CurrentPageNumber());
        else if (code.InnerText.Contains(DocXFormatConstants.TotalPagesNumberPlaceholder))
            questDocXTextStyling.WrapTextStyle(
                descriptor.TotalPages());
    }

    public void ResolveDrawing(DocumentFormat.OpenXml.Wordprocessing.Document document, Drawing drawing, IContainer descriptor)
    {
        using var picture = new QuestDocXPicture(drawing, document);
        descriptor
            .MaxHeight(picture.Size.Height.GetPxFromEmu())
            .MaxWidth(picture.Size.Width.GetPxFromEmu())
            .Image(picture.ImageStream, ImageScaling.Resize);
    }

    public void ResolveBreak(TextDescriptor descriptor, Break @break)
    {
        switch (@break.Type.Value)
        {
            case BreakValues.Column:
                descriptor.EmptyLine();
                break;
            case BreakValues.TextWrapping:
                descriptor.EmptyLine();
                break;
        }
    }
}
