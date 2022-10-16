using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QuestReports.Converters.DocXToPdf.Extensions;
using QuestPDF.Infrastructure;

namespace QuestReports.Converters.DocXToPdf.CommonAssets;

public class QuestDocXPicture : IDisposable
{
    public Size Size { get; }

    public Stream ImageStream { get; } = new MemoryStream();

    public QuestDocXPicture(Drawing drawing, Document document)
    {
        var picture = drawing.GetPicture();
        var blip = picture?.GetBlip();
        if (blip is null)
            return;
        var img = GetImagePart(document, blip);
        var extent = drawing.GetExtent();
        Size = new Size(extent?.Cx ?? 0, extent?.Cy ?? 0);
        ImageStream = img!.GetStream();
    }

    private static ImagePart? GetImagePart(Document document, string partId)
    {
        var mainPart = document.GetMainPartById(partId);
        if (mainPart is ImagePart imageMainPart)
            return imageMainPart;
        var headerPart = document.GetHeaderPartById(partId);
        if (headerPart is ImagePart imageHeaderPart)
            return imageHeaderPart;
        var footerPart = document.GetFooterPartById(partId);
        if (footerPart is ImagePart footerFooterPart)
            return footerFooterPart;
        return null;
    }

    public void Dispose()
        => ImageStream.Dispose();
}
