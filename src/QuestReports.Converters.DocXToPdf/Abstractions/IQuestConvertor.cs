using DocumentFormat.OpenXml.Packaging;
using JetBrains.Annotations;

namespace QuestReports.Converters.DocXToPdf.Abstractions;

[PublicAPI]
public interface IQuestConvertor
{
    byte[] ConvertToPdf(WordprocessingDocument[] documents, Action<DocXQuestConversionOptions> options);

    byte[] ConvertToPdf(WordprocessingDocument[] documents);

    IEnumerable<byte[]> ConvertToImages(WordprocessingDocument[] documents, Action<DocXQuestConversionOptions> options);

    IEnumerable<byte[]> ConvertToImages(WordprocessingDocument[] documents);
}
