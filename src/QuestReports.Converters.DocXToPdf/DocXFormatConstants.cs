namespace QuestReports.Converters.DocXToPdf;

public static class DocXFormatConstants
{
    public const float DxaScale = 20f;
    public const float PtFontScale = 2f;
    public const float EmuScale = 9525f;

    public const float DefaultStrokeScale = 0.2f;

    public const int DefaultMaxPagesNum = 1000;

    public const string PageNumberPlaceholder = @" PAGE   \* MERGEFORMAT ";
    public const string TotalPagesNumberPlaceholder = @" NUMPAGES   \* MERGEFORMAT ";

    // TODO: There is a requirement to add minified fonts to make PDF more light weight
    public static readonly IReadOnlyDictionary<string, string> Fonts = new Dictionary<string, string>
    {
        // {"Calibri", "Calibri"}, TODO: There is no support for linux currently
        {"Arial", "Arial"},
        {"Cambria", "Cambria"}, {"Candara", "Candara"},
        {"Comic Sans MS", "Comic Sans MS"}, {"Consolas", "Consolas"},
        {"Corbel", "Corbel"}, {"Courier", "Courier"},
        {"Courier New", "Courier New"}, {"Georgia", "Georgia"},
        {"Impact", "Impact"}, {"Lucida Console", "Lucida Console"},
        {"Segoe SD", "Segoe SD"}, {"Segoe UI", "Segoe UI"},
        {"Tahoma", "Tahoma"}, {"Times New Roman", "Times New Roman"},
        {"Times Roman", "Times Roman"}, {"Trebuchet MS", "Trebuchet MS"},
        {"Verdana", "Verdana"}
    };
}
