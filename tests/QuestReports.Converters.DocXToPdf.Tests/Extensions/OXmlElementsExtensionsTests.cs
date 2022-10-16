using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using QuestReports.Converters.DocXToPdf.Extensions;
using QuestPDF.Helpers;
using Xunit;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using GridColumn = DocumentFormat.OpenXml.Wordprocessing.GridColumn;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Picture = DocumentFormat.OpenXml.Drawing.Pictures.Picture;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableCellProperties = DocumentFormat.OpenXml.Wordprocessing.TableCellProperties;
using TableGrid = DocumentFormat.OpenXml.Wordprocessing.TableGrid;
using TableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace QuestReports.Converters.DocXToPdf.Tests.Extensions;

public class OXmlElementsExtensionsTests
{
    #region GetParagraphRightIndentation

    [Theory]
    [InlineData(-1001.11f)]
    [InlineData(-1000f)]
    [InlineData(-1f)]
    [InlineData(-0.5f)]
    [InlineData(0f)]
    [InlineData(0.5f)]
    [InlineData(1f)]
    [InlineData(1000f)]
    [InlineData(1001.11f)]
    public void GetParagraphRightIndentation_CorrectValue_CorrectResult(float value)
    {
        // Arrange
        var paragraph = new Paragraph();
        var properties = new ParagraphProperties();
        var indentation = new Indentation
        {
            Right = value.ToString(CultureInfo.CurrentCulture)
        };
        properties.AppendChild(indentation);
        paragraph.AppendChild(properties);

        var expected = value / DocXFormatConstants.DxaScale;

        // Act
        var actual = paragraph.GetParagraphRightIndentation();

        // Assert
        actual.Should().Be(expected);
    }

    [Fact]
    public void GetParagraphRightIndentation_NullValue_ZeroResult()
    {
        // Arrange
        var paragraph = new Paragraph();
        var properties = new ParagraphProperties();
        var indentation = new Indentation();
        properties.AppendChild(indentation);
        paragraph.AppendChild(properties);

        var expected = 0;

        // Act
        var actual = paragraph.GetParagraphRightIndentation();

        // Assert
        actual.Should().Be(expected);
    }

    #endregion

    #region GetParagraphLeftIndentation

    [Theory]
    [InlineData(-1001.11f)]
    [InlineData(-1000f)]
    [InlineData(-1f)]
    [InlineData(-0.5f)]
    [InlineData(0f)]
    [InlineData(0.5f)]
    [InlineData(1f)]
    [InlineData(1000f)]
    [InlineData(1001.11f)]
    public void GetParagraphLeftIndentation_CorrectValue_CorrectResult(float value)
    {
        // Arrange
        var paragraph = new Paragraph();
        var properties = new ParagraphProperties();
        var indentation = new Indentation
        {
            Left = value.ToString(CultureInfo.CurrentCulture)
        };
        properties.AppendChild(indentation);
        paragraph.AppendChild(properties);

        var expected = value / DocXFormatConstants.DxaScale;

        // Act
        var actual = paragraph.GetParagraphLeftIndentation();

        // Assert
        actual.Should().Be(expected);
    }

    [Theory]
    [InlineData(-1001.11f)]
    [InlineData(-1000f)]
    [InlineData(-1f)]
    [InlineData(-0.5f)]
    [InlineData(0f)]
    [InlineData(0.5f)]
    [InlineData(1f)]
    [InlineData(1000f)]
    [InlineData(1001.11f)]
    public void GetParagraphLeftIndentation_CorrectFirstLineValue_CorrectResult(float value)
    {
        // Arrange
        var paragraph = new Paragraph();
        var properties = new ParagraphProperties();
        var indentation = new Indentation
        {
            FirstLine = value.ToString(CultureInfo.CurrentCulture)
        };
        properties.AppendChild(indentation);
        paragraph.AppendChild(properties);

        var expected = value / DocXFormatConstants.DxaScale;

        // Act
        var actual = paragraph.GetParagraphLeftIndentation();

        // Assert
        actual.Should().Be(expected);
    }

    [Fact]
    public void GetParagraphLeftIndentation_NullValue_ZeroResult()
    {
        // Arrange
        var paragraph = new Paragraph();
        var properties = new ParagraphProperties();
        var indentation = new Indentation();
        properties.AppendChild(indentation);
        paragraph.AppendChild(properties);

        var expected = 0;

        // Act
        var actual = paragraph.GetParagraphLeftIndentation();

        // Assert
        actual.Should().Be(expected);
    }

    #endregion

    #region GetParagraphTextDirection

    [Theory]
    [InlineData(TextDirectionValues.LefToRightTopToBottom)]
    [InlineData(TextDirectionValues.LeftToRightTopToBottom2010)]
    [InlineData(TextDirectionValues.TopToBottomRightToLeft)]
    [InlineData(TextDirectionValues.TopToBottomRightToLeft2010)]
    [InlineData(TextDirectionValues.BottomToTopLeftToRight)]
    [InlineData(TextDirectionValues.BottomToTopLeftToRight2010)]
    [InlineData(TextDirectionValues.LefttoRightTopToBottomRotated)]
    [InlineData(TextDirectionValues.LeftToRightTopToBottomRotated2010)]
    [InlineData(TextDirectionValues.TopToBottomRightToLeftRotated)]
    [InlineData(TextDirectionValues.TopToBottomRightToLeftRotated2010)]
    [InlineData(TextDirectionValues.TopToBottomLeftToRightRotated)]
    [InlineData(TextDirectionValues.TopToBottomLeftToRightRotated2010)]
    [InlineData(null)]
    public void GetParagraphTextDirection_CorrectValue_CorrectResult(TextDirectionValues? value)
    {
        // Arrange
        var paragraph = new Paragraph();
        var properties = new ParagraphProperties();
        var indentation = new TextDirection
        {
            Val = value
        };
        properties.AppendChild(indentation);
        paragraph.AppendChild(properties);

        // Act
        var actual = paragraph.GetParagraphTextDirection()?.Value;

        // Assert
        actual.Should().Be(value);
    }

    #endregion

    #region GetRows

    [Fact]
    public void GetRows_CorrectValue_CorrectResult()
    {
        // Arrange
        var table = new Table();
        var tableRows = new List<TableRow>();
        for (var i = 0; i < 10; i++)
        {
            var tableRow = new TableRow();
            tableRows.Add(tableRow);
            table.AppendChild(tableRow);
        }

        // Act
        var actual = table.GetRows();

        // Assert
        actual.Should().BeEquivalentTo(tableRows);
    }

    #endregion

    #region GetCells

    [Fact]
    public void GetCells_CorrectValue_CorrectResult()
    {
        // Arrange
        var tableRow = new TableRow();
        var tableCells = new List<TableCell>();
        for (var i = 0; i < 10; i++)
        {
            var tableCell = new TableCell();
            tableCells.Add(tableCell);
            tableRow.AppendChild(tableCell);
        }

        // Act
        var actual = tableRow.GetCells();

        // Assert
        actual.Should().BeEquivalentTo(tableCells);
    }

    #endregion

    #region GetHorizontalMergeFromGridSpan

    [Theory]
    [InlineData(1)]
    [InlineData(10)]
    [InlineData(100)]
    public void GetHorizontalMergeFromGridSpan_CorrectValue_CorrectResult(int value)
    {
        // Arrange
        var tableCellProperties = new TableCellProperties();
        var gridSpan = new GridSpan
        {
            Val = value
        };
        tableCellProperties.AppendChild(gridSpan);

        // Act
        var actual = tableCellProperties.GetHorizontalMergeFromGridSpan();

        // Assert
        actual.Should().Be((uint)value);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(null)]
    public void GetHorizontalMergeFromGridSpan_ZeroOrNullValue_OneAsResult(int? value)
    {
        // Arrange
        var tableCellProperties = new TableCellProperties();
        var gridSpan = new GridSpan
        {
            Val = value
        };
        tableCellProperties.AppendChild(gridSpan);

        // Act
        var actual = tableCellProperties.GetHorizontalMergeFromGridSpan();

        // Assert
        actual.Should().Be(1);
    }

    #endregion

    #region IsStartOfVerticalMerge

    [Fact]
    public void IsStartOfVerticalMerge_RestartValue_ReturnsTrue()
    {
        // Arrange
        var tableCellProperties = new TableCellProperties();
        var verticalMerge = new VerticalMerge
        {
            Val = MergedCellValues.Restart
        };
        tableCellProperties.AppendChild(verticalMerge);

        // Act
        var actual = tableCellProperties.IsStartOfVerticalMerge();

        // Assert
        actual.Should().Be(true);
    }

    [Theory]
    [InlineData(MergedCellValues.Continue)]
    [InlineData(null)]
    public void IsStartOfVerticalMerge_ContinueOrNullValue_ReturnsFalse(MergedCellValues? value)
    {
        // Arrange
        var tableCellProperties = new TableCellProperties();
        var verticalMerge = new VerticalMerge
        {
            Val = value
        };
        tableCellProperties.AppendChild(verticalMerge);

        // Act
        var actual = tableCellProperties.IsStartOfVerticalMerge();

        // Assert
        actual.Should().Be(false);
    }

    #endregion

    #region GetVerticalCellAlignmentOrDefault

    [Theory]
    [InlineData(TableVerticalAlignmentValues.Bottom)]
    [InlineData(TableVerticalAlignmentValues.Top)]
    [InlineData(TableVerticalAlignmentValues.Center)]
    public void GetVerticalCellAlignmentOrDefault_CorrectValue_CorrectResult(TableVerticalAlignmentValues value)
    {
        // Arrange
        var tableCellProperties = new TableCellProperties();
        var verticalAlignment = new TableCellVerticalAlignment
        {
            Val = value
        };
        tableCellProperties.AppendChild(verticalAlignment);

        // Act
        var actual = tableCellProperties.GetVerticalCellAlignmentOrDefault();

        // Assert
        actual.Should().Be(value);
    }

    [Fact]
    public void GetVerticalCellAlignmentOrDefault_NullValue_DefaultResult()
    {
        // Arrange
        var tableCellProperties = new TableCellProperties();
        var verticalAlignment = new TableCellVerticalAlignment();
        tableCellProperties.AppendChild(verticalAlignment);

        // Act
        var actual = tableCellProperties.GetVerticalCellAlignmentOrDefault();

        // Assert
        actual.Should().Be(TableVerticalAlignmentValues.Center);
    }

    #endregion

    #region GetTableBorders

    [Fact]
    public void GetTableBorders_CorrectValue_CorrectResult()
    {
        // Arrange
        var table = new Table();
        var tableProperties = new TableProperties();
        var tableBorders = new TableBorders();
        tableProperties.AppendChild(tableBorders);
        table.AppendChild(tableProperties);

        // Act
        var actual = table.GetTableBorders();

        // Assert
        actual.Should().BeEquivalentTo(tableBorders);
    }

    [Fact]
    public void GetTableBorders_NullProperties_NullResult()
    {
        // Arrange
        var table = new Table();
        var tableProperties = new TableProperties();
        table.AppendChild(tableProperties);

        // Act
        var actual = table.GetTableBorders();

        // Assert
        actual.Should().BeNull();
    }

    #endregion

    #region GetCellFillHexColorOrDefault

    [Theory]
    [InlineData(Colors.Cyan.Accent1)]
    [InlineData(Colors.Black)]
    public void GetCellFillHexColorOrDefault_CorrectValue_CorrectResult(string value)
    {
        // Arrange
        var tableCellProperties = new TableCellProperties();
        var shading = new Shading
        {
            Fill = value
        };
        tableCellProperties.AppendChild(shading);

        // Act
        var actual = tableCellProperties.GetCellFillHexColorOrDefault();

        // Assert
        actual.Should().Be(value);
    }

    [Fact]
    public void GetCellFillHexColorOrDefault_NullValue_WhiteAsResult()
    {
        // Arrange
        var tableCellProperties = new TableCellProperties();
        var shading = new Shading();
        tableCellProperties.AppendChild(shading);

        // Act
        var actual = tableCellProperties.GetCellFillHexColorOrDefault();

        // Assert
        actual.Should().Be(Colors.White);
    }

    #endregion

    #region GetTableRowHeight

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(10)]
    public void GetTableRowHeight_CorrectValue_CorrectResult(uint value)
    {
        // Arrange
        var tableRow = new TableRow();
        var tableRowProps = new TableRowProperties();
        var tableRowHeight = new TableRowHeight()
        {
            Val = value
        };
        tableRowProps.AppendChild(tableRowHeight);
        tableRow.AppendChild(tableRowProps);

        // Act
        var actual = tableRow.GetTableRowHeight();

        // Assert
        actual.Should().Be(value);
    }

    [Fact]
    public void GetTableRowHeight_NullValue_NullResult()
    {
        // Arrange
        var tableRow = new TableRow();

        // Act
        var actual = tableRow.GetTableRowHeight();

        // Assert
        actual.Should().BeNull();
    }

    #endregion

    #region GetTableRowHeightByInnerContent

    [Theory]
    [InlineData("0")]
    [InlineData("1")]
    [InlineData("10")]
    [InlineData("12.4")]
    [InlineData("153.34")]
    public void GetTableRowHeightByInnerContent_CorrectValue_CorrectResult(string value)
    {
        // Arrange
        var tableRow = new TableRow();
        var runProperties = new RunProperties()
        {
            FontSize = new FontSize()
            {
                Val = value
            }
        };
        tableRow.AppendChild(runProperties);

        // Act
        var actual = tableRow.GetTableRowHeightByInnerContent();

        // Assert
        actual.Should().Be(value);
    }

    [Fact]
    public void GetTableRowHeightByInnerContent_NullValue_NullResult()
    {
        // Arrange
        var tableRow = new TableRow();
        var runProperties = new RunProperties();
        tableRow.AppendChild(runProperties);

        // Act
        var actual = tableRow.GetTableRowHeightByInnerContent();

        // Assert
        actual.Should().BeNull();
    }

    #endregion

    #region GetTableGrid

    [Fact]
    public void GetTableGrid_CorrectValue_CorrectResult()
    {
        // Arrange
        var table = new Table();
        var tableGrid = new TableGrid();
        table.AppendChild(tableGrid);

        // Act
        var actual = table.GetTableGrid();

        // Assert
        actual.Should().BeEquivalentTo(tableGrid);
    }

    [Fact]
    public void GetTableGrid_NullValue_NullResult()
    {
        // Arrange
        var table = new Table();
        // Act
        var actual = table.GetTableGrid();

        // Assert
        actual.Should().BeNull();
    }

    #endregion

    #region GetGridColumns

    [Fact]
    public void GetGridColumns_CorrectValue_CorrectResult()
    {
        // Arrange
        var tableGrid = new TableGrid();
        var expectedGripColumns = new List<GridColumn>();
        for (var i = 0; i < 10; i++)
        {
            var gridColumn = new GridColumn{Width = i.ToString()};
            tableGrid.AppendChild(gridColumn);
            expectedGripColumns.Add(gridColumn);
        }

        // Act
        var actual = tableGrid.GetGridColumns();

        // Assert
        actual.Should().BeEquivalentTo(expectedGripColumns.ToArray());
    }

    [Fact]
    public void GetGridColumns_NullValue_EmptyResult()
    {
        // Arrange
        var tableGrid = new TableGrid();

        // Act
        var actual = tableGrid.GetGridColumns();

        // Assert
        actual.Should().BeEmpty();
    }

    #endregion

    #region GetText

    [Fact]
    public void GetText_CorrectValue_CorrectResult()
    {
        // Arrange
        var run = new Run();
        var text = new Text();
        run.AppendChild(text);

        // Act
        var actual = run.Get<Text>();

        // Assert
        actual.Should().BeEquivalentTo(text);
    }

    [Fact]
    public void GetText_NullValue_NullResult()
    {
        // Arrange
        var run = new Run();

        // Act
        var actual = run.Get<Text>();

        // Assert
        actual.Should().BeNull();
    }

    #endregion

    #region GetDrawing

    [Fact]
    public void GetDrawing_CorrectValue_CorrectResult()
    {
        // Arrange
        var run = new Run();
        var drawing = new Drawing();
        run.AppendChild(drawing);

        // Act
        var actual = run.Get<Drawing>();

        // Assert
        actual.Should().BeEquivalentTo(drawing);
    }

    [Fact]
    public void GetDrawing_NullValue_NullResult()
    {
        // Arrange
        var run = new Run();

        // Act
        var actual = run.Get<Drawing>();

        // Assert
        actual.Should().BeNull();
    }

    #endregion

    #region GetBreak

    [Fact]
    public void GetBreak_CorrectValue_CorrectResult()
    {
        // Arrange
        var run = new Run();
        var @break = new Break();
        run.AppendChild(@break);

        // Act
        var actual = run.Get<Break>();

        // Assert
        actual.Should().BeEquivalentTo(@break);
    }

    [Fact]
    public void GetBreak_NullValue_NullResult()
    {
        // Arrange
        var run = new Run();

        // Act
        var actual = run.Get<Break>();

        // Assert
        actual.Should().BeNull();
    }

    #endregion

    #region GetFieldCode

    [Fact]
    public void GetFieldCode_CorrectValue_CorrectResult()
    {
        // Arrange
        var run = new Run();
        var fieldCode = new FieldCode();
        run.AppendChild(fieldCode);

        // Act
        var actual = run.Get<FieldCode>();

        // Assert
        actual.Should().BeEquivalentTo(fieldCode);
    }

    [Fact]
    public void GetFieldCode_NullValue_NullResult()
    {
        // Arrange
        var run = new Run();

        // Act
        var actual = run.Get<FieldCode>();

        // Assert
        actual.Should().BeNull();
    }

    #endregion

    #region GetPicture

    [Fact]
    public void GetPicture_CorrectValue_CorrectResult()
    {
        // Arrange
        var drawing = new Drawing();
        var inline = new Inline();
        var graphic = new Graphic();
        var graphicData = new GraphicData();
        var picture = new Picture();
        graphicData.AppendChild(picture);
        graphic.AppendChild(graphicData);
        inline.AppendChild(graphic);
        drawing.AppendChild(inline);

        // Act
        var actual = drawing.GetPicture();

        // Assert
        actual.Should().BeEquivalentTo(picture);
    }

    [Fact]
    public void GetPicture_NullValue_NullResult()
    {
        // Arrange
        var drawing = new Drawing();

        // Act
        var actual = drawing.GetPicture();

        // Assert
        actual.Should().BeNull();
    }

    #endregion

    #region GetBlip

    [Theory]
    [InlineData("1b")]
    [InlineData("10b")]
    [InlineData("3333b")]
    public void GetBlip_CorrectValue_CorrectResult(string value)
    {
        // Arrange
        var picture = new Picture
        {
            BlipFill = new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill
            {
                Blip = new Blip
                {
                    Embed = value
                }
            }
        };

        // Act
        var actual = picture.GetBlip();

        // Assert
        actual.Should().Be(value);
    }

    [Fact]
    public void GetBlip_NullValue_NullResult()
    {
        // Arrange
        var picture = new Picture();

        // Act
        var actual = picture.GetBlip();

        // Assert
        actual.Should().BeNull();
    }

    #endregion

    #region GetExtent

    [Fact]
    public void GetExtent_CorrectValue_CorrectResult()
    {
        // Arrange
        var drawing = new Drawing();
        var extent = new Extent();
        drawing.AppendChild(extent);

        // Act
        var actual = drawing.GetExtent();

        // Assert
        actual.Should().BeEquivalentTo(extent);
    }

    [Fact]
    public void GetExtent_NullValue_NullResult()
    {
        // Arrange
        var drawing = new Drawing();

        // Act
        var actual = drawing.GetExtent();

        // Assert
        actual.Should().BeNull();
    }

    #endregion

    #region GetMainPartById

    [Fact]
    public void GetMainPartById_CorrectValue_CorrectResult()
    {
        // Arrange
        const string partId = "testId";
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainDocumentPart = document.AddMainDocumentPart();
        mainDocumentPart.Document = new Document();
        mainDocumentPart.Document.MainDocumentPart.AddNewPart<ChartPart>(partId);

        // Act
        var actual = document.MainDocumentPart.Document.GetMainPartById(partId);

        // Assert
        actual.Should().BeOfType<ChartPart>();
    }

    [Fact]
    public void GetMainPartById_NullValue_ThrowEx()
    {
        // Arrange
        const string partId = "testId";
        using var stream = new MemoryStream();
        var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainDocumentPart = document.AddMainDocumentPart();
        mainDocumentPart.Document = new Document();

        // Act
        var actual = () => document.MainDocumentPart.Document.GetMainPartById(partId);

        // Assert
        actual.Should().Throw<ArgumentOutOfRangeException>();
    }

    #endregion

    #region GetHeaderPartById

    [Fact]
    public void GetHeaderPartById_CorrectValue_CorrectResult()
    {
        // Arrange
        const string partId = "testId";
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainDocumentPart = document.AddMainDocumentPart();
        mainDocumentPart.Document = new Document();
        mainDocumentPart.Document.MainDocumentPart.AddNewPart<HeaderPart>().AddNewPart<ChartPart>(partId);

        // Act
        var actual = document.MainDocumentPart.Document.GetHeaderPartById(partId);

        // Assert
        actual.Should().BeOfType<ChartPart>();
    }

    [Fact]
    public void GetHeaderPartById_NullUpperValue_NullResult()
    {
        // Arrange
        const string partId = "testId";
        using var stream = new MemoryStream();
        var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainDocumentPart = document.AddMainDocumentPart();
        mainDocumentPart.Document = new Document();

        // Act
        var actual = document.MainDocumentPart.Document.GetHeaderPartById(partId);

        // Assert
        actual.Should().BeNull();
    }

    [Fact]
    public void GetHeaderPartById_NullValue_ThrowEx()
    {
        // Arrange
        const string partId = "testId";
        using var stream = new MemoryStream();
        var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainDocumentPart = document.AddMainDocumentPart();
        mainDocumentPart.Document = new Document();
        mainDocumentPart.Document.MainDocumentPart.AddNewPart<HeaderPart>();

        // Act
        var actual = () => document.MainDocumentPart.Document.GetHeaderPartById(partId);

        // Assert
        actual.Should().Throw<ArgumentOutOfRangeException>();
    }

    #endregion

    #region GetFooterPartById

    [Fact]
    public void GetFooterPartById_CorrectValue_CorrectResult()
    {
        // Arrange
        const string partId = "testId";
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainDocumentPart = document.AddMainDocumentPart();
        mainDocumentPart.Document = new Document();
        mainDocumentPart.Document.MainDocumentPart.AddNewPart<FooterPart>().AddNewPart<ChartPart>(partId);

        // Act
        var actual = document.MainDocumentPart.Document.GetFooterPartById(partId);

        // Assert
        actual.Should().BeOfType<ChartPart>();
    }

    [Fact]
    public void GetFooterPartById_NullUpperValue_NullResult()
    {
        // Arrange
        const string partId = "testId";
        using var stream = new MemoryStream();
        var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainDocumentPart = document.AddMainDocumentPart();
        mainDocumentPart.Document = new Document();

        // Act
        var actual = document.MainDocumentPart.Document.GetFooterPartById(partId);

        // Assert
        actual.Should().BeNull();
    }

    [Fact]
    public void GetFooterPartById_NullValue_ThrowEx()
    {
        // Arrange
        const string partId = "testId";
        using var stream = new MemoryStream();
        var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainDocumentPart = document.AddMainDocumentPart();
        mainDocumentPart.Document = new Document();
        mainDocumentPart.Document.MainDocumentPart.AddNewPart<FooterPart>();

        // Act
        var actual = () => document.MainDocumentPart.Document.GetFooterPartById(partId);

        // Assert
        actual.Should().Throw<ArgumentOutOfRangeException>();
    }

    #endregion

    #region GetFontSizeInParagraph

    [Theory]
    [InlineData(1)]
    [InlineData(10)]
    [InlineData(13.4f)]
    [InlineData(null)]
    public void GetFontSizeInParagraph_CorrectValue_CorrectResult(float? value)
    {
        // Arrange
        var paragraph = new Paragraph();
        var paragraphProps = new ParagraphProperties();
        var paragraphMarkRunProperties = new ParagraphMarkRunProperties();
        var fontSize = new FontSize
        {
            Val = value?.ToString(CultureInfo.CurrentCulture)
        };
        paragraphMarkRunProperties.AppendChild(fontSize);
        paragraphProps.AppendChild(paragraphMarkRunProperties);
        paragraph.AppendChild(paragraphProps);

        var expected = value / DocXFormatConstants.PtFontScale;

        // Act
        var actual = paragraph.GetFontSizeInParagraph();

        // Assert
        actual.Should().Be(expected);
    }

    #endregion

    #region ContainsTable

    [Fact]
    public void ContainsTable_Contains_ReturnsTrue()
    {
        // Arrange
        var tableCell = new TableCell();
        var table = new Table();
        tableCell.AppendChild(table);

        // Act
        var actual = tableCell.ContainsTable();

        // Assert
        actual.Should().BeTrue();
    }

    [Fact]
    public void ContainsTable_Empty_ReturnsFalse()
    {
        // Arrange
        var tableCell = new TableCell();

        // Act
        var actual = tableCell.ContainsTable();

        // Assert
        actual.Should().BeFalse();
    }

    #endregion

    #region ContainsPageBreak

    [Fact]
    public void ContainsPageBreak_Contains_ReturnsTrue()
    {
        // Arrange
        var paragraph = new Paragraph();
        var pageBreak = new Break
        {
            Type = BreakValues.Page
        };
        paragraph.AppendChild(pageBreak);

        // Act
        var actual = paragraph.ContainsPageBreak();

        // Assert
        actual.Should().BeTrue();
    }

    [Fact]
    public void ContainsPageBreak_Empty_ReturnsFalse()
    {
        // Arrange
        var paragraph = new Paragraph();

        // Act
        var actual = paragraph.ContainsPageBreak();

        // Assert
        actual.Should().BeFalse();
    }

    [Fact]
    public void ContainsPageBreak_ContainsNotPageBreak_ReturnsFalse()
    {
        // Arrange
        var paragraph = new Paragraph();
        var pageBreak = new Break
        {
            Type = BreakValues.TextWrapping
        };
        paragraph.AppendChild(pageBreak);

        // Act
        var actual = paragraph.ContainsPageBreak();

        // Assert
        actual.Should().BeFalse();
    }

    #endregion

    #region GetDocument

    [Fact]
    public void GetDocument_CorrectValue_CorrectResult()
    {
        // Arrange
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainDocumentPart = document.AddMainDocumentPart();
        mainDocumentPart.Document = new Document();

        // Act
        var actual = document.GetDocument();

        // Assert
        actual.Should().BeEquivalentTo(mainDocumentPart.Document);
    }

    [Fact]
    public void GetDocument_NullValue_ThrowEx()
    {
        // Arrange
        using var stream = new MemoryStream();
        var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);

        // Act
        var actual = () => document.GetDocument();

        // Assert
        actual.Should().Throw<NullReferenceException>();
    }

    #endregion

    #region GetDocumentBodyChildElements

    [Fact]
    public void GetDocumentBodyChildElements_CorrectValue_CorrectResult()
    {
        // Arrange
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainDocumentPart = document.AddMainDocumentPart();
        var innerDocument = new Document();
        var body = new Body();
        var elements = new OpenXmlElement[] { new Paragraph(), new Table() };
        foreach (var element in elements)
            body.AppendChild(element);
        innerDocument.AppendChild(body);
        mainDocumentPart.Document = innerDocument;

        // Act
        var actual = document.GetDocumentBodyChildElements();

        // Assert
        actual.Should().BeEquivalentTo(body.ChildElements);
    }

    [Fact]
    public void GetDocumentBodyChildElements_NullValue_ThrowEx()
    {
        // Arrange
        using var stream = new MemoryStream();
        var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);

        // Act
        var actual = () => document.GetDocumentBodyChildElements();

        // Assert
        actual.Should().Throw<NullReferenceException>();
    }

    #endregion

    #region GetParagraphSpacingOrDefault

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(100)]
    [InlineData(1000)]
    public void GetParagraphSpacingOrDefault_CorrectValue_CorrectResult(int value)
    {
        // Arrange
        var paragraph = new Paragraph();
        var runProperties = new RunProperties
        {
            Spacing = new Spacing
            {
                Val = value
            }
        };
        paragraph.AppendChild(runProperties);
        var expected = value / DocXFormatConstants.DxaScale;

        // Act
        var actual = paragraph.GetParagraphSpacingOrDefault();

        // Assert
        actual.Should().Be(expected);
    }

    [Fact]
    public void GetParagraphSpacingOrDefault_NullValue_ZeroResult()
    {
        // Arrange
        var paragraph = new Paragraph();
        var runProperties = new RunProperties();
        paragraph.AppendChild(runProperties);

        // Act
        var actual = paragraph.GetParagraphSpacingOrDefault();

        // Assert
        actual.Should().Be(0);
    }

    #endregion

    #region GetParagraphAlignmentOrDefault

    [Theory]
    [InlineData(JustificationValues.Both)]
    [InlineData(JustificationValues.Center)]
    [InlineData(JustificationValues.Right)]
    [InlineData(JustificationValues.Left)]
    public void GetParagraphAlignmentOrDefault_CorrectValue_CorrectResult(JustificationValues value)
    {
        // Arrange
        var paragraph = new Paragraph();
        var paragraphProperties = new ParagraphProperties
        {
            Justification = new Justification
            {
                Val = value
            }
        };
        paragraph.AppendChild(paragraphProperties);

        // Act
        var actual = paragraph.GetParagraphAlignmentOrDefault();

        // Assert
        actual.Should().Be(value);
    }

    [Fact]
    public void GetParagraphAlignmentOrDefault_NullValue_DefaultResult()
    {
        // Arrange
        var paragraph = new Paragraph();
        var paragraphProperties = new ParagraphProperties();
        paragraph.AppendChild(paragraphProperties);

        // Act
        var actual = paragraph.GetParagraphAlignmentOrDefault();

        // Assert
        actual.Should().Be(JustificationValues.Left);
    }

    #endregion
}
