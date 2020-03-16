using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using System.Xml;
using System.Xml.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AbsWordDocument
{
    public class PropriedadesNumeracao
    {
        public NumberFormatValues numberFormatValue;
        public string levelTextValue;
        public LevelJustificationValues levelJustificationValue;
        public int tabStopPositionValue;
        public int indentationStartValue;
        public int indentationHangingValue;
        public FontTypeHintValues fontTypeHintValue;
        public RunFonts runFonts;
        public int fontSize;

        public PropriedadesNumeracao()
        {
            numberFormatValue = NumberFormatValues.Bullet;
            levelTextValue = null;
            levelJustificationValue = LevelJustificationValues.Left;
            tabStopPositionValue = -1;
            indentationStartValue = -1;
            indentationHangingValue = -1;
            fontTypeHintValue = FontTypeHintValues.Default;
            runFonts = null;
            fontSize = -1;
        }
    }

    public static class WordDocUtilities
    {
        static int bookmark = 0;
        static int figureCounter = 0;
        static int tableCounter = 0;

        public static string NewBookmark() { return (bookmark++).ToString(); }
        public static int NewFigureCounter() { return figureCounter++; }
        public static int NewTableCounter() { return tableCounter++; }

        public static string NSID { get { return Guid.NewGuid().ToString("D").Substring(0, 8).ToUpper(); } }

        // Add a paragraph with style
        public static Paragraph CreateParagraphWithStyle(string text, string parastyleid)
        {
            // Add a paragraph with a run and some text.
            Paragraph p =
                new Paragraph(
                    new Run(
                        new Text(text)));

            // If the paragraph has no ParagraphProperties object, create one.
            if (p.Elements<ParagraphProperties>().Count() == 0)
            {
                p.PrependChild<ParagraphProperties>(new ParagraphProperties());
            }

            // Get a reference to the ParagraphProperties object.
            ParagraphProperties pPr = p.ParagraphProperties;

            // If a ParagraphStyleId object doesn't exist, create one.
            if (pPr.ParagraphStyleId == null)
                pPr.ParagraphStyleId = new ParagraphStyleId();

            // Set the style of the paragraph.
            pPr.ParagraphStyleId.Val = parastyleid;

            return p;
        }

        #region Styles
        // Extract the styles or stylesWithEffects part from a 
        // word processing document as an XDocument instance.
        public static XDocument ExtractStylesPart(string fileName, bool getStylesWithEffectsPart = true)
        {
            // Declare a variable to hold the XDocument.
            XDocument styles = null;

            WordprocessingDocument document = null;

            // Open the document for read access and get a reference.
            try
            {
                document = WordprocessingDocument.Open(fileName, false);

                // Get a reference to the main document part.
                var docPart = document.MainDocumentPart;

                // Assign a reference to the appropriate part to the stylesPart variable.
                StylesPart stylesPart = null;
                if (getStylesWithEffectsPart)
                    stylesPart = docPart.StylesWithEffectsPart;
                else
                    stylesPart = docPart.StyleDefinitionsPart;

                // If the part exists, read it into the XDocument.
                if (stylesPart != null)
                {
                    using (var reader = XmlNodeReader.Create(
                      stylesPart.GetStream(FileMode.Open, FileAccess.Read)))
                    {
                        // Create the XDocument.
                        styles = XDocument.Load(reader);
                    }
                }
            }
            finally
            {
                if (document != null)
                    document.Close();
            }

            // Return the XDocument instance.
            return styles;
        }

        // Adds child parts and generates content of the specified part.
        private static void GenerateFooterPartContent(FooterPart footerPart)
        {
            Footer footer1 = new Footer() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid wp14" } };
            footer1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            footer1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            footer1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            footer1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            footer1.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            footer1.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            footer1.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            footer1.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            footer1.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            footer1.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            footer1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footer1.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            footer1.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            footer1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footer1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footer1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footer1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footer1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            footer1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footer1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footer1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footer1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            footer1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            footer1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            footer1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            footer1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            footer1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            footer1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            footer1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Table table3 = new Table();

            TableProperties tableProperties3 = new TableProperties();
            TableWidth tableWidth3 = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };
            TableJustification tableJustification1 = new TableJustification() { Val = TableRowAlignmentValues.Center };

            TableCellMarginDefault tableCellMarginDefault2 = new TableCellMarginDefault();
            TopMargin topMargin4 = new TopMargin() { Width = "144", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin2 = new TableCellLeftMargin() { Width = 115, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin4 = new BottomMargin() { Width = "144", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin2 = new TableCellRightMargin() { Width = 115, Type = TableWidthValues.Dxa };

            tableCellMarginDefault2.Append(topMargin4);
            tableCellMarginDefault2.Append(tableCellLeftMargin2);
            tableCellMarginDefault2.Append(bottomMargin4);
            tableCellMarginDefault2.Append(tableCellRightMargin2);
            TableLook tableLook3 = new TableLook() { Val = "04A0" };

            tableProperties3.Append(tableWidth3);
            tableProperties3.Append(tableJustification1);
            tableProperties3.Append(tableCellMarginDefault2);
            tableProperties3.Append(tableLook3);

            TableGrid tableGrid3 = new TableGrid();
            GridColumn gridColumn3 = new GridColumn() { Width = "4257" };
            GridColumn gridColumn4 = new GridColumn() { Width = "4247" };

            tableGrid3.Append(gridColumn3);
            tableGrid3.Append(gridColumn4);

            TableRow tableRow5 = new TableRow() { RsidTableRowAddition = "00632C30", RsidTableRowProperties = "00632C30" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableJustification tableJustification2 = new TableJustification() { Val = TableRowAlignmentValues.Center };

            tableRowProperties1.Append(tableJustification2);

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "4257", Type = TableWidthUnitValues.Dxa };
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties5.Append(tableCellWidth5);
            tableCellProperties5.Append(shading1);
            tableCellProperties5.Append(tableCellVerticalAlignment1);

            Paragraph paragraph56 = new Paragraph() { RsidParagraphAddition = "00632C30", RsidRunAdditionDefault = "00A72496" };

            ParagraphProperties paragraphProperties27 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId15 = new ParagraphStyleId() { Val = "Rodap" };

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            Caps caps1 = new Caps();
            Color color19 = new Color() { Val = "808080", ThemeColor = ThemeColorValues.Background1, ThemeShade = "80" };
            FontSize fontSize30 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "18" };

            paragraphMarkRunProperties16.Append(caps1);
            paragraphMarkRunProperties16.Append(color19);
            paragraphMarkRunProperties16.Append(fontSize30);
            paragraphMarkRunProperties16.Append(fontSizeComplexScript16);

            paragraphProperties27.Append(paragraphStyleId15);
            paragraphProperties27.Append(paragraphMarkRunProperties16);

            paragraph56.Append(paragraphProperties27);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph56);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "4247", Type = TableWidthUnitValues.Dxa };
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties6.Append(tableCellWidth6);
            tableCellProperties6.Append(shading2);
            tableCellProperties6.Append(tableCellVerticalAlignment2);

            Paragraph paragraph57 = new Paragraph() { RsidParagraphAddition = "00632C30", RsidRunAdditionDefault = "00A72496" };

            ParagraphProperties paragraphProperties28 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId16 = new ParagraphStyleId() { Val = "Rodap" };
            Justification justification8 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            Caps caps2 = new Caps();
            Color color20 = new Color() { Val = "808080", ThemeColor = ThemeColorValues.Background1, ThemeShade = "80" };
            FontSize fontSize31 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "18" };

            paragraphMarkRunProperties17.Append(caps2);
            paragraphMarkRunProperties17.Append(color20);
            paragraphMarkRunProperties17.Append(fontSize31);
            paragraphMarkRunProperties17.Append(fontSizeComplexScript17);

            paragraphProperties28.Append(paragraphStyleId16);
            paragraphProperties28.Append(justification8);
            paragraphProperties28.Append(paragraphMarkRunProperties17);

            paragraph57.Append(paragraphProperties28);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph57);

            tableRow5.Append(tableRowProperties1);
            tableRow5.Append(tableCell5);
            tableRow5.Append(tableCell6);

            table3.Append(tableProperties3);
            table3.Append(tableGrid3);
            table3.Append(tableRow5);

            Paragraph paragraph58 = new Paragraph() { RsidParagraphAddition = "00632C30", RsidRunAdditionDefault = "00A72496" };

            ParagraphProperties paragraphProperties29 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId17 = new ParagraphStyleId() { Val = "Rodap" };

            paragraphProperties29.Append(paragraphStyleId17);

            paragraph58.Append(paragraphProperties29);

            footer1.Append(table3);
            footer1.Append(paragraph58);

            footerPart.Footer = footer1;
        }

        // Generates content of fontTablePart1.
        private static void GenerateFontTablePartContent(FontTablePart fontTablePart)
        {
            Fonts fonts1 = new Fonts() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid" } };
            fonts1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            fonts1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            fonts1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            fonts1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            fonts1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            fonts1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            fonts1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");

            Font font1 = new Font() { Name = "Symbol" };
            Panose1Number panose1Number1 = new Panose1Number() { Val = "05050102010706020507" };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = "02" };
            FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature1 = new FontSignature() { UnicodeSignature0 = "00000000", UnicodeSignature1 = "10000000", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "80000000", CodePageSignature1 = "00000000" };

            font1.Append(panose1Number1);
            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);
            font1.Append(fontSignature1);

            Font font2 = new Font() { Name = "Times New Roman" };
            Panose1Number panose1Number2 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature2 = new FontSignature() { UnicodeSignature0 = "E0002AFF", UnicodeSignature1 = "C0007843", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font2.Append(panose1Number2);
            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);
            font2.Append(fontSignature2);

            Font font3 = new Font() { Name = "Courier New" };
            Panose1Number panose1Number3 = new Panose1Number() { Val = "02070309020205020404" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Modern };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Fixed };
            FontSignature fontSignature3 = new FontSignature() { UnicodeSignature0 = "E0002EFF", UnicodeSignature1 = "C0007843", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font3.Append(panose1Number3);
            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);
            font3.Append(fontSignature3);

            Font font4 = new Font() { Name = "Wingdings" };
            Panose1Number panose1Number4 = new Panose1Number() { Val = "05000000000000000000" };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = "02" };
            FontFamily fontFamily4 = new FontFamily() { Val = FontFamilyValues.Auto };
            Pitch pitch4 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature4 = new FontSignature() { UnicodeSignature0 = "00000000", UnicodeSignature1 = "10000000", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "80000000", CodePageSignature1 = "00000000" };

            font4.Append(panose1Number4);
            font4.Append(fontCharSet4);
            font4.Append(fontFamily4);
            font4.Append(pitch4);
            font4.Append(fontSignature4);

            Font font5 = new Font() { Name = "Calibri Light" };
            Panose1Number panose1Number5 = new Panose1Number() { Val = "020F0302020204030204" };
            FontCharSet fontCharSet5 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily5 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch5 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature5 = new FontSignature() { UnicodeSignature0 = "A0002AEF", UnicodeSignature1 = "4000207B", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font5.Append(panose1Number5);
            font5.Append(fontCharSet5);
            font5.Append(fontFamily5);
            font5.Append(pitch5);
            font5.Append(fontSignature5);

            Font font6 = new Font() { Name = "Calibri" };
            Panose1Number panose1Number6 = new Panose1Number() { Val = "020F0502020204030204" };
            FontCharSet fontCharSet6 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily6 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch6 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature6 = new FontSignature() { UnicodeSignature0 = "E0002AFF", UnicodeSignature1 = "4000ACFF", UnicodeSignature2 = "00000001", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font6.Append(panose1Number6);
            font6.Append(fontCharSet6);
            font6.Append(fontFamily6);
            font6.Append(pitch6);
            font6.Append(fontSignature6);

            Font font7 = new Font() { Name = "Segoe UI" };
            Panose1Number panose1Number7 = new Panose1Number() { Val = "020B0502040204020203" };
            FontCharSet fontCharSet7 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily7 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch7 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature7 = new FontSignature() { UnicodeSignature0 = "E4002EFF", UnicodeSignature1 = "C000E47F", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font7.Append(panose1Number7);
            font7.Append(fontCharSet7);
            font7.Append(fontFamily7);
            font7.Append(pitch7);
            font7.Append(fontSignature7);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);
            fonts1.Append(font6);
            fonts1.Append(font7);

            fontTablePart.Fonts = fonts1;
        }



        // Generates content of styleDefinitionsPart.
        public static Styles GenerateStyleDefinitionsPartContent()
        {
            Styles styles1 = new Styles() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid" } };
            styles1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            styles1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            styles1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            styles1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            styles1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts12 = new RunFonts() { Ascii = "Calibri Light", HighAnsi = "Calibri", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
            Languages languages9 = new Languages() { Val = "pt-BR", EastAsia = "pt-BR", Bidi = "ar-SA" };

            runPropertiesBaseStyle1.Append(runFonts12);
            runPropertiesBaseStyle1.Append(languages9);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);
            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            LatentStyles latentStyles1 = new LatentStyles() { DefaultLockedState = false, DefaultUiPriority = 99, DefaultSemiHidden = false, DefaultUnhideWhenUsed = false, DefaultPrimaryStyle = false, Count = 376 };
            LatentStyleExceptionInfo latentStyleExceptionInfo1 = new LatentStyleExceptionInfo() { Name = "Normal", UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo2 = new LatentStyleExceptionInfo() { Name = "heading 1", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo3 = new LatentStyleExceptionInfo() { Name = "heading 2", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo4 = new LatentStyleExceptionInfo() { Name = "heading 3", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo5 = new LatentStyleExceptionInfo() { Name = "heading 4", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo6 = new LatentStyleExceptionInfo() { Name = "heading 5", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo7 = new LatentStyleExceptionInfo() { Name = "heading 6", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo8 = new LatentStyleExceptionInfo() { Name = "heading 7", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo9 = new LatentStyleExceptionInfo() { Name = "heading 8", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo10 = new LatentStyleExceptionInfo() { Name = "heading 9", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo11 = new LatentStyleExceptionInfo() { Name = "index 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo12 = new LatentStyleExceptionInfo() { Name = "index 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo13 = new LatentStyleExceptionInfo() { Name = "index 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo14 = new LatentStyleExceptionInfo() { Name = "index 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo15 = new LatentStyleExceptionInfo() { Name = "index 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo16 = new LatentStyleExceptionInfo() { Name = "index 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo17 = new LatentStyleExceptionInfo() { Name = "index 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo18 = new LatentStyleExceptionInfo() { Name = "index 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo19 = new LatentStyleExceptionInfo() { Name = "index 9", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo20 = new LatentStyleExceptionInfo() { Name = "toc 1", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo21 = new LatentStyleExceptionInfo() { Name = "toc 2", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo22 = new LatentStyleExceptionInfo() { Name = "toc 3", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo23 = new LatentStyleExceptionInfo() { Name = "toc 4", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo24 = new LatentStyleExceptionInfo() { Name = "toc 5", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo25 = new LatentStyleExceptionInfo() { Name = "toc 6", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo26 = new LatentStyleExceptionInfo() { Name = "toc 7", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo27 = new LatentStyleExceptionInfo() { Name = "toc 8", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo28 = new LatentStyleExceptionInfo() { Name = "toc 9", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo29 = new LatentStyleExceptionInfo() { Name = "Normal Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo30 = new LatentStyleExceptionInfo() { Name = "footnote text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo31 = new LatentStyleExceptionInfo() { Name = "annotation text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo32 = new LatentStyleExceptionInfo() { Name = "header", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo33 = new LatentStyleExceptionInfo() { Name = "footer", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo34 = new LatentStyleExceptionInfo() { Name = "index heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo35 = new LatentStyleExceptionInfo() { Name = "caption", UiPriority = 35, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo36 = new LatentStyleExceptionInfo() { Name = "table of figures", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo37 = new LatentStyleExceptionInfo() { Name = "envelope address", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo38 = new LatentStyleExceptionInfo() { Name = "envelope return", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo39 = new LatentStyleExceptionInfo() { Name = "footnote reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo40 = new LatentStyleExceptionInfo() { Name = "annotation reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo41 = new LatentStyleExceptionInfo() { Name = "line number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo42 = new LatentStyleExceptionInfo() { Name = "page number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo43 = new LatentStyleExceptionInfo() { Name = "endnote reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo44 = new LatentStyleExceptionInfo() { Name = "endnote text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo45 = new LatentStyleExceptionInfo() { Name = "table of authorities", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo46 = new LatentStyleExceptionInfo() { Name = "macro", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo47 = new LatentStyleExceptionInfo() { Name = "toa heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo48 = new LatentStyleExceptionInfo() { Name = "List", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo49 = new LatentStyleExceptionInfo() { Name = "List Bullet", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo50 = new LatentStyleExceptionInfo() { Name = "List Number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo51 = new LatentStyleExceptionInfo() { Name = "List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo52 = new LatentStyleExceptionInfo() { Name = "List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo53 = new LatentStyleExceptionInfo() { Name = "List 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo54 = new LatentStyleExceptionInfo() { Name = "List 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo55 = new LatentStyleExceptionInfo() { Name = "List Bullet 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo56 = new LatentStyleExceptionInfo() { Name = "List Bullet 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo57 = new LatentStyleExceptionInfo() { Name = "List Bullet 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo58 = new LatentStyleExceptionInfo() { Name = "List Bullet 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo59 = new LatentStyleExceptionInfo() { Name = "List Number 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo60 = new LatentStyleExceptionInfo() { Name = "List Number 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo61 = new LatentStyleExceptionInfo() { Name = "List Number 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo62 = new LatentStyleExceptionInfo() { Name = "List Number 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo63 = new LatentStyleExceptionInfo() { Name = "Title", UiPriority = 10, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo64 = new LatentStyleExceptionInfo() { Name = "Closing", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo65 = new LatentStyleExceptionInfo() { Name = "Signature", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo66 = new LatentStyleExceptionInfo() { Name = "Default Paragraph Font", UiPriority = 1, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo67 = new LatentStyleExceptionInfo() { Name = "Body Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo68 = new LatentStyleExceptionInfo() { Name = "Body Text Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo69 = new LatentStyleExceptionInfo() { Name = "List Continue", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo70 = new LatentStyleExceptionInfo() { Name = "List Continue 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo71 = new LatentStyleExceptionInfo() { Name = "List Continue 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo72 = new LatentStyleExceptionInfo() { Name = "List Continue 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo73 = new LatentStyleExceptionInfo() { Name = "List Continue 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo74 = new LatentStyleExceptionInfo() { Name = "Message Header", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo75 = new LatentStyleExceptionInfo() { Name = "Subtitle", UiPriority = 11, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo76 = new LatentStyleExceptionInfo() { Name = "Salutation", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo77 = new LatentStyleExceptionInfo() { Name = "Date", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo78 = new LatentStyleExceptionInfo() { Name = "Body Text First Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo79 = new LatentStyleExceptionInfo() { Name = "Body Text First Indent 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo80 = new LatentStyleExceptionInfo() { Name = "Note Heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo81 = new LatentStyleExceptionInfo() { Name = "Body Text 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo82 = new LatentStyleExceptionInfo() { Name = "Body Text 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo83 = new LatentStyleExceptionInfo() { Name = "Body Text Indent 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo84 = new LatentStyleExceptionInfo() { Name = "Body Text Indent 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo85 = new LatentStyleExceptionInfo() { Name = "Block Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo86 = new LatentStyleExceptionInfo() { Name = "Hyperlink", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo87 = new LatentStyleExceptionInfo() { Name = "FollowedHyperlink", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo88 = new LatentStyleExceptionInfo() { Name = "Strong", UiPriority = 22, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo89 = new LatentStyleExceptionInfo() { Name = "Emphasis", UiPriority = 20, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo90 = new LatentStyleExceptionInfo() { Name = "Document Map", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo91 = new LatentStyleExceptionInfo() { Name = "Plain Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo92 = new LatentStyleExceptionInfo() { Name = "E-mail Signature", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo93 = new LatentStyleExceptionInfo() { Name = "HTML Top of Form", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo94 = new LatentStyleExceptionInfo() { Name = "HTML Bottom of Form", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo95 = new LatentStyleExceptionInfo() { Name = "Normal (Web)", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo96 = new LatentStyleExceptionInfo() { Name = "HTML Acronym", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo97 = new LatentStyleExceptionInfo() { Name = "HTML Address", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo98 = new LatentStyleExceptionInfo() { Name = "HTML Cite", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo99 = new LatentStyleExceptionInfo() { Name = "HTML Code", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo100 = new LatentStyleExceptionInfo() { Name = "HTML Definition", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo101 = new LatentStyleExceptionInfo() { Name = "HTML Keyboard", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo102 = new LatentStyleExceptionInfo() { Name = "HTML Preformatted", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo103 = new LatentStyleExceptionInfo() { Name = "HTML Sample", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo104 = new LatentStyleExceptionInfo() { Name = "HTML Typewriter", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo105 = new LatentStyleExceptionInfo() { Name = "HTML Variable", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo106 = new LatentStyleExceptionInfo() { Name = "annotation subject", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo107 = new LatentStyleExceptionInfo() { Name = "No List", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo108 = new LatentStyleExceptionInfo() { Name = "Outline List 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo109 = new LatentStyleExceptionInfo() { Name = "Outline List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo110 = new LatentStyleExceptionInfo() { Name = "Outline List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo111 = new LatentStyleExceptionInfo() { Name = "Table Simple 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo112 = new LatentStyleExceptionInfo() { Name = "Table Simple 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo113 = new LatentStyleExceptionInfo() { Name = "Table Simple 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo114 = new LatentStyleExceptionInfo() { Name = "Table Classic 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo115 = new LatentStyleExceptionInfo() { Name = "Table Classic 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo116 = new LatentStyleExceptionInfo() { Name = "Table Classic 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo117 = new LatentStyleExceptionInfo() { Name = "Table Classic 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo118 = new LatentStyleExceptionInfo() { Name = "Table Colorful 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo119 = new LatentStyleExceptionInfo() { Name = "Table Colorful 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo120 = new LatentStyleExceptionInfo() { Name = "Table Colorful 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo121 = new LatentStyleExceptionInfo() { Name = "Table Columns 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo122 = new LatentStyleExceptionInfo() { Name = "Table Columns 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo123 = new LatentStyleExceptionInfo() { Name = "Table Columns 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo124 = new LatentStyleExceptionInfo() { Name = "Table Columns 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo125 = new LatentStyleExceptionInfo() { Name = "Table Columns 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo126 = new LatentStyleExceptionInfo() { Name = "Table Grid 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo127 = new LatentStyleExceptionInfo() { Name = "Table Grid 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo128 = new LatentStyleExceptionInfo() { Name = "Table Grid 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo129 = new LatentStyleExceptionInfo() { Name = "Table Grid 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo130 = new LatentStyleExceptionInfo() { Name = "Table Grid 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo131 = new LatentStyleExceptionInfo() { Name = "Table Grid 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo132 = new LatentStyleExceptionInfo() { Name = "Table Grid 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo133 = new LatentStyleExceptionInfo() { Name = "Table Grid 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo134 = new LatentStyleExceptionInfo() { Name = "Table List 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo135 = new LatentStyleExceptionInfo() { Name = "Table List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo136 = new LatentStyleExceptionInfo() { Name = "Table List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo137 = new LatentStyleExceptionInfo() { Name = "Table List 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo138 = new LatentStyleExceptionInfo() { Name = "Table List 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo139 = new LatentStyleExceptionInfo() { Name = "Table List 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo140 = new LatentStyleExceptionInfo() { Name = "Table List 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo141 = new LatentStyleExceptionInfo() { Name = "Table List 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo142 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo143 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo144 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo145 = new LatentStyleExceptionInfo() { Name = "Table Contemporary", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo146 = new LatentStyleExceptionInfo() { Name = "Table Elegant", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo147 = new LatentStyleExceptionInfo() { Name = "Table Professional", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo148 = new LatentStyleExceptionInfo() { Name = "Table Subtle 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo149 = new LatentStyleExceptionInfo() { Name = "Table Subtle 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo150 = new LatentStyleExceptionInfo() { Name = "Table Web 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo151 = new LatentStyleExceptionInfo() { Name = "Table Web 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo152 = new LatentStyleExceptionInfo() { Name = "Balloon Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo153 = new LatentStyleExceptionInfo() { Name = "Table Grid", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo154 = new LatentStyleExceptionInfo() { Name = "Placeholder Text", SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo155 = new LatentStyleExceptionInfo() { Name = "No Spacing", UiPriority = 1, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo156 = new LatentStyleExceptionInfo() { Name = "Light Shading", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo157 = new LatentStyleExceptionInfo() { Name = "Light List", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo158 = new LatentStyleExceptionInfo() { Name = "Light Grid", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo159 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo160 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo161 = new LatentStyleExceptionInfo() { Name = "Medium List 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo162 = new LatentStyleExceptionInfo() { Name = "Medium List 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo163 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo164 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo165 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo166 = new LatentStyleExceptionInfo() { Name = "Dark List", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo167 = new LatentStyleExceptionInfo() { Name = "Colorful Shading", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo168 = new LatentStyleExceptionInfo() { Name = "Colorful List", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo169 = new LatentStyleExceptionInfo() { Name = "Colorful Grid", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo170 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 1", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo171 = new LatentStyleExceptionInfo() { Name = "Light List Accent 1", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo172 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 1", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo173 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo174 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 1", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo175 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo176 = new LatentStyleExceptionInfo() { Name = "Revision", SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo177 = new LatentStyleExceptionInfo() { Name = "List Paragraph", UiPriority = 34, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo178 = new LatentStyleExceptionInfo() { Name = "Quote", UiPriority = 29, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo179 = new LatentStyleExceptionInfo() { Name = "Intense Quote", UiPriority = 30, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo180 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 1", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo181 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo182 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 1", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo183 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 1", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo184 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 1", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo185 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 1", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo186 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 1", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo187 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 1", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo188 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 2", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo189 = new LatentStyleExceptionInfo() { Name = "Light List Accent 2", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo190 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 2", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo191 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 2", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo192 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo193 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 2", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo194 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo195 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 2", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo196 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo197 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 2", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo198 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 2", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo199 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 2", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo200 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 2", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo201 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 2", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo202 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 3", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo203 = new LatentStyleExceptionInfo() { Name = "Light List Accent 3", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo204 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 3", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo205 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 3", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo206 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 3", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo207 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 3", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo208 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 3", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo209 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 3", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo210 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 3", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo211 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo212 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 3", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo213 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 3", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo214 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 3", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo215 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 3", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo216 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 4", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo217 = new LatentStyleExceptionInfo() { Name = "Light List Accent 4", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo218 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 4", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo219 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 4", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo220 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 4", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo221 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 4", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo222 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 4", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo223 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 4", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo224 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 4", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo225 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 4", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo226 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 4", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo227 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 4", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo228 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 4", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo229 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 4", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo230 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 5", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo231 = new LatentStyleExceptionInfo() { Name = "Light List Accent 5", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo232 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 5", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo233 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 5", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo234 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 5", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo235 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 5", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo236 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 5", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo237 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 5", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo238 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 5", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo239 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 5", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo240 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 5", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo241 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 5", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo242 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 5", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo243 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 5", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo244 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 6", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo245 = new LatentStyleExceptionInfo() { Name = "Light List Accent 6", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo246 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 6", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo247 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 6", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo248 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 6", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo249 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 6", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo250 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 6", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo251 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 6", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo252 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 6", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo253 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 6", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo254 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 6", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo255 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 6", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo256 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 6", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo257 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 6", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo258 = new LatentStyleExceptionInfo() { Name = "Subtle Emphasis", UiPriority = 19, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo259 = new LatentStyleExceptionInfo() { Name = "Intense Emphasis", UiPriority = 21, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo260 = new LatentStyleExceptionInfo() { Name = "Subtle Reference", UiPriority = 31, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo261 = new LatentStyleExceptionInfo() { Name = "Intense Reference", UiPriority = 32, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo262 = new LatentStyleExceptionInfo() { Name = "Book Title", UiPriority = 33, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo263 = new LatentStyleExceptionInfo() { Name = "Bibliography", UiPriority = 37, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo264 = new LatentStyleExceptionInfo() { Name = "TOC Heading", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo265 = new LatentStyleExceptionInfo() { Name = "Plain Table 1", UiPriority = 41 };
            LatentStyleExceptionInfo latentStyleExceptionInfo266 = new LatentStyleExceptionInfo() { Name = "Plain Table 2", UiPriority = 42 };
            LatentStyleExceptionInfo latentStyleExceptionInfo267 = new LatentStyleExceptionInfo() { Name = "Plain Table 3", UiPriority = 43 };
            LatentStyleExceptionInfo latentStyleExceptionInfo268 = new LatentStyleExceptionInfo() { Name = "Plain Table 4", UiPriority = 44 };
            LatentStyleExceptionInfo latentStyleExceptionInfo269 = new LatentStyleExceptionInfo() { Name = "Plain Table 5", UiPriority = 45 };
            LatentStyleExceptionInfo latentStyleExceptionInfo270 = new LatentStyleExceptionInfo() { Name = "Grid Table Light", UiPriority = 40 };
            LatentStyleExceptionInfo latentStyleExceptionInfo271 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo272 = new LatentStyleExceptionInfo() { Name = "Grid Table 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo273 = new LatentStyleExceptionInfo() { Name = "Grid Table 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo274 = new LatentStyleExceptionInfo() { Name = "Grid Table 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo275 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo276 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo277 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo278 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 1", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo279 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 1", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo280 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 1", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo281 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 1", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo282 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 1", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo283 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 1", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo284 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 1", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo285 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 2", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo286 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo287 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 2", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo288 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 2", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo289 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 2", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo290 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 2", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo291 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 2", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo292 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 3", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo293 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 3", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo294 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo295 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 3", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo296 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 3", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo297 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 3", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo298 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 3", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo299 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 4", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo300 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 4", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo301 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 4", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo302 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo303 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 4", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo304 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 4", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo305 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 4", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo306 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 5", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo307 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 5", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo308 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 5", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo309 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 5", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo310 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 5", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo311 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 5", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo312 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 5", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo313 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 6", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo314 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 6", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo315 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 6", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo316 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 6", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo317 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 6", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo318 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 6", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo319 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 6", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo320 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo321 = new LatentStyleExceptionInfo() { Name = "List Table 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo322 = new LatentStyleExceptionInfo() { Name = "List Table 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo323 = new LatentStyleExceptionInfo() { Name = "List Table 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo324 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo325 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo326 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo327 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 1", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo328 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 1", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo329 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 1", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo330 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 1", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo331 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 1", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo332 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 1", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo333 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 1", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo334 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 2", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo335 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo336 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 2", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo337 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 2", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo338 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 2", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo339 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 2", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo340 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 2", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo341 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 3", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo342 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 3", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo343 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo344 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 3", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo345 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 3", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo346 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 3", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo347 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 3", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo348 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 4", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo349 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 4", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo350 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 4", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo351 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo352 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 4", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo353 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 4", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo354 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 4", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo355 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 5", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo356 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 5", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo357 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 5", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo358 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 5", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo359 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 5", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo360 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 5", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo361 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 5", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo362 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 6", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo363 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 6", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo364 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 6", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo365 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 6", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo366 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 6", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo367 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 6", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo368 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 6", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo369 = new LatentStyleExceptionInfo() { Name = "Mention", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo370 = new LatentStyleExceptionInfo() { Name = "Smart Hyperlink", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo371 = new LatentStyleExceptionInfo() { Name = "Hashtag", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo372 = new LatentStyleExceptionInfo() { Name = "Unresolved Mention", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo373 = new LatentStyleExceptionInfo() { Name = "Smart Link", SemiHidden = true, UnhideWhenUsed = true };

            latentStyles1.Append(latentStyleExceptionInfo1);
            latentStyles1.Append(latentStyleExceptionInfo2);
            latentStyles1.Append(latentStyleExceptionInfo3);
            latentStyles1.Append(latentStyleExceptionInfo4);
            latentStyles1.Append(latentStyleExceptionInfo5);
            latentStyles1.Append(latentStyleExceptionInfo6);
            latentStyles1.Append(latentStyleExceptionInfo7);
            latentStyles1.Append(latentStyleExceptionInfo8);
            latentStyles1.Append(latentStyleExceptionInfo9);
            latentStyles1.Append(latentStyleExceptionInfo10);
            latentStyles1.Append(latentStyleExceptionInfo11);
            latentStyles1.Append(latentStyleExceptionInfo12);
            latentStyles1.Append(latentStyleExceptionInfo13);
            latentStyles1.Append(latentStyleExceptionInfo14);
            latentStyles1.Append(latentStyleExceptionInfo15);
            latentStyles1.Append(latentStyleExceptionInfo16);
            latentStyles1.Append(latentStyleExceptionInfo17);
            latentStyles1.Append(latentStyleExceptionInfo18);
            latentStyles1.Append(latentStyleExceptionInfo19);
            latentStyles1.Append(latentStyleExceptionInfo20);
            latentStyles1.Append(latentStyleExceptionInfo21);
            latentStyles1.Append(latentStyleExceptionInfo22);
            latentStyles1.Append(latentStyleExceptionInfo23);
            latentStyles1.Append(latentStyleExceptionInfo24);
            latentStyles1.Append(latentStyleExceptionInfo25);
            latentStyles1.Append(latentStyleExceptionInfo26);
            latentStyles1.Append(latentStyleExceptionInfo27);
            latentStyles1.Append(latentStyleExceptionInfo28);
            latentStyles1.Append(latentStyleExceptionInfo29);
            latentStyles1.Append(latentStyleExceptionInfo30);
            latentStyles1.Append(latentStyleExceptionInfo31);
            latentStyles1.Append(latentStyleExceptionInfo32);
            latentStyles1.Append(latentStyleExceptionInfo33);
            latentStyles1.Append(latentStyleExceptionInfo34);
            latentStyles1.Append(latentStyleExceptionInfo35);
            latentStyles1.Append(latentStyleExceptionInfo36);
            latentStyles1.Append(latentStyleExceptionInfo37);
            latentStyles1.Append(latentStyleExceptionInfo38);
            latentStyles1.Append(latentStyleExceptionInfo39);
            latentStyles1.Append(latentStyleExceptionInfo40);
            latentStyles1.Append(latentStyleExceptionInfo41);
            latentStyles1.Append(latentStyleExceptionInfo42);
            latentStyles1.Append(latentStyleExceptionInfo43);
            latentStyles1.Append(latentStyleExceptionInfo44);
            latentStyles1.Append(latentStyleExceptionInfo45);
            latentStyles1.Append(latentStyleExceptionInfo46);
            latentStyles1.Append(latentStyleExceptionInfo47);
            latentStyles1.Append(latentStyleExceptionInfo48);
            latentStyles1.Append(latentStyleExceptionInfo49);
            latentStyles1.Append(latentStyleExceptionInfo50);
            latentStyles1.Append(latentStyleExceptionInfo51);
            latentStyles1.Append(latentStyleExceptionInfo52);
            latentStyles1.Append(latentStyleExceptionInfo53);
            latentStyles1.Append(latentStyleExceptionInfo54);
            latentStyles1.Append(latentStyleExceptionInfo55);
            latentStyles1.Append(latentStyleExceptionInfo56);
            latentStyles1.Append(latentStyleExceptionInfo57);
            latentStyles1.Append(latentStyleExceptionInfo58);
            latentStyles1.Append(latentStyleExceptionInfo59);
            latentStyles1.Append(latentStyleExceptionInfo60);
            latentStyles1.Append(latentStyleExceptionInfo61);
            latentStyles1.Append(latentStyleExceptionInfo62);
            latentStyles1.Append(latentStyleExceptionInfo63);
            latentStyles1.Append(latentStyleExceptionInfo64);
            latentStyles1.Append(latentStyleExceptionInfo65);
            latentStyles1.Append(latentStyleExceptionInfo66);
            latentStyles1.Append(latentStyleExceptionInfo67);
            latentStyles1.Append(latentStyleExceptionInfo68);
            latentStyles1.Append(latentStyleExceptionInfo69);
            latentStyles1.Append(latentStyleExceptionInfo70);
            latentStyles1.Append(latentStyleExceptionInfo71);
            latentStyles1.Append(latentStyleExceptionInfo72);
            latentStyles1.Append(latentStyleExceptionInfo73);
            latentStyles1.Append(latentStyleExceptionInfo74);
            latentStyles1.Append(latentStyleExceptionInfo75);
            latentStyles1.Append(latentStyleExceptionInfo76);
            latentStyles1.Append(latentStyleExceptionInfo77);
            latentStyles1.Append(latentStyleExceptionInfo78);
            latentStyles1.Append(latentStyleExceptionInfo79);
            latentStyles1.Append(latentStyleExceptionInfo80);
            latentStyles1.Append(latentStyleExceptionInfo81);
            latentStyles1.Append(latentStyleExceptionInfo82);
            latentStyles1.Append(latentStyleExceptionInfo83);
            latentStyles1.Append(latentStyleExceptionInfo84);
            latentStyles1.Append(latentStyleExceptionInfo85);
            latentStyles1.Append(latentStyleExceptionInfo86);
            latentStyles1.Append(latentStyleExceptionInfo87);
            latentStyles1.Append(latentStyleExceptionInfo88);
            latentStyles1.Append(latentStyleExceptionInfo89);
            latentStyles1.Append(latentStyleExceptionInfo90);
            latentStyles1.Append(latentStyleExceptionInfo91);
            latentStyles1.Append(latentStyleExceptionInfo92);
            latentStyles1.Append(latentStyleExceptionInfo93);
            latentStyles1.Append(latentStyleExceptionInfo94);
            latentStyles1.Append(latentStyleExceptionInfo95);
            latentStyles1.Append(latentStyleExceptionInfo96);
            latentStyles1.Append(latentStyleExceptionInfo97);
            latentStyles1.Append(latentStyleExceptionInfo98);
            latentStyles1.Append(latentStyleExceptionInfo99);
            latentStyles1.Append(latentStyleExceptionInfo100);
            latentStyles1.Append(latentStyleExceptionInfo101);
            latentStyles1.Append(latentStyleExceptionInfo102);
            latentStyles1.Append(latentStyleExceptionInfo103);
            latentStyles1.Append(latentStyleExceptionInfo104);
            latentStyles1.Append(latentStyleExceptionInfo105);
            latentStyles1.Append(latentStyleExceptionInfo106);
            latentStyles1.Append(latentStyleExceptionInfo107);
            latentStyles1.Append(latentStyleExceptionInfo108);
            latentStyles1.Append(latentStyleExceptionInfo109);
            latentStyles1.Append(latentStyleExceptionInfo110);
            latentStyles1.Append(latentStyleExceptionInfo111);
            latentStyles1.Append(latentStyleExceptionInfo112);
            latentStyles1.Append(latentStyleExceptionInfo113);
            latentStyles1.Append(latentStyleExceptionInfo114);
            latentStyles1.Append(latentStyleExceptionInfo115);
            latentStyles1.Append(latentStyleExceptionInfo116);
            latentStyles1.Append(latentStyleExceptionInfo117);
            latentStyles1.Append(latentStyleExceptionInfo118);
            latentStyles1.Append(latentStyleExceptionInfo119);
            latentStyles1.Append(latentStyleExceptionInfo120);
            latentStyles1.Append(latentStyleExceptionInfo121);
            latentStyles1.Append(latentStyleExceptionInfo122);
            latentStyles1.Append(latentStyleExceptionInfo123);
            latentStyles1.Append(latentStyleExceptionInfo124);
            latentStyles1.Append(latentStyleExceptionInfo125);
            latentStyles1.Append(latentStyleExceptionInfo126);
            latentStyles1.Append(latentStyleExceptionInfo127);
            latentStyles1.Append(latentStyleExceptionInfo128);
            latentStyles1.Append(latentStyleExceptionInfo129);
            latentStyles1.Append(latentStyleExceptionInfo130);
            latentStyles1.Append(latentStyleExceptionInfo131);
            latentStyles1.Append(latentStyleExceptionInfo132);
            latentStyles1.Append(latentStyleExceptionInfo133);
            latentStyles1.Append(latentStyleExceptionInfo134);
            latentStyles1.Append(latentStyleExceptionInfo135);
            latentStyles1.Append(latentStyleExceptionInfo136);
            latentStyles1.Append(latentStyleExceptionInfo137);
            latentStyles1.Append(latentStyleExceptionInfo138);
            latentStyles1.Append(latentStyleExceptionInfo139);
            latentStyles1.Append(latentStyleExceptionInfo140);
            latentStyles1.Append(latentStyleExceptionInfo141);
            latentStyles1.Append(latentStyleExceptionInfo142);
            latentStyles1.Append(latentStyleExceptionInfo143);
            latentStyles1.Append(latentStyleExceptionInfo144);
            latentStyles1.Append(latentStyleExceptionInfo145);
            latentStyles1.Append(latentStyleExceptionInfo146);
            latentStyles1.Append(latentStyleExceptionInfo147);
            latentStyles1.Append(latentStyleExceptionInfo148);
            latentStyles1.Append(latentStyleExceptionInfo149);
            latentStyles1.Append(latentStyleExceptionInfo150);
            latentStyles1.Append(latentStyleExceptionInfo151);
            latentStyles1.Append(latentStyleExceptionInfo152);
            latentStyles1.Append(latentStyleExceptionInfo153);
            latentStyles1.Append(latentStyleExceptionInfo154);
            latentStyles1.Append(latentStyleExceptionInfo155);
            latentStyles1.Append(latentStyleExceptionInfo156);
            latentStyles1.Append(latentStyleExceptionInfo157);
            latentStyles1.Append(latentStyleExceptionInfo158);
            latentStyles1.Append(latentStyleExceptionInfo159);
            latentStyles1.Append(latentStyleExceptionInfo160);
            latentStyles1.Append(latentStyleExceptionInfo161);
            latentStyles1.Append(latentStyleExceptionInfo162);
            latentStyles1.Append(latentStyleExceptionInfo163);
            latentStyles1.Append(latentStyleExceptionInfo164);
            latentStyles1.Append(latentStyleExceptionInfo165);
            latentStyles1.Append(latentStyleExceptionInfo166);
            latentStyles1.Append(latentStyleExceptionInfo167);
            latentStyles1.Append(latentStyleExceptionInfo168);
            latentStyles1.Append(latentStyleExceptionInfo169);
            latentStyles1.Append(latentStyleExceptionInfo170);
            latentStyles1.Append(latentStyleExceptionInfo171);
            latentStyles1.Append(latentStyleExceptionInfo172);
            latentStyles1.Append(latentStyleExceptionInfo173);
            latentStyles1.Append(latentStyleExceptionInfo174);
            latentStyles1.Append(latentStyleExceptionInfo175);
            latentStyles1.Append(latentStyleExceptionInfo176);
            latentStyles1.Append(latentStyleExceptionInfo177);
            latentStyles1.Append(latentStyleExceptionInfo178);
            latentStyles1.Append(latentStyleExceptionInfo179);
            latentStyles1.Append(latentStyleExceptionInfo180);
            latentStyles1.Append(latentStyleExceptionInfo181);
            latentStyles1.Append(latentStyleExceptionInfo182);
            latentStyles1.Append(latentStyleExceptionInfo183);
            latentStyles1.Append(latentStyleExceptionInfo184);
            latentStyles1.Append(latentStyleExceptionInfo185);
            latentStyles1.Append(latentStyleExceptionInfo186);
            latentStyles1.Append(latentStyleExceptionInfo187);
            latentStyles1.Append(latentStyleExceptionInfo188);
            latentStyles1.Append(latentStyleExceptionInfo189);
            latentStyles1.Append(latentStyleExceptionInfo190);
            latentStyles1.Append(latentStyleExceptionInfo191);
            latentStyles1.Append(latentStyleExceptionInfo192);
            latentStyles1.Append(latentStyleExceptionInfo193);
            latentStyles1.Append(latentStyleExceptionInfo194);
            latentStyles1.Append(latentStyleExceptionInfo195);
            latentStyles1.Append(latentStyleExceptionInfo196);
            latentStyles1.Append(latentStyleExceptionInfo197);
            latentStyles1.Append(latentStyleExceptionInfo198);
            latentStyles1.Append(latentStyleExceptionInfo199);
            latentStyles1.Append(latentStyleExceptionInfo200);
            latentStyles1.Append(latentStyleExceptionInfo201);
            latentStyles1.Append(latentStyleExceptionInfo202);
            latentStyles1.Append(latentStyleExceptionInfo203);
            latentStyles1.Append(latentStyleExceptionInfo204);
            latentStyles1.Append(latentStyleExceptionInfo205);
            latentStyles1.Append(latentStyleExceptionInfo206);
            latentStyles1.Append(latentStyleExceptionInfo207);
            latentStyles1.Append(latentStyleExceptionInfo208);
            latentStyles1.Append(latentStyleExceptionInfo209);
            latentStyles1.Append(latentStyleExceptionInfo210);
            latentStyles1.Append(latentStyleExceptionInfo211);
            latentStyles1.Append(latentStyleExceptionInfo212);
            latentStyles1.Append(latentStyleExceptionInfo213);
            latentStyles1.Append(latentStyleExceptionInfo214);
            latentStyles1.Append(latentStyleExceptionInfo215);
            latentStyles1.Append(latentStyleExceptionInfo216);
            latentStyles1.Append(latentStyleExceptionInfo217);
            latentStyles1.Append(latentStyleExceptionInfo218);
            latentStyles1.Append(latentStyleExceptionInfo219);
            latentStyles1.Append(latentStyleExceptionInfo220);
            latentStyles1.Append(latentStyleExceptionInfo221);
            latentStyles1.Append(latentStyleExceptionInfo222);
            latentStyles1.Append(latentStyleExceptionInfo223);
            latentStyles1.Append(latentStyleExceptionInfo224);
            latentStyles1.Append(latentStyleExceptionInfo225);
            latentStyles1.Append(latentStyleExceptionInfo226);
            latentStyles1.Append(latentStyleExceptionInfo227);
            latentStyles1.Append(latentStyleExceptionInfo228);
            latentStyles1.Append(latentStyleExceptionInfo229);
            latentStyles1.Append(latentStyleExceptionInfo230);
            latentStyles1.Append(latentStyleExceptionInfo231);
            latentStyles1.Append(latentStyleExceptionInfo232);
            latentStyles1.Append(latentStyleExceptionInfo233);
            latentStyles1.Append(latentStyleExceptionInfo234);
            latentStyles1.Append(latentStyleExceptionInfo235);
            latentStyles1.Append(latentStyleExceptionInfo236);
            latentStyles1.Append(latentStyleExceptionInfo237);
            latentStyles1.Append(latentStyleExceptionInfo238);
            latentStyles1.Append(latentStyleExceptionInfo239);
            latentStyles1.Append(latentStyleExceptionInfo240);
            latentStyles1.Append(latentStyleExceptionInfo241);
            latentStyles1.Append(latentStyleExceptionInfo242);
            latentStyles1.Append(latentStyleExceptionInfo243);
            latentStyles1.Append(latentStyleExceptionInfo244);
            latentStyles1.Append(latentStyleExceptionInfo245);
            latentStyles1.Append(latentStyleExceptionInfo246);
            latentStyles1.Append(latentStyleExceptionInfo247);
            latentStyles1.Append(latentStyleExceptionInfo248);
            latentStyles1.Append(latentStyleExceptionInfo249);
            latentStyles1.Append(latentStyleExceptionInfo250);
            latentStyles1.Append(latentStyleExceptionInfo251);
            latentStyles1.Append(latentStyleExceptionInfo252);
            latentStyles1.Append(latentStyleExceptionInfo253);
            latentStyles1.Append(latentStyleExceptionInfo254);
            latentStyles1.Append(latentStyleExceptionInfo255);
            latentStyles1.Append(latentStyleExceptionInfo256);
            latentStyles1.Append(latentStyleExceptionInfo257);
            latentStyles1.Append(latentStyleExceptionInfo258);
            latentStyles1.Append(latentStyleExceptionInfo259);
            latentStyles1.Append(latentStyleExceptionInfo260);
            latentStyles1.Append(latentStyleExceptionInfo261);
            latentStyles1.Append(latentStyleExceptionInfo262);
            latentStyles1.Append(latentStyleExceptionInfo263);
            latentStyles1.Append(latentStyleExceptionInfo264);
            latentStyles1.Append(latentStyleExceptionInfo265);
            latentStyles1.Append(latentStyleExceptionInfo266);
            latentStyles1.Append(latentStyleExceptionInfo267);
            latentStyles1.Append(latentStyleExceptionInfo268);
            latentStyles1.Append(latentStyleExceptionInfo269);
            latentStyles1.Append(latentStyleExceptionInfo270);
            latentStyles1.Append(latentStyleExceptionInfo271);
            latentStyles1.Append(latentStyleExceptionInfo272);
            latentStyles1.Append(latentStyleExceptionInfo273);
            latentStyles1.Append(latentStyleExceptionInfo274);
            latentStyles1.Append(latentStyleExceptionInfo275);
            latentStyles1.Append(latentStyleExceptionInfo276);
            latentStyles1.Append(latentStyleExceptionInfo277);
            latentStyles1.Append(latentStyleExceptionInfo278);
            latentStyles1.Append(latentStyleExceptionInfo279);
            latentStyles1.Append(latentStyleExceptionInfo280);
            latentStyles1.Append(latentStyleExceptionInfo281);
            latentStyles1.Append(latentStyleExceptionInfo282);
            latentStyles1.Append(latentStyleExceptionInfo283);
            latentStyles1.Append(latentStyleExceptionInfo284);
            latentStyles1.Append(latentStyleExceptionInfo285);
            latentStyles1.Append(latentStyleExceptionInfo286);
            latentStyles1.Append(latentStyleExceptionInfo287);
            latentStyles1.Append(latentStyleExceptionInfo288);
            latentStyles1.Append(latentStyleExceptionInfo289);
            latentStyles1.Append(latentStyleExceptionInfo290);
            latentStyles1.Append(latentStyleExceptionInfo291);
            latentStyles1.Append(latentStyleExceptionInfo292);
            latentStyles1.Append(latentStyleExceptionInfo293);
            latentStyles1.Append(latentStyleExceptionInfo294);
            latentStyles1.Append(latentStyleExceptionInfo295);
            latentStyles1.Append(latentStyleExceptionInfo296);
            latentStyles1.Append(latentStyleExceptionInfo297);
            latentStyles1.Append(latentStyleExceptionInfo298);
            latentStyles1.Append(latentStyleExceptionInfo299);
            latentStyles1.Append(latentStyleExceptionInfo300);
            latentStyles1.Append(latentStyleExceptionInfo301);
            latentStyles1.Append(latentStyleExceptionInfo302);
            latentStyles1.Append(latentStyleExceptionInfo303);
            latentStyles1.Append(latentStyleExceptionInfo304);
            latentStyles1.Append(latentStyleExceptionInfo305);
            latentStyles1.Append(latentStyleExceptionInfo306);
            latentStyles1.Append(latentStyleExceptionInfo307);
            latentStyles1.Append(latentStyleExceptionInfo308);
            latentStyles1.Append(latentStyleExceptionInfo309);
            latentStyles1.Append(latentStyleExceptionInfo310);
            latentStyles1.Append(latentStyleExceptionInfo311);
            latentStyles1.Append(latentStyleExceptionInfo312);
            latentStyles1.Append(latentStyleExceptionInfo313);
            latentStyles1.Append(latentStyleExceptionInfo314);
            latentStyles1.Append(latentStyleExceptionInfo315);
            latentStyles1.Append(latentStyleExceptionInfo316);
            latentStyles1.Append(latentStyleExceptionInfo317);
            latentStyles1.Append(latentStyleExceptionInfo318);
            latentStyles1.Append(latentStyleExceptionInfo319);
            latentStyles1.Append(latentStyleExceptionInfo320);
            latentStyles1.Append(latentStyleExceptionInfo321);
            latentStyles1.Append(latentStyleExceptionInfo322);
            latentStyles1.Append(latentStyleExceptionInfo323);
            latentStyles1.Append(latentStyleExceptionInfo324);
            latentStyles1.Append(latentStyleExceptionInfo325);
            latentStyles1.Append(latentStyleExceptionInfo326);
            latentStyles1.Append(latentStyleExceptionInfo327);
            latentStyles1.Append(latentStyleExceptionInfo328);
            latentStyles1.Append(latentStyleExceptionInfo329);
            latentStyles1.Append(latentStyleExceptionInfo330);
            latentStyles1.Append(latentStyleExceptionInfo331);
            latentStyles1.Append(latentStyleExceptionInfo332);
            latentStyles1.Append(latentStyleExceptionInfo333);
            latentStyles1.Append(latentStyleExceptionInfo334);
            latentStyles1.Append(latentStyleExceptionInfo335);
            latentStyles1.Append(latentStyleExceptionInfo336);
            latentStyles1.Append(latentStyleExceptionInfo337);
            latentStyles1.Append(latentStyleExceptionInfo338);
            latentStyles1.Append(latentStyleExceptionInfo339);
            latentStyles1.Append(latentStyleExceptionInfo340);
            latentStyles1.Append(latentStyleExceptionInfo341);
            latentStyles1.Append(latentStyleExceptionInfo342);
            latentStyles1.Append(latentStyleExceptionInfo343);
            latentStyles1.Append(latentStyleExceptionInfo344);
            latentStyles1.Append(latentStyleExceptionInfo345);
            latentStyles1.Append(latentStyleExceptionInfo346);
            latentStyles1.Append(latentStyleExceptionInfo347);
            latentStyles1.Append(latentStyleExceptionInfo348);
            latentStyles1.Append(latentStyleExceptionInfo349);
            latentStyles1.Append(latentStyleExceptionInfo350);
            latentStyles1.Append(latentStyleExceptionInfo351);
            latentStyles1.Append(latentStyleExceptionInfo352);
            latentStyles1.Append(latentStyleExceptionInfo353);
            latentStyles1.Append(latentStyleExceptionInfo354);
            latentStyles1.Append(latentStyleExceptionInfo355);
            latentStyles1.Append(latentStyleExceptionInfo356);
            latentStyles1.Append(latentStyleExceptionInfo357);
            latentStyles1.Append(latentStyleExceptionInfo358);
            latentStyles1.Append(latentStyleExceptionInfo359);
            latentStyles1.Append(latentStyleExceptionInfo360);
            latentStyles1.Append(latentStyleExceptionInfo361);
            latentStyles1.Append(latentStyleExceptionInfo362);
            latentStyles1.Append(latentStyleExceptionInfo363);
            latentStyles1.Append(latentStyleExceptionInfo364);
            latentStyles1.Append(latentStyleExceptionInfo365);
            latentStyles1.Append(latentStyleExceptionInfo366);
            latentStyles1.Append(latentStyleExceptionInfo367);
            latentStyles1.Append(latentStyleExceptionInfo368);
            latentStyles1.Append(latentStyleExceptionInfo369);
            latentStyles1.Append(latentStyleExceptionInfo370);
            latentStyles1.Append(latentStyleExceptionInfo371);
            latentStyles1.Append(latentStyleExceptionInfo372);
            latentStyles1.Append(latentStyleExceptionInfo373);

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines9 = new SpacingBetweenLines() { After = "100", Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation3 = new Indentation() { FirstLine = "500" };
            Justification justification9 = new Justification() { Val = JustificationValues.Both };

            styleParagraphProperties1.Append(spacingBetweenLines9);
            styleParagraphProperties1.Append(indentation3);
            styleParagraphProperties1.Append(justification9);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts13 = new RunFonts() { HighAnsi = "Calibri Light", ComplexScript = "Calibri Light" };
            Color color21 = new Color() { Val = "000000" };

            styleRunProperties1.Append(runFonts13);
            styleRunProperties1.Append(color21);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);


            Style style2 = new Style() { Type = StyleValues.Paragraph, StyleId = "Ttulo1" };
            StyleName styleName2 = new StyleName() { Val = "heading 1" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            UIPriority uIPriority1 = new UIPriority() { Val = 9 };
            PrimaryStyle primaryStyle2 = new PrimaryStyle();
            Rsid rsid1 = new Rsid() { Val = "006054CD" };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            KeepLines keepLines7 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines10 = new SpacingBetweenLines() { Before = "240", After = "0", Line = "259", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation4 = new Indentation() { FirstLine = "0" };
            Justification justification10 = new Justification() { Val = JustificationValues.Left };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 0 };

            styleParagraphProperties2.Append(keepNext1);
            styleParagraphProperties2.Append(keepLines7);
            styleParagraphProperties2.Append(spacingBetweenLines10);
            styleParagraphProperties2.Append(indentation4);
            styleParagraphProperties2.Append(justification10);
            styleParagraphProperties2.Append(outlineLevel1);

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            RunFonts runFonts14 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color22 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize32 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "32" };
            Languages languages10 = new Languages() { EastAsia = "en-US" };

            styleRunProperties2.Append(runFonts14);
            styleRunProperties2.Append(color22);
            styleRunProperties2.Append(fontSize32);
            styleRunProperties2.Append(fontSizeComplexScript18);
            styleRunProperties2.Append(languages10);

            style2.Append(styleName2);
            style2.Append(basedOn1);
            style2.Append(nextParagraphStyle1);
            style2.Append(uIPriority1);
            style2.Append(primaryStyle2);
            style2.Append(rsid1);
            style2.Append(styleParagraphProperties2);
            style2.Append(styleRunProperties2);

            Style style3 = new Style() { Type = StyleValues.Paragraph, StyleId = "Ttulo2" };
            StyleName styleName3 = new StyleName() { Val = "heading 2" };
            BasedOn basedOn2 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle2 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Ttulo2Char1" };
            UIPriority uIPriority2 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle3 = new PrimaryStyle();
            Rsid rsid2 = new Rsid() { Val = "00033F32" };

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();
            KeepNext keepNext2 = new KeepNext();
            KeepLines keepLines8 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines11 = new SpacingBetweenLines() { Before = "40", After = "0" };
            OutlineLevel outlineLevel2 = new OutlineLevel() { Val = 1 };

            styleParagraphProperties3.Append(keepNext2);
            styleParagraphProperties3.Append(keepLines8);
            styleParagraphProperties3.Append(spacingBetweenLines11);
            styleParagraphProperties3.Append(outlineLevel2);

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            RunFonts runFonts15 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color23 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize33 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "26" };

            styleRunProperties3.Append(runFonts15);
            styleRunProperties3.Append(color23);
            styleRunProperties3.Append(fontSize33);
            styleRunProperties3.Append(fontSizeComplexScript19);

            style3.Append(styleName3);
            style3.Append(basedOn2);
            style3.Append(nextParagraphStyle2);
            style3.Append(linkedStyle1);
            style3.Append(uIPriority2);
            style3.Append(semiHidden1);
            style3.Append(unhideWhenUsed1);
            style3.Append(primaryStyle3);
            style3.Append(rsid2);
            style3.Append(styleParagraphProperties3);
            style3.Append(styleRunProperties3);

            Style style4 = new Style() { Type = StyleValues.Paragraph, StyleId = "Ttulo3" };
            StyleName styleName4 = new StyleName() { Val = "heading 3" };
            BasedOn basedOn3 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle3 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "Ttulo3Char1" };
            UIPriority uIPriority3 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden2 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed2 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle4 = new PrimaryStyle();
            Rsid rsid3 = new Rsid() { Val = "00033F32" };

            StyleParagraphProperties styleParagraphProperties4 = new StyleParagraphProperties();
            KeepNext keepNext3 = new KeepNext();
            KeepLines keepLines9 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines12 = new SpacingBetweenLines() { Before = "40", After = "0" };
            OutlineLevel outlineLevel3 = new OutlineLevel() { Val = 2 };

            styleParagraphProperties4.Append(keepNext3);
            styleParagraphProperties4.Append(keepLines9);
            styleParagraphProperties4.Append(spacingBetweenLines12);
            styleParagraphProperties4.Append(outlineLevel3);

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            RunFonts runFonts16 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color24 = new Color() { Val = "1F3763", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "7F" };
            FontSize fontSize34 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties4.Append(runFonts16);
            styleRunProperties4.Append(color24);
            styleRunProperties4.Append(fontSize34);
            styleRunProperties4.Append(fontSizeComplexScript20);

            style4.Append(styleName4);
            style4.Append(basedOn3);
            style4.Append(nextParagraphStyle3);
            style4.Append(linkedStyle2);
            style4.Append(uIPriority3);
            style4.Append(semiHidden2);
            style4.Append(unhideWhenUsed2);
            style4.Append(primaryStyle4);
            style4.Append(rsid3);
            style4.Append(styleParagraphProperties4);
            style4.Append(styleRunProperties4);

            Style style5 = new Style() { Type = StyleValues.Paragraph, StyleId = "Ttulo4" };
            StyleName styleName5 = new StyleName() { Val = "heading 4" };
            BasedOn basedOn4 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle4 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle3 = new LinkedStyle() { Val = "Ttulo4Char1" };
            UIPriority uIPriority4 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden3 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed3 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle5 = new PrimaryStyle();
            Rsid rsid4 = new Rsid() { Val = "00033F32" };

            StyleParagraphProperties styleParagraphProperties5 = new StyleParagraphProperties();
            KeepNext keepNext4 = new KeepNext();
            KeepLines keepLines10 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines13 = new SpacingBetweenLines() { Before = "40", After = "0" };
            OutlineLevel outlineLevel4 = new OutlineLevel() { Val = 3 };

            styleParagraphProperties5.Append(keepNext4);
            styleParagraphProperties5.Append(keepLines10);
            styleParagraphProperties5.Append(spacingBetweenLines13);
            styleParagraphProperties5.Append(outlineLevel4);

            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            RunFonts runFonts17 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Italic italic1 = new Italic();
            ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
            Color color25 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };

            styleRunProperties5.Append(runFonts17);
            styleRunProperties5.Append(italic1);
            styleRunProperties5.Append(italicComplexScript1);
            styleRunProperties5.Append(color25);

            style5.Append(styleName5);
            style5.Append(basedOn4);
            style5.Append(nextParagraphStyle4);
            style5.Append(linkedStyle3);
            style5.Append(uIPriority4);
            style5.Append(semiHidden3);
            style5.Append(unhideWhenUsed3);
            style5.Append(primaryStyle5);
            style5.Append(rsid4);
            style5.Append(styleParagraphProperties5);
            style5.Append(styleRunProperties5);

            Style style6 = new Style() { Type = StyleValues.Character, StyleId = "Fontepargpadro", Default = true };
            StyleName styleName6 = new StyleName() { Val = "Default Paragraph Font" };
            UIPriority uIPriority5 = new UIPriority() { Val = 1 };
            SemiHidden semiHidden4 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed4 = new UnhideWhenUsed();

            style6.Append(styleName6);
            style6.Append(uIPriority5);
            style6.Append(semiHidden4);
            style6.Append(unhideWhenUsed4);

            Style style7 = new Style() { Type = StyleValues.Table, StyleId = "Tabelanormal", Default = true };
            StyleName styleName7 = new StyleName() { Val = "Normal Table" };
            UIPriority uIPriority6 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden5 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed5 = new UnhideWhenUsed();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation1 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault3 = new TableCellMarginDefault();
            TopMargin topMargin5 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin3 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin5 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin3 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault3.Append(topMargin5);
            tableCellMarginDefault3.Append(tableCellLeftMargin3);
            tableCellMarginDefault3.Append(bottomMargin5);
            tableCellMarginDefault3.Append(tableCellRightMargin3);

            styleTableProperties1.Append(tableIndentation1);
            styleTableProperties1.Append(tableCellMarginDefault3);

            style7.Append(styleName7);
            style7.Append(uIPriority6);
            style7.Append(semiHidden5);
            style7.Append(unhideWhenUsed5);
            style7.Append(styleTableProperties1);

            Style style8 = new Style() { Type = StyleValues.Numbering, StyleId = "Semlista", Default = true };
            StyleName styleName8 = new StyleName() { Val = "No List" };
            UIPriority uIPriority7 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden6 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed6 = new UnhideWhenUsed();

            style8.Append(styleName8);
            style8.Append(uIPriority7);
            style8.Append(semiHidden6);
            style8.Append(unhideWhenUsed6);

            Style style9 = new Style() { Type = StyleValues.Paragraph, StyleId = "Ttulo11", CustomStyle = true };
            StyleName styleName9 = new StyleName() { Val = "Título 11" };
            BasedOn basedOn5 = new BasedOn() { Val = "Ttulo1" };
            NextParagraphStyle nextParagraphStyle5 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle4 = new LinkedStyle() { Val = "Ttulo1Char" };
            UIPriority uIPriority8 = new UIPriority() { Val = 9 };
            PrimaryStyle primaryStyle6 = new PrimaryStyle();
            Rsid rsid5 = new Rsid() { Val = "00B54F93" };

            StyleParagraphProperties styleParagraphProperties6 = new StyleParagraphProperties();

            NumberingProperties numberingProperties1 = new NumberingProperties();
            NumberingId numberingId1 = new NumberingId() { Val = 1 };

            numberingProperties1.Append(numberingId1);

            ParagraphBorders paragraphBorders1 = new ParagraphBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "4F81BD", Size = (UInt32Value)24U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "4F81BD", Size = (UInt32Value)24U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "4F81BD", Size = (UInt32Value)24U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "4F81BD", Size = (UInt32Value)24U, Space = (UInt32Value)0U };

            paragraphBorders1.Append(topBorder1);
            paragraphBorders1.Append(leftBorder2);
            paragraphBorders1.Append(bottomBorder1);
            paragraphBorders1.Append(rightBorder1);
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "4F81BD" };
            SpacingBetweenLines spacingBetweenLines14 = new SpacingBetweenLines() { Before = "0", After = "300" };

            styleParagraphProperties6.Append(numberingProperties1);
            styleParagraphProperties6.Append(paragraphBorders1);
            styleParagraphProperties6.Append(shading3);
            styleParagraphProperties6.Append(spacingBetweenLines14);

            StyleRunProperties styleRunProperties6 = new StyleRunProperties();
            RunFonts runFonts18 = new RunFonts() { Ascii = "Calibri Light", HighAnsi = "Calibri Light", EastAsia = "Times New Roman" };
            SmallCaps smallCaps1 = new SmallCaps();
            Color color26 = new Color() { Val = "FFFFFF" };
            FontSize fontSize35 = new FontSize() { Val = "40" };
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "40" };

            styleRunProperties6.Append(runFonts18);
            styleRunProperties6.Append(smallCaps1);
            styleRunProperties6.Append(color26);
            styleRunProperties6.Append(fontSize35);
            styleRunProperties6.Append(fontSizeComplexScript21);

            style9.Append(styleName9);
            style9.Append(basedOn5);
            style9.Append(nextParagraphStyle5);
            style9.Append(linkedStyle4);
            style9.Append(uIPriority8);
            style9.Append(primaryStyle6);
            style9.Append(rsid5);
            style9.Append(styleParagraphProperties6);
            style9.Append(styleRunProperties6);

            Style style10 = new Style() { Type = StyleValues.Paragraph, StyleId = "Ttulo21", CustomStyle = true };
            StyleName styleName10 = new StyleName() { Val = "Título 21" };
            BasedOn basedOn6 = new BasedOn() { Val = "Ttulo2" };
            NextParagraphStyle nextParagraphStyle6 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle5 = new LinkedStyle() { Val = "Ttulo2Char" };
            UIPriority uIPriority9 = new UIPriority() { Val = 9 };
            PrimaryStyle primaryStyle7 = new PrimaryStyle();
            Rsid rsid6 = new Rsid() { Val = "00B54F93" };

            StyleParagraphProperties styleParagraphProperties7 = new StyleParagraphProperties();

            NumberingProperties numberingProperties2 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 1 };
            NumberingId numberingId2 = new NumberingId() { Val = 1 };

            numberingProperties2.Append(numberingLevelReference1);
            numberingProperties2.Append(numberingId2);

            ParagraphBorders paragraphBorders2 = new ParagraphBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "DBE5F1", Size = (UInt32Value)24U, Space = (UInt32Value)0U };
            LeftBorder leftBorder3 = new LeftBorder() { Val = BorderValues.Single, Color = "DBE5F1", Size = (UInt32Value)24U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "DBE5F1", Size = (UInt32Value)24U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "DBE5F1", Size = (UInt32Value)24U, Space = (UInt32Value)0U };

            paragraphBorders2.Append(topBorder2);
            paragraphBorders2.Append(leftBorder3);
            paragraphBorders2.Append(bottomBorder2);
            paragraphBorders2.Append(rightBorder2);
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "DBE5F1" };
            SpacingBetweenLines spacingBetweenLines15 = new SpacingBetweenLines() { Before = "450", After = "150" };

            styleParagraphProperties7.Append(numberingProperties2);
            styleParagraphProperties7.Append(paragraphBorders2);
            styleParagraphProperties7.Append(shading4);
            styleParagraphProperties7.Append(spacingBetweenLines15);

            StyleRunProperties styleRunProperties7 = new StyleRunProperties();
            RunFonts runFonts19 = new RunFonts() { Ascii = "Calibri Light", HighAnsi = "Calibri Light", EastAsia = "Times New Roman" };
            Color color27 = new Color() { Val = "0070C0" };
            FontSize fontSize36 = new FontSize() { Val = "34" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "34" };

            styleRunProperties7.Append(runFonts19);
            styleRunProperties7.Append(color27);
            styleRunProperties7.Append(fontSize36);
            styleRunProperties7.Append(fontSizeComplexScript22);

            style10.Append(styleName10);
            style10.Append(basedOn6);
            style10.Append(nextParagraphStyle6);
            style10.Append(linkedStyle5);
            style10.Append(uIPriority9);
            style10.Append(primaryStyle7);
            style10.Append(rsid6);
            style10.Append(styleParagraphProperties7);
            style10.Append(styleRunProperties7);

            Style style11 = new Style() { Type = StyleValues.Paragraph, StyleId = "Ttulo31", CustomStyle = true };
            StyleName styleName11 = new StyleName() { Val = "Título 31" };
            BasedOn basedOn7 = new BasedOn() { Val = "Ttulo3" };
            NextParagraphStyle nextParagraphStyle7 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle6 = new LinkedStyle() { Val = "Ttulo3Char" };
            UIPriority uIPriority10 = new UIPriority() { Val = 9 };
            PrimaryStyle primaryStyle8 = new PrimaryStyle();
            Rsid rsid7 = new Rsid() { Val = "00B54F93" };

            StyleParagraphProperties styleParagraphProperties8 = new StyleParagraphProperties();

            NumberingProperties numberingProperties3 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference2 = new NumberingLevelReference() { Val = 2 };
            NumberingId numberingId3 = new NumberingId() { Val = 1 };

            numberingProperties3.Append(numberingLevelReference2);
            numberingProperties3.Append(numberingId3);

            ParagraphBorders paragraphBorders3 = new ParagraphBorders();
            TopBorder topBorder3 = new TopBorder() { Val = BorderValues.Single, Color = "4F81BD", Size = (UInt32Value)6U, Space = (UInt32Value)2U };
            LeftBorder leftBorder4 = new LeftBorder() { Val = BorderValues.Single, Color = "4F81BD", Size = (UInt32Value)6U, Space = (UInt32Value)2U };

            paragraphBorders3.Append(topBorder3);
            paragraphBorders3.Append(leftBorder4);
            SpacingBetweenLines spacingBetweenLines16 = new SpacingBetweenLines() { Before = "450", After = "150" };

            styleParagraphProperties8.Append(numberingProperties3);
            styleParagraphProperties8.Append(paragraphBorders3);
            styleParagraphProperties8.Append(spacingBetweenLines16);

            StyleRunProperties styleRunProperties8 = new StyleRunProperties();
            RunFonts runFonts20 = new RunFonts() { Ascii = "Calibri Light", HighAnsi = "Calibri Light", EastAsia = "Times New Roman" };
            Italic italic2 = new Italic();
            Color color28 = new Color() { Val = "1F497D" };
            FontSize fontSize37 = new FontSize() { Val = "30" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "36" };

            styleRunProperties8.Append(runFonts20);
            styleRunProperties8.Append(italic2);
            styleRunProperties8.Append(color28);
            styleRunProperties8.Append(fontSize37);
            styleRunProperties8.Append(fontSizeComplexScript23);

            style11.Append(styleName11);
            style11.Append(basedOn7);
            style11.Append(nextParagraphStyle7);
            style11.Append(linkedStyle6);
            style11.Append(uIPriority10);
            style11.Append(primaryStyle8);
            style11.Append(rsid7);
            style11.Append(styleParagraphProperties8);
            style11.Append(styleRunProperties8);

            Style style12 = new Style() { Type = StyleValues.Paragraph, StyleId = "Ttulo41", CustomStyle = true };
            StyleName styleName12 = new StyleName() { Val = "Título 41" };
            BasedOn basedOn8 = new BasedOn() { Val = "Ttulo4" };
            NextParagraphStyle nextParagraphStyle8 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle7 = new LinkedStyle() { Val = "Ttulo4Char" };
            UIPriority uIPriority11 = new UIPriority() { Val = 9 };
            PrimaryStyle primaryStyle9 = new PrimaryStyle();
            Rsid rsid8 = new Rsid() { Val = "00B54F93" };

            StyleParagraphProperties styleParagraphProperties9 = new StyleParagraphProperties();

            NumberingProperties numberingProperties4 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference3 = new NumberingLevelReference() { Val = 3 };
            NumberingId numberingId4 = new NumberingId() { Val = 1 };

            numberingProperties4.Append(numberingLevelReference3);
            numberingProperties4.Append(numberingId4);
            SpacingBetweenLines spacingBetweenLines17 = new SpacingBetweenLines() { Before = "450", After = "150" };

            styleParagraphProperties9.Append(numberingProperties4);
            styleParagraphProperties9.Append(spacingBetweenLines17);

            StyleRunProperties styleRunProperties9 = new StyleRunProperties();
            RunFonts runFonts21 = new RunFonts() { Ascii = "Calibri Light", HighAnsi = "Calibri Light", EastAsia = "Times New Roman" };
            Italic italic3 = new Italic() { Val = false };
            ItalicComplexScript italicComplexScript2 = new ItalicComplexScript() { Val = false };
            Color color29 = new Color() { Val = "0070C0" };
            FontSize fontSize38 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "26" };

            styleRunProperties9.Append(runFonts21);
            styleRunProperties9.Append(italic3);
            styleRunProperties9.Append(italicComplexScript2);
            styleRunProperties9.Append(color29);
            styleRunProperties9.Append(fontSize38);
            styleRunProperties9.Append(fontSizeComplexScript24);

            style12.Append(styleName12);
            style12.Append(basedOn8);
            style12.Append(nextParagraphStyle8);
            style12.Append(linkedStyle7);
            style12.Append(uIPriority11);
            style12.Append(primaryStyle9);
            style12.Append(rsid8);
            style12.Append(styleParagraphProperties9);
            style12.Append(styleRunProperties9);

            Style style13 = new Style() { Type = StyleValues.Paragraph, StyleId = "Ttulo51", CustomStyle = true };
            StyleName styleName13 = new StyleName() { Val = "Título 51" };
            BasedOn basedOn9 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle9 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle8 = new LinkedStyle() { Val = "Ttulo5Char" };
            UIPriority uIPriority12 = new UIPriority() { Val = 9 };
            UnhideWhenUsed unhideWhenUsed7 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle10 = new PrimaryStyle();
            Rsid rsid9 = new Rsid() { Val = "00B54F93" };

            StyleParagraphProperties styleParagraphProperties10 = new StyleParagraphProperties();
            KeepNext keepNext5 = new KeepNext();
            KeepLines keepLines11 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines18 = new SpacingBetweenLines() { Before = "40", After = "0" };
            Indentation indentation5 = new Indentation() { Start = "600", Hanging = "600" };
            OutlineLevel outlineLevel5 = new OutlineLevel() { Val = 4 };

            styleParagraphProperties10.Append(keepNext5);
            styleParagraphProperties10.Append(keepLines11);
            styleParagraphProperties10.Append(spacingBetweenLines18);
            styleParagraphProperties10.Append(indentation5);
            styleParagraphProperties10.Append(outlineLevel5);

            StyleRunProperties styleRunProperties10 = new StyleRunProperties();
            RunFonts runFonts22 = new RunFonts() { EastAsia = "Times New Roman" };
            Italic italic4 = new Italic();
            ItalicComplexScript italicComplexScript3 = new ItalicComplexScript();
            Color color30 = new Color() { Val = "000077" };
            FontSize fontSize39 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties10.Append(runFonts22);
            styleRunProperties10.Append(italic4);
            styleRunProperties10.Append(italicComplexScript3);
            styleRunProperties10.Append(color30);
            styleRunProperties10.Append(fontSize39);
            styleRunProperties10.Append(fontSizeComplexScript25);

            style13.Append(styleName13);
            style13.Append(basedOn9);
            style13.Append(nextParagraphStyle9);
            style13.Append(linkedStyle8);
            style13.Append(uIPriority12);
            style13.Append(unhideWhenUsed7);
            style13.Append(primaryStyle10);
            style13.Append(rsid9);
            style13.Append(styleParagraphProperties10);
            style13.Append(styleRunProperties10);

            Style style14 = new Style() { Type = StyleValues.Paragraph, StyleId = "Ttulo61", CustomStyle = true };
            StyleName styleName14 = new StyleName() { Val = "Título 61" };
            BasedOn basedOn10 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle10 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle9 = new LinkedStyle() { Val = "Ttulo6Char" };
            UIPriority uIPriority13 = new UIPriority() { Val = 9 };
            UnhideWhenUsed unhideWhenUsed8 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle11 = new PrimaryStyle();
            Rsid rsid10 = new Rsid() { Val = "00B54F93" };

            StyleParagraphProperties styleParagraphProperties11 = new StyleParagraphProperties();
            KeepNext keepNext6 = new KeepNext();
            KeepLines keepLines12 = new KeepLines();

            NumberingProperties numberingProperties5 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference4 = new NumberingLevelReference() { Val = 5 };
            NumberingId numberingId5 = new NumberingId() { Val = 1 };

            numberingProperties5.Append(numberingLevelReference4);
            numberingProperties5.Append(numberingId5);
            SpacingBetweenLines spacingBetweenLines19 = new SpacingBetweenLines() { Before = "40", After = "0" };
            OutlineLevel outlineLevel6 = new OutlineLevel() { Val = 5 };

            styleParagraphProperties11.Append(keepNext6);
            styleParagraphProperties11.Append(keepLines12);
            styleParagraphProperties11.Append(numberingProperties5);
            styleParagraphProperties11.Append(spacingBetweenLines19);
            styleParagraphProperties11.Append(outlineLevel6);

            StyleRunProperties styleRunProperties11 = new StyleRunProperties();
            RunFonts runFonts23 = new RunFonts() { EastAsia = "Times New Roman" };
            Italic italic5 = new Italic();
            ItalicComplexScript italicComplexScript4 = new ItalicComplexScript();
            Color color31 = new Color() { Val = "000077" };
            FontSize fontSize40 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "22" };

            styleRunProperties11.Append(runFonts23);
            styleRunProperties11.Append(italic5);
            styleRunProperties11.Append(italicComplexScript4);
            styleRunProperties11.Append(color31);
            styleRunProperties11.Append(fontSize40);
            styleRunProperties11.Append(fontSizeComplexScript26);

            style14.Append(styleName14);
            style14.Append(basedOn10);
            style14.Append(nextParagraphStyle10);
            style14.Append(linkedStyle9);
            style14.Append(uIPriority13);
            style14.Append(unhideWhenUsed8);
            style14.Append(primaryStyle11);
            style14.Append(rsid10);
            style14.Append(styleParagraphProperties11);
            style14.Append(styleRunProperties11);

            Style style15 = new Style() { Type = StyleValues.Paragraph, StyleId = "Ttulo71", CustomStyle = true };
            StyleName styleName15 = new StyleName() { Val = "Título 71" };
            BasedOn basedOn11 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle11 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle10 = new LinkedStyle() { Val = "Ttulo7Char" };
            UIPriority uIPriority14 = new UIPriority() { Val = 9 };
            UnhideWhenUsed unhideWhenUsed9 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle12 = new PrimaryStyle();
            Rsid rsid11 = new Rsid() { Val = "00B54F93" };

            StyleParagraphProperties styleParagraphProperties12 = new StyleParagraphProperties();
            KeepNext keepNext7 = new KeepNext();
            KeepLines keepLines13 = new KeepLines();

            NumberingProperties numberingProperties6 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference5 = new NumberingLevelReference() { Val = 6 };
            NumberingId numberingId6 = new NumberingId() { Val = 1 };

            numberingProperties6.Append(numberingLevelReference5);
            numberingProperties6.Append(numberingId6);
            SpacingBetweenLines spacingBetweenLines20 = new SpacingBetweenLines() { Before = "40", After = "0" };
            OutlineLevel outlineLevel7 = new OutlineLevel() { Val = 6 };

            styleParagraphProperties12.Append(keepNext7);
            styleParagraphProperties12.Append(keepLines13);
            styleParagraphProperties12.Append(numberingProperties6);
            styleParagraphProperties12.Append(spacingBetweenLines20);
            styleParagraphProperties12.Append(outlineLevel7);

            StyleRunProperties styleRunProperties12 = new StyleRunProperties();
            RunFonts runFonts24 = new RunFonts() { EastAsia = "Times New Roman" };
            Italic italic6 = new Italic();
            ItalicComplexScript italicComplexScript5 = new ItalicComplexScript();
            Color color32 = new Color() { Val = "000077" };
            FontSize fontSize41 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "22" };

            styleRunProperties12.Append(runFonts24);
            styleRunProperties12.Append(italic6);
            styleRunProperties12.Append(italicComplexScript5);
            styleRunProperties12.Append(color32);
            styleRunProperties12.Append(fontSize41);
            styleRunProperties12.Append(fontSizeComplexScript27);

            style15.Append(styleName15);
            style15.Append(basedOn11);
            style15.Append(nextParagraphStyle11);
            style15.Append(linkedStyle10);
            style15.Append(uIPriority14);
            style15.Append(unhideWhenUsed9);
            style15.Append(primaryStyle12);
            style15.Append(rsid11);
            style15.Append(styleParagraphProperties12);
            style15.Append(styleRunProperties12);

            Style style16 = new Style() { Type = StyleValues.Character, StyleId = "Ttulo1Char", CustomStyle = true };
            StyleName styleName16 = new StyleName() { Val = "Título 1 Char" };
            LinkedStyle linkedStyle11 = new LinkedStyle() { Val = "Ttulo11" };
            UIPriority uIPriority15 = new UIPriority() { Val = 9 };
            Rsid rsid12 = new Rsid() { Val = "00033F32" };

            StyleRunProperties styleRunProperties13 = new StyleRunProperties();
            RunFonts runFonts25 = new RunFonts() { HighAnsi = "Calibri Light", EastAsia = "Times New Roman", ComplexScriptTheme = ThemeFontValues.MajorBidi };
            SmallCaps smallCaps2 = new SmallCaps();
            Color color33 = new Color() { Val = "FFFFFF" };
            FontSize fontSize42 = new FontSize() { Val = "40" };
            FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "40" };
            Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "4F81BD" };
            Languages languages11 = new Languages() { EastAsia = "en-US" };

            styleRunProperties13.Append(runFonts25);
            styleRunProperties13.Append(smallCaps2);
            styleRunProperties13.Append(color33);
            styleRunProperties13.Append(fontSize42);
            styleRunProperties13.Append(fontSizeComplexScript28);
            styleRunProperties13.Append(shading5);
            styleRunProperties13.Append(languages11);

            style16.Append(styleName16);
            style16.Append(linkedStyle11);
            style16.Append(uIPriority15);
            style16.Append(rsid12);
            style16.Append(styleRunProperties13);

            Style style17 = new Style() { Type = StyleValues.Character, StyleId = "Ttulo2Char", CustomStyle = true };
            StyleName styleName17 = new StyleName() { Val = "Título 2 Char" };
            LinkedStyle linkedStyle12 = new LinkedStyle() { Val = "Ttulo21" };
            UIPriority uIPriority16 = new UIPriority() { Val = 9 };
            Rsid rsid13 = new Rsid() { Val = "00033F32" };

            StyleRunProperties styleRunProperties14 = new StyleRunProperties();
            RunFonts runFonts26 = new RunFonts() { HighAnsi = "Calibri Light", EastAsia = "Times New Roman", ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color34 = new Color() { Val = "0070C0" };
            FontSize fontSize43 = new FontSize() { Val = "34" };
            FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "34" };
            Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "DBE5F1" };

            styleRunProperties14.Append(runFonts26);
            styleRunProperties14.Append(color34);
            styleRunProperties14.Append(fontSize43);
            styleRunProperties14.Append(fontSizeComplexScript29);
            styleRunProperties14.Append(shading6);

            style17.Append(styleName17);
            style17.Append(linkedStyle12);
            style17.Append(uIPriority16);
            style17.Append(rsid13);
            style17.Append(styleRunProperties14);

            Style style18 = new Style() { Type = StyleValues.Character, StyleId = "Ttulo3Char", CustomStyle = true };
            StyleName styleName18 = new StyleName() { Val = "Título 3 Char" };
            LinkedStyle linkedStyle13 = new LinkedStyle() { Val = "Ttulo31" };
            UIPriority uIPriority17 = new UIPriority() { Val = 9 };
            Rsid rsid14 = new Rsid() { Val = "00033F32" };

            StyleRunProperties styleRunProperties15 = new StyleRunProperties();
            RunFonts runFonts27 = new RunFonts() { HighAnsi = "Calibri Light", EastAsia = "Times New Roman", ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Italic italic7 = new Italic();
            Color color35 = new Color() { Val = "1F497D" };
            FontSize fontSize44 = new FontSize() { Val = "30" };
            FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "36" };

            styleRunProperties15.Append(runFonts27);
            styleRunProperties15.Append(italic7);
            styleRunProperties15.Append(color35);
            styleRunProperties15.Append(fontSize44);
            styleRunProperties15.Append(fontSizeComplexScript30);

            style18.Append(styleName18);
            style18.Append(linkedStyle13);
            style18.Append(uIPriority17);
            style18.Append(rsid14);
            style18.Append(styleRunProperties15);

            Style style19 = new Style() { Type = StyleValues.Character, StyleId = "Ttulo4Char", CustomStyle = true };
            StyleName styleName19 = new StyleName() { Val = "Título 4 Char" };
            LinkedStyle linkedStyle14 = new LinkedStyle() { Val = "Ttulo41" };
            UIPriority uIPriority18 = new UIPriority() { Val = 9 };
            Rsid rsid15 = new Rsid() { Val = "00033F32" };

            StyleRunProperties styleRunProperties16 = new StyleRunProperties();
            RunFonts runFonts28 = new RunFonts() { HighAnsi = "Calibri Light", EastAsia = "Times New Roman", ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color36 = new Color() { Val = "0070C0" };
            FontSize fontSize45 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "26" };

            styleRunProperties16.Append(runFonts28);
            styleRunProperties16.Append(color36);
            styleRunProperties16.Append(fontSize45);
            styleRunProperties16.Append(fontSizeComplexScript31);

            style19.Append(styleName19);
            style19.Append(linkedStyle14);
            style19.Append(uIPriority18);
            style19.Append(rsid15);
            style19.Append(styleRunProperties16);

            Style style20 = new Style() { Type = StyleValues.Character, StyleId = "Ttulo5Char", CustomStyle = true };
            StyleName styleName20 = new StyleName() { Val = "Título 5 Char" };
            LinkedStyle linkedStyle15 = new LinkedStyle() { Val = "Ttulo51" };
            UIPriority uIPriority19 = new UIPriority() { Val = 9 };
            Rsid rsid16 = new Rsid() { Val = "00B54F93" };

            StyleRunProperties styleRunProperties17 = new StyleRunProperties();
            RunFonts runFonts29 = new RunFonts() { HighAnsi = "Calibri Light", EastAsia = "Times New Roman", ComplexScript = "Calibri Light" };
            Italic italic8 = new Italic();
            ItalicComplexScript italicComplexScript6 = new ItalicComplexScript();
            Color color37 = new Color() { Val = "000077" };
            FontSize fontSize46 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties17.Append(runFonts29);
            styleRunProperties17.Append(italic8);
            styleRunProperties17.Append(italicComplexScript6);
            styleRunProperties17.Append(color37);
            styleRunProperties17.Append(fontSize46);
            styleRunProperties17.Append(fontSizeComplexScript32);

            style20.Append(styleName20);
            style20.Append(linkedStyle15);
            style20.Append(uIPriority19);
            style20.Append(rsid16);
            style20.Append(styleRunProperties17);

            Style style21 = new Style() { Type = StyleValues.Character, StyleId = "Ttulo6Char", CustomStyle = true };
            StyleName styleName21 = new StyleName() { Val = "Título 6 Char" };
            LinkedStyle linkedStyle16 = new LinkedStyle() { Val = "Ttulo61" };
            UIPriority uIPriority20 = new UIPriority() { Val = 9 };
            Rsid rsid17 = new Rsid() { Val = "00B54F93" };

            StyleRunProperties styleRunProperties18 = new StyleRunProperties();
            RunFonts runFonts30 = new RunFonts() { Ascii = "Calibri Light", HighAnsi = "Calibri Light", EastAsia = "Times New Roman" };
            Color color38 = new Color() { Val = "000077" };
            FontSize fontSize47 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "22" };
            Languages languages12 = new Languages() { Val = "pt-BR", EastAsia = "pt-BR", Bidi = "ar-SA" };

            styleRunProperties18.Append(runFonts30);
            styleRunProperties18.Append(color38);
            styleRunProperties18.Append(fontSize47);
            styleRunProperties18.Append(fontSizeComplexScript33);
            styleRunProperties18.Append(languages12);

            style21.Append(styleName21);
            style21.Append(linkedStyle16);
            style21.Append(uIPriority20);
            style21.Append(rsid17);
            style21.Append(styleRunProperties18);

            Style style22 = new Style() { Type = StyleValues.Character, StyleId = "Ttulo7Char", CustomStyle = true };
            StyleName styleName22 = new StyleName() { Val = "Título 7 Char" };
            LinkedStyle linkedStyle17 = new LinkedStyle() { Val = "Ttulo71" };
            UIPriority uIPriority21 = new UIPriority() { Val = 9 };
            Rsid rsid18 = new Rsid() { Val = "00B54F93" };

            StyleRunProperties styleRunProperties19 = new StyleRunProperties();
            RunFonts runFonts31 = new RunFonts() { Ascii = "Calibri Light", HighAnsi = "Calibri Light", EastAsia = "Times New Roman" };
            Italic italic9 = new Italic();
            ItalicComplexScript italicComplexScript7 = new ItalicComplexScript();
            Color color39 = new Color() { Val = "000077" };
            FontSize fontSize48 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "22" };
            Languages languages13 = new Languages() { Val = "pt-BR", EastAsia = "pt-BR", Bidi = "ar-SA" };

            styleRunProperties19.Append(runFonts31);
            styleRunProperties19.Append(italic9);
            styleRunProperties19.Append(italicComplexScript7);
            styleRunProperties19.Append(color39);
            styleRunProperties19.Append(fontSize48);
            styleRunProperties19.Append(fontSizeComplexScript34);
            styleRunProperties19.Append(languages13);

            style22.Append(styleName22);
            style22.Append(linkedStyle17);
            style22.Append(uIPriority21);
            style22.Append(rsid18);
            style22.Append(styleRunProperties19);




            Style style23 = new Style() { Type = StyleValues.Paragraph, StyleId = "NormalTextTable", CustomStyle = true };
            StyleName styleName23 = new StyleName() { Val = "NormalTextTable" };
            PrimaryStyle primaryStyle13 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties13 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines21 = new SpacingBetweenLines() { Before = "90", After = "90", Line = "180", LineRule = LineSpacingRuleValues.Auto };
            Justification justification11 = new Justification() { Val = JustificationValues.Both };

            styleParagraphProperties13.Append(spacingBetweenLines21);
            styleParagraphProperties13.Append(justification11);

            StyleRunProperties styleRunProperties20 = new StyleRunProperties();
            RunFonts runFonts32 = new RunFonts() { HighAnsi = "Calibri Light", ComplexScript = "Calibri Light" };
            Color color40 = new Color() { Val = "000000" };
            FontSize fontSize49 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "16" };

            styleRunProperties20.Append(runFonts32);
            styleRunProperties20.Append(color40);
            styleRunProperties20.Append(fontSize49);
            styleRunProperties20.Append(fontSizeComplexScript35);

            style23.Append(styleName23);
            style23.Append(primaryStyle13);
            style23.Append(styleParagraphProperties13);
            style23.Append(styleRunProperties20);

            Style style24 = new Style() { Type = StyleValues.Paragraph, StyleId = "CenteredTextTable", CustomStyle = true };
            StyleName styleName24 = new StyleName() { Val = "CenteredTextTable" };
            PrimaryStyle primaryStyle14 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties14 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines22 = new SpacingBetweenLines() { Before = "90", After = "90", Line = "180", LineRule = LineSpacingRuleValues.Auto };
            Justification justification12 = new Justification() { Val = JustificationValues.Center };

            styleParagraphProperties14.Append(spacingBetweenLines22);
            styleParagraphProperties14.Append(justification12);

            StyleRunProperties styleRunProperties21 = new StyleRunProperties();
            RunFonts runFonts33 = new RunFonts() { HighAnsi = "Calibri Light", ComplexScript = "Calibri Light" };
            Color color41 = new Color() { Val = "000000" };
            FontSize fontSize50 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "16" };

            styleRunProperties21.Append(runFonts33);
            styleRunProperties21.Append(color41);
            styleRunProperties21.Append(fontSize50);
            styleRunProperties21.Append(fontSizeComplexScript36);

            style24.Append(styleName24);
            style24.Append(primaryStyle14);
            style24.Append(styleParagraphProperties14);
            style24.Append(styleRunProperties21);

            Style style25 = new Style() { Type = StyleValues.Paragraph, StyleId = "LeftTextTable", CustomStyle = true };
            StyleName styleName25 = new StyleName() { Val = "LeftTextTable" };
            PrimaryStyle primaryStyle15 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties15 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines23 = new SpacingBetweenLines() { Before = "90", After = "90", Line = "180", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties15.Append(spacingBetweenLines23);

            StyleRunProperties styleRunProperties22 = new StyleRunProperties();
            RunFonts runFonts34 = new RunFonts() { HighAnsi = "Calibri Light", ComplexScript = "Calibri Light" };
            Color color42 = new Color() { Val = "000000" };
            FontSize fontSize51 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "16" };

            styleRunProperties22.Append(runFonts34);
            styleRunProperties22.Append(color42);
            styleRunProperties22.Append(fontSize51);
            styleRunProperties22.Append(fontSizeComplexScript37);

            style25.Append(styleName25);
            style25.Append(primaryStyle15);
            style25.Append(styleParagraphProperties15);
            style25.Append(styleRunProperties22);

            Style style26 = new Style() { Type = StyleValues.Paragraph, StyleId = "RightTextTable", CustomStyle = true };
            StyleName styleName26 = new StyleName() { Val = "RightTextTable" };
            PrimaryStyle primaryStyle16 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties16 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines24 = new SpacingBetweenLines() { Before = "90", After = "90", Line = "180", LineRule = LineSpacingRuleValues.Auto };
            Justification justification13 = new Justification() { Val = JustificationValues.Right };

            styleParagraphProperties16.Append(spacingBetweenLines24);
            styleParagraphProperties16.Append(justification13);

            StyleRunProperties styleRunProperties23 = new StyleRunProperties();
            RunFonts runFonts35 = new RunFonts() { HighAnsi = "Calibri Light", ComplexScript = "Calibri Light" };
            Color color43 = new Color() { Val = "000000" };
            FontSize fontSize52 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "16" };

            styleRunProperties23.Append(runFonts35);
            styleRunProperties23.Append(color43);
            styleRunProperties23.Append(fontSize52);
            styleRunProperties23.Append(fontSizeComplexScript38);

            style26.Append(styleName26);
            style26.Append(primaryStyle16);
            style26.Append(styleParagraphProperties16);
            style26.Append(styleRunProperties23);

            Style style27 = new Style() { Type = StyleValues.Paragraph, StyleId = "Textodecomentrio" };
            StyleName styleName27 = new StyleName() { Val = "annotation text" };
            BasedOn basedOn12 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle18 = new LinkedStyle() { Val = "TextodecomentrioChar" };
            UIPriority uIPriority22 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden7 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed10 = new UnhideWhenUsed();

            StyleParagraphProperties styleParagraphProperties17 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines25 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties17.Append(spacingBetweenLines25);

            style27.Append(styleName27);
            style27.Append(basedOn12);
            style27.Append(linkedStyle18);
            style27.Append(uIPriority22);
            style27.Append(semiHidden7);
            style27.Append(unhideWhenUsed10);
            style27.Append(styleParagraphProperties17);

            Style style28 = new Style() { Type = StyleValues.Character, StyleId = "TextodecomentrioChar", CustomStyle = true };
            StyleName styleName28 = new StyleName() { Val = "Texto de comentário Char" };
            BasedOn basedOn13 = new BasedOn() { Val = "Fontepargpadro" };
            LinkedStyle linkedStyle19 = new LinkedStyle() { Val = "Textodecomentrio" };
            UIPriority uIPriority23 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden8 = new SemiHidden();

            StyleRunProperties styleRunProperties24 = new StyleRunProperties();
            RunFonts runFonts36 = new RunFonts() { HighAnsi = "Calibri Light", ComplexScript = "Calibri Light" };
            Color color44 = new Color() { Val = "000000" };

            styleRunProperties24.Append(runFonts36);
            styleRunProperties24.Append(color44);

            style28.Append(styleName28);
            style28.Append(basedOn13);
            style28.Append(linkedStyle19);
            style28.Append(uIPriority23);
            style28.Append(semiHidden8);
            style28.Append(styleRunProperties24);

            Style style29 = new Style() { Type = StyleValues.Character, StyleId = "Refdecomentrio" };
            StyleName styleName29 = new StyleName() { Val = "annotation reference" };
            BasedOn basedOn14 = new BasedOn() { Val = "Fontepargpadro" };
            UIPriority uIPriority24 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden9 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed11 = new UnhideWhenUsed();

            StyleRunProperties styleRunProperties25 = new StyleRunProperties();
            FontSize fontSize53 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript39 = new FontSizeComplexScript() { Val = "16" };

            styleRunProperties25.Append(fontSize53);
            styleRunProperties25.Append(fontSizeComplexScript39);

            style29.Append(styleName29);
            style29.Append(basedOn14);
            style29.Append(uIPriority24);
            style29.Append(semiHidden9);
            style29.Append(unhideWhenUsed11);
            style29.Append(styleRunProperties25);

            Style style30 = new Style() { Type = StyleValues.Character, StyleId = "Ttulo1Char1", CustomStyle = true };
            StyleName styleName30 = new StyleName() { Val = "Título 1 Char1" };
            BasedOn basedOn15 = new BasedOn() { Val = "Fontepargpadro" };
            UIPriority uIPriority25 = new UIPriority() { Val = 9 };
            Rsid rsid19 = new Rsid() { Val = "006054CD" };

            StyleRunProperties styleRunProperties26 = new StyleRunProperties();
            RunFonts runFonts37 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color45 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize54 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript40 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties26.Append(runFonts37);
            styleRunProperties26.Append(color45);
            styleRunProperties26.Append(fontSize54);
            styleRunProperties26.Append(fontSizeComplexScript40);

            style30.Append(styleName30);
            style30.Append(basedOn15);
            style30.Append(uIPriority25);
            style30.Append(rsid19);
            style30.Append(styleRunProperties26);

            Style style31 = new Style() { Type = StyleValues.Paragraph, StyleId = "SemEspaamento" };
            StyleName styleName31 = new StyleName() { Val = "No Spacing" };
            LinkedStyle linkedStyle20 = new LinkedStyle() { Val = "SemEspaamentoChar" };
            UIPriority uIPriority26 = new UIPriority() { Val = 1 };
            PrimaryStyle primaryStyle17 = new PrimaryStyle();
            Rsid rsid20 = new Rsid() { Val = "006054CD" };

            StyleRunProperties styleRunProperties27 = new StyleRunProperties();
            RunFonts runFonts38 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorEastAsia, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            FontSize fontSize55 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "22" };

            styleRunProperties27.Append(runFonts38);
            styleRunProperties27.Append(fontSize55);
            styleRunProperties27.Append(fontSizeComplexScript41);

            style31.Append(styleName31);
            style31.Append(linkedStyle20);
            style31.Append(uIPriority26);
            style31.Append(primaryStyle17);
            style31.Append(rsid20);
            style31.Append(styleRunProperties27);

            Style style32 = new Style() { Type = StyleValues.Character, StyleId = "SemEspaamentoChar", CustomStyle = true };
            StyleName styleName32 = new StyleName() { Val = "Sem Espaçamento Char" };
            BasedOn basedOn16 = new BasedOn() { Val = "Fontepargpadro" };
            LinkedStyle linkedStyle21 = new LinkedStyle() { Val = "SemEspaamento" };
            UIPriority uIPriority27 = new UIPriority() { Val = 1 };
            Rsid rsid21 = new Rsid() { Val = "006054CD" };

            StyleRunProperties styleRunProperties28 = new StyleRunProperties();
            RunFonts runFonts39 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorEastAsia, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            FontSize fontSize56 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "22" };

            styleRunProperties28.Append(runFonts39);
            styleRunProperties28.Append(fontSize56);
            styleRunProperties28.Append(fontSizeComplexScript42);

            style32.Append(styleName32);
            style32.Append(basedOn16);
            style32.Append(linkedStyle21);
            style32.Append(uIPriority27);
            style32.Append(rsid21);
            style32.Append(styleRunProperties28);

            Style style33 = new Style() { Type = StyleValues.Paragraph, StyleId = "CabealhodoSumrio" };
            StyleName styleName33 = new StyleName() { Val = "TOC Heading" };
            BasedOn basedOn17 = new BasedOn() { Val = "Ttulo1" };
            NextParagraphStyle nextParagraphStyle12 = new NextParagraphStyle() { Val = "Normal" };
            UIPriority uIPriority28 = new UIPriority() { Val = 39 };
            UnhideWhenUsed unhideWhenUsed12 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle18 = new PrimaryStyle();
            Rsid rsid22 = new Rsid() { Val = "006054CD" };

            StyleParagraphProperties styleParagraphProperties18 = new StyleParagraphProperties();
            OutlineLevel outlineLevel8 = new OutlineLevel() { Val = 9 };

            styleParagraphProperties18.Append(outlineLevel8);

            StyleRunProperties styleRunProperties29 = new StyleRunProperties();
            Languages languages14 = new Languages() { EastAsia = "pt-BR" };

            styleRunProperties29.Append(languages14);

            style33.Append(styleName33);
            style33.Append(basedOn17);
            style33.Append(nextParagraphStyle12);
            style33.Append(uIPriority28);
            style33.Append(unhideWhenUsed12);
            style33.Append(primaryStyle18);
            style33.Append(rsid22);
            style33.Append(styleParagraphProperties18);
            style33.Append(styleRunProperties29);

            Style style34 = new Style() { Type = StyleValues.Paragraph, StyleId = "Rodap" };
            StyleName styleName34 = new StyleName() { Val = "footer" };
            BasedOn basedOn18 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle22 = new LinkedStyle() { Val = "RodapChar" };
            UIPriority uIPriority29 = new UIPriority() { Val = 99 };
            UnhideWhenUsed unhideWhenUsed13 = new UnhideWhenUsed();
            Rsid rsid23 = new Rsid() { Val = "006054CD" };

            StyleParagraphProperties styleParagraphProperties19 = new StyleParagraphProperties();

            Tabs tabs4 = new Tabs();
            TabStop tabStop7 = new TabStop() { Val = TabStopValues.Center, Position = 4252 };
            TabStop tabStop8 = new TabStop() { Val = TabStopValues.Right, Position = 8504 };

            tabs4.Append(tabStop7);
            tabs4.Append(tabStop8);
            SpacingBetweenLines spacingBetweenLines26 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation6 = new Indentation() { FirstLine = "0" };
            Justification justification14 = new Justification() { Val = JustificationValues.Left };

            styleParagraphProperties19.Append(tabs4);
            styleParagraphProperties19.Append(spacingBetweenLines26);
            styleParagraphProperties19.Append(indentation6);
            styleParagraphProperties19.Append(justification14);

            StyleRunProperties styleRunProperties30 = new StyleRunProperties();
            RunFonts runFonts40 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            Color color46 = new Color() { Val = "auto" };
            FontSize fontSize57 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "22" };
            Languages languages15 = new Languages() { EastAsia = "en-US" };

            styleRunProperties30.Append(runFonts40);
            styleRunProperties30.Append(color46);
            styleRunProperties30.Append(fontSize57);
            styleRunProperties30.Append(fontSizeComplexScript43);
            styleRunProperties30.Append(languages15);

            style34.Append(styleName34);
            style34.Append(basedOn18);
            style34.Append(linkedStyle22);
            style34.Append(uIPriority29);
            style34.Append(unhideWhenUsed13);
            style34.Append(rsid23);
            style34.Append(styleParagraphProperties19);
            style34.Append(styleRunProperties30);

            Style style35 = new Style() { Type = StyleValues.Character, StyleId = "RodapChar", CustomStyle = true };
            StyleName styleName35 = new StyleName() { Val = "Rodapé Char" };
            BasedOn basedOn19 = new BasedOn() { Val = "Fontepargpadro" };
            LinkedStyle linkedStyle23 = new LinkedStyle() { Val = "Rodap" };
            UIPriority uIPriority30 = new UIPriority() { Val = 99 };
            Rsid rsid24 = new Rsid() { Val = "006054CD" };

            StyleRunProperties styleRunProperties31 = new StyleRunProperties();
            RunFonts runFonts41 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            FontSize fontSize58 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "22" };
            Languages languages16 = new Languages() { EastAsia = "en-US" };

            styleRunProperties31.Append(runFonts41);
            styleRunProperties31.Append(fontSize58);
            styleRunProperties31.Append(fontSizeComplexScript44);
            styleRunProperties31.Append(languages16);

            style35.Append(styleName35);
            style35.Append(basedOn19);
            style35.Append(linkedStyle23);
            style35.Append(uIPriority30);
            style35.Append(rsid24);
            style35.Append(styleRunProperties31);

            Style style36 = new Style() { Type = StyleValues.Paragraph, StyleId = "Sumrio1" };
            StyleName styleName36 = new StyleName() { Val = "toc 1" };
            BasedOn basedOn20 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle13 = new NextParagraphStyle() { Val = "Normal" };
            AutoRedefine autoRedefine1 = new AutoRedefine();
            UIPriority uIPriority31 = new UIPriority() { Val = 39 };
            UnhideWhenUsed unhideWhenUsed14 = new UnhideWhenUsed();
            Rsid rsid25 = new Rsid() { Val = "006054CD" };

            StyleParagraphProperties styleParagraphProperties20 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines27 = new SpacingBetweenLines() { Line = "259", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation7 = new Indentation() { FirstLine = "0" };
            Justification justification15 = new Justification() { Val = JustificationValues.Left };

            styleParagraphProperties20.Append(spacingBetweenLines27);
            styleParagraphProperties20.Append(indentation7);
            styleParagraphProperties20.Append(justification15);

            StyleRunProperties styleRunProperties32 = new StyleRunProperties();
            RunFonts runFonts42 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            Color color47 = new Color() { Val = "auto" };
            FontSize fontSize59 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript45 = new FontSizeComplexScript() { Val = "22" };
            Languages languages17 = new Languages() { EastAsia = "en-US" };

            styleRunProperties32.Append(runFonts42);
            styleRunProperties32.Append(color47);
            styleRunProperties32.Append(fontSize59);
            styleRunProperties32.Append(fontSizeComplexScript45);
            styleRunProperties32.Append(languages17);

            style36.Append(styleName36);
            style36.Append(basedOn20);
            style36.Append(nextParagraphStyle13);
            style36.Append(autoRedefine1);
            style36.Append(uIPriority31);
            style36.Append(unhideWhenUsed14);
            style36.Append(rsid25);
            style36.Append(styleParagraphProperties20);
            style36.Append(styleRunProperties32);

            Style style37 = new Style() { Type = StyleValues.Character, StyleId = "Hyperlink" };
            StyleName styleName37 = new StyleName() { Val = "Hyperlink" };
            BasedOn basedOn21 = new BasedOn() { Val = "Fontepargpadro" };
            UIPriority uIPriority32 = new UIPriority() { Val = 99 };
            UnhideWhenUsed unhideWhenUsed15 = new UnhideWhenUsed();
            Rsid rsid26 = new Rsid() { Val = "006054CD" };

            StyleRunProperties styleRunProperties33 = new StyleRunProperties();
            Color color48 = new Color() { Val = "0563C1", ThemeColor = ThemeColorValues.Hyperlink };
            Underline underline1 = new Underline() { Val = UnderlineValues.Single };

            styleRunProperties33.Append(color48);
            styleRunProperties33.Append(underline1);

            style37.Append(styleName37);
            style37.Append(basedOn21);
            style37.Append(uIPriority32);
            style37.Append(unhideWhenUsed15);
            style37.Append(rsid26);
            style37.Append(styleRunProperties33);

            Style style38 = new Style() { Type = StyleValues.Paragraph, StyleId = "Sumrio2" };
            StyleName styleName38 = new StyleName() { Val = "toc 2" };
            BasedOn basedOn22 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle14 = new NextParagraphStyle() { Val = "Normal" };
            AutoRedefine autoRedefine2 = new AutoRedefine();
            UIPriority uIPriority33 = new UIPriority() { Val = 39 };
            UnhideWhenUsed unhideWhenUsed16 = new UnhideWhenUsed();
            Rsid rsid27 = new Rsid() { Val = "006054CD" };

            StyleParagraphProperties styleParagraphProperties21 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines28 = new SpacingBetweenLines() { Line = "259", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation8 = new Indentation() { Start = "220", FirstLine = "0" };
            Justification justification16 = new Justification() { Val = JustificationValues.Left };

            styleParagraphProperties21.Append(spacingBetweenLines28);
            styleParagraphProperties21.Append(indentation8);
            styleParagraphProperties21.Append(justification16);

            StyleRunProperties styleRunProperties34 = new StyleRunProperties();
            RunFonts runFonts43 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            Color color49 = new Color() { Val = "auto" };
            FontSize fontSize60 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript46 = new FontSizeComplexScript() { Val = "22" };
            Languages languages18 = new Languages() { EastAsia = "en-US" };

            styleRunProperties34.Append(runFonts43);
            styleRunProperties34.Append(color49);
            styleRunProperties34.Append(fontSize60);
            styleRunProperties34.Append(fontSizeComplexScript46);
            styleRunProperties34.Append(languages18);

            style38.Append(styleName38);
            style38.Append(basedOn22);
            style38.Append(nextParagraphStyle14);
            style38.Append(autoRedefine2);
            style38.Append(uIPriority33);
            style38.Append(unhideWhenUsed16);
            style38.Append(rsid27);
            style38.Append(styleParagraphProperties21);
            style38.Append(styleRunProperties34);

            Style style39 = new Style() { Type = StyleValues.Paragraph, StyleId = "Sumrio3" };
            StyleName styleName39 = new StyleName() { Val = "toc 3" };
            BasedOn basedOn23 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle15 = new NextParagraphStyle() { Val = "Normal" };
            AutoRedefine autoRedefine3 = new AutoRedefine();
            UIPriority uIPriority34 = new UIPriority() { Val = 39 };
            UnhideWhenUsed unhideWhenUsed17 = new UnhideWhenUsed();
            Rsid rsid28 = new Rsid() { Val = "006054CD" };

            StyleParagraphProperties styleParagraphProperties22 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines29 = new SpacingBetweenLines() { Line = "259", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation9 = new Indentation() { Start = "440", FirstLine = "0" };
            Justification justification17 = new Justification() { Val = JustificationValues.Left };

            styleParagraphProperties22.Append(spacingBetweenLines29);
            styleParagraphProperties22.Append(indentation9);
            styleParagraphProperties22.Append(justification17);

            StyleRunProperties styleRunProperties35 = new StyleRunProperties();
            RunFonts runFonts44 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            Color color50 = new Color() { Val = "auto" };
            FontSize fontSize61 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript47 = new FontSizeComplexScript() { Val = "22" };
            Languages languages19 = new Languages() { EastAsia = "en-US" };

            styleRunProperties35.Append(runFonts44);
            styleRunProperties35.Append(color50);
            styleRunProperties35.Append(fontSize61);
            styleRunProperties35.Append(fontSizeComplexScript47);
            styleRunProperties35.Append(languages19);

            style39.Append(styleName39);
            style39.Append(basedOn23);
            style39.Append(nextParagraphStyle15);
            style39.Append(autoRedefine3);
            style39.Append(uIPriority34);
            style39.Append(unhideWhenUsed17);
            style39.Append(rsid28);
            style39.Append(styleParagraphProperties22);
            style39.Append(styleRunProperties35);

            Style style40 = new Style() { Type = StyleValues.Paragraph, StyleId = "Textodebalo" };
            StyleName styleName40 = new StyleName() { Val = "Balloon Text" };
            BasedOn basedOn24 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle24 = new LinkedStyle() { Val = "TextodebaloChar" };
            UIPriority uIPriority35 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden10 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed18 = new UnhideWhenUsed();
            Rsid rsid29 = new Rsid() { Val = "00033F32" };

            StyleParagraphProperties styleParagraphProperties23 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines30 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties23.Append(spacingBetweenLines30);

            StyleRunProperties styleRunProperties36 = new StyleRunProperties();
            RunFonts runFonts45 = new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI", ComplexScript = "Segoe UI" };
            FontSize fontSize62 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript48 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties36.Append(runFonts45);
            styleRunProperties36.Append(fontSize62);
            styleRunProperties36.Append(fontSizeComplexScript48);

            style40.Append(styleName40);
            style40.Append(basedOn24);
            style40.Append(linkedStyle24);
            style40.Append(uIPriority35);
            style40.Append(semiHidden10);
            style40.Append(unhideWhenUsed18);
            style40.Append(rsid29);
            style40.Append(styleParagraphProperties23);
            style40.Append(styleRunProperties36);

            Style style41 = new Style() { Type = StyleValues.Character, StyleId = "Ttulo2Char1", CustomStyle = true };
            StyleName styleName41 = new StyleName() { Val = "Título 2 Char1" };
            BasedOn basedOn25 = new BasedOn() { Val = "Fontepargpadro" };
            LinkedStyle linkedStyle25 = new LinkedStyle() { Val = "Ttulo2" };
            UIPriority uIPriority36 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden11 = new SemiHidden();
            Rsid rsid30 = new Rsid() { Val = "00033F32" };

            StyleRunProperties styleRunProperties37 = new StyleRunProperties();
            RunFonts runFonts46 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color51 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize63 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript49 = new FontSizeComplexScript() { Val = "26" };

            styleRunProperties37.Append(runFonts46);
            styleRunProperties37.Append(color51);
            styleRunProperties37.Append(fontSize63);
            styleRunProperties37.Append(fontSizeComplexScript49);

            style41.Append(styleName41);
            style41.Append(basedOn25);
            style41.Append(linkedStyle25);
            style41.Append(uIPriority36);
            style41.Append(semiHidden11);
            style41.Append(rsid30);
            style41.Append(styleRunProperties37);

            Style style42 = new Style() { Type = StyleValues.Character, StyleId = "Ttulo3Char1", CustomStyle = true };
            StyleName styleName42 = new StyleName() { Val = "Título 3 Char1" };
            BasedOn basedOn26 = new BasedOn() { Val = "Fontepargpadro" };
            LinkedStyle linkedStyle26 = new LinkedStyle() { Val = "Ttulo3" };
            UIPriority uIPriority37 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden12 = new SemiHidden();
            Rsid rsid31 = new Rsid() { Val = "00033F32" };

            StyleRunProperties styleRunProperties38 = new StyleRunProperties();
            RunFonts runFonts47 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color52 = new Color() { Val = "1F3763", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "7F" };
            FontSize fontSize64 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript50 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties38.Append(runFonts47);
            styleRunProperties38.Append(color52);
            styleRunProperties38.Append(fontSize64);
            styleRunProperties38.Append(fontSizeComplexScript50);

            style42.Append(styleName42);
            style42.Append(basedOn26);
            style42.Append(linkedStyle26);
            style42.Append(uIPriority37);
            style42.Append(semiHidden12);
            style42.Append(rsid31);
            style42.Append(styleRunProperties38);

            Style style43 = new Style() { Type = StyleValues.Character, StyleId = "Ttulo4Char1", CustomStyle = true };
            StyleName styleName43 = new StyleName() { Val = "Título 4 Char1" };
            BasedOn basedOn27 = new BasedOn() { Val = "Fontepargpadro" };
            LinkedStyle linkedStyle27 = new LinkedStyle() { Val = "Ttulo4" };
            UIPriority uIPriority38 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden13 = new SemiHidden();
            Rsid rsid32 = new Rsid() { Val = "00033F32" };

            StyleRunProperties styleRunProperties39 = new StyleRunProperties();
            RunFonts runFonts48 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Italic italic10 = new Italic();
            ItalicComplexScript italicComplexScript8 = new ItalicComplexScript();
            Color color53 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };

            styleRunProperties39.Append(runFonts48);
            styleRunProperties39.Append(italic10);
            styleRunProperties39.Append(italicComplexScript8);
            styleRunProperties39.Append(color53);

            style43.Append(styleName43);
            style43.Append(basedOn27);
            style43.Append(linkedStyle27);
            style43.Append(uIPriority38);
            style43.Append(semiHidden13);
            style43.Append(rsid32);
            style43.Append(styleRunProperties39);

            Style style44 = new Style() { Type = StyleValues.Character, StyleId = "TextodebaloChar", CustomStyle = true };
            StyleName styleName44 = new StyleName() { Val = "Texto de balão Char" };
            BasedOn basedOn28 = new BasedOn() { Val = "Fontepargpadro" };
            LinkedStyle linkedStyle28 = new LinkedStyle() { Val = "Textodebalo" };
            UIPriority uIPriority39 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden14 = new SemiHidden();
            Rsid rsid33 = new Rsid() { Val = "00033F32" };

            StyleRunProperties styleRunProperties40 = new StyleRunProperties();
            RunFonts runFonts49 = new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI", ComplexScript = "Segoe UI" };
            Color color54 = new Color() { Val = "000000" };
            FontSize fontSize65 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript51 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties40.Append(runFonts49);
            styleRunProperties40.Append(color54);
            styleRunProperties40.Append(fontSize65);
            styleRunProperties40.Append(fontSizeComplexScript51);

            style44.Append(styleName44);
            style44.Append(basedOn28);
            style44.Append(linkedStyle28);
            style44.Append(uIPriority39);
            style44.Append(semiHidden14);
            style44.Append(rsid33);
            style44.Append(styleRunProperties40);

            Style style45 = new Style() { Type = StyleValues.Paragraph, StyleId = "Cabealho" };
            StyleName styleName45 = new StyleName() { Val = "header" };
            BasedOn basedOn29 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle29 = new LinkedStyle() { Val = "CabealhoChar" };
            UIPriority uIPriority40 = new UIPriority() { Val = 99 };
            UnhideWhenUsed unhideWhenUsed19 = new UnhideWhenUsed();
            Rsid rsid34 = new Rsid() { Val = "003D4FFA" };

            StyleParagraphProperties styleParagraphProperties24 = new StyleParagraphProperties();

            Tabs tabs5 = new Tabs();
            TabStop tabStop9 = new TabStop() { Val = TabStopValues.Center, Position = 4252 };
            TabStop tabStop10 = new TabStop() { Val = TabStopValues.Right, Position = 8504 };

            tabs5.Append(tabStop9);
            tabs5.Append(tabStop10);
            SpacingBetweenLines spacingBetweenLines31 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties24.Append(tabs5);
            styleParagraphProperties24.Append(spacingBetweenLines31);

            style45.Append(styleName45);
            style45.Append(basedOn29);
            style45.Append(linkedStyle29);
            style45.Append(uIPriority40);
            style45.Append(unhideWhenUsed19);
            style45.Append(rsid34);
            style45.Append(styleParagraphProperties24);

            Style style46 = new Style() { Type = StyleValues.Character, StyleId = "CabealhoChar", CustomStyle = true };
            StyleName styleName46 = new StyleName() { Val = "Cabeçalho Char" };
            BasedOn basedOn30 = new BasedOn() { Val = "Fontepargpadro" };
            LinkedStyle linkedStyle30 = new LinkedStyle() { Val = "Cabealho" };
            UIPriority uIPriority41 = new UIPriority() { Val = 99 };
            Rsid rsid35 = new Rsid() { Val = "003D4FFA" };

            StyleRunProperties styleRunProperties41 = new StyleRunProperties();
            RunFonts runFonts50 = new RunFonts() { HighAnsi = "Calibri Light", ComplexScript = "Calibri Light" };
            Color color55 = new Color() { Val = "000000" };

            styleRunProperties41.Append(runFonts50);
            styleRunProperties41.Append(color55);

            style46.Append(styleName46);
            style46.Append(basedOn30);
            style46.Append(linkedStyle30);
            style46.Append(uIPriority41);
            style46.Append(rsid35);
            style46.Append(styleRunProperties41);

            Style style47 = new Style() { Type = StyleValues.Character, StyleId = "TextodoEspaoReservado" };
            StyleName styleName47 = new StyleName() { Val = "Placeholder Text" };
            BasedOn basedOn31 = new BasedOn() { Val = "Fontepargpadro" };
            UIPriority uIPriority42 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden15 = new SemiHidden();
            Rsid rsid36 = new Rsid() { Val = "003D4FFA" };

            StyleRunProperties styleRunProperties42 = new StyleRunProperties();
            Color color56 = new Color() { Val = "808080" };

            styleRunProperties42.Append(color56);

            style47.Append(styleName47);
            style47.Append(basedOn31);
            style47.Append(uIPriority42);
            style47.Append(semiHidden15);
            style47.Append(rsid36);
            style47.Append(styleRunProperties42);

            styles1.Append(docDefaults1);
            styles1.Append(latentStyles1);
            styles1.Append(style1);
            styles1.Append(style2);
            styles1.Append(style3);
            styles1.Append(style4);
            styles1.Append(style5);
            styles1.Append(style6);
            styles1.Append(style7);
            styles1.Append(style8);
            styles1.Append(style9);
            styles1.Append(style10);
            styles1.Append(style11);
            styles1.Append(style12);
            styles1.Append(style13);
            styles1.Append(style14);
            styles1.Append(style15);
            styles1.Append(style16);
            styles1.Append(style17);
            styles1.Append(style18);
            styles1.Append(style19);
            styles1.Append(style20);
            styles1.Append(style21);
            styles1.Append(style22);
            styles1.Append(style23);
            styles1.Append(style24);
            styles1.Append(style25);
            styles1.Append(style26);
            styles1.Append(style27);
            styles1.Append(style28);
            styles1.Append(style29);
            styles1.Append(style30);
            styles1.Append(style31);
            styles1.Append(style32);
            styles1.Append(style33);
            styles1.Append(style34);
            styles1.Append(style35);
            styles1.Append(style36);
            styles1.Append(style37);
            styles1.Append(style38);
            styles1.Append(style39);
            styles1.Append(style40);
            styles1.Append(style41);
            styles1.Append(style42);
            styles1.Append(style43);
            styles1.Append(style44);
            styles1.Append(style45);
            styles1.Append(style46);
            styles1.Append(style47);

            return styles1;
        }



        // Add a StylesDefinitionsPart to the document. Returns a reference to it.
        public static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument doc, Styles root = null)
        {
            StyleDefinitionsPart part;
            part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            if (root == null)
                root = new Styles();
            root.Save(part);
            return part;
        }
        #endregion

        #region Numbering
        // Extract the numbering part from a 
        // word processing document as an XDocument instance.
        public static XDocument ExtractNumberingPart(string fileName)
        {
            // Declare a variable to hold the XDocument.
            XDocument numbering = null;

            WordprocessingDocument document = null;

            // Open the document for read access and get a reference.
            try
            {
                document = WordprocessingDocument.Open(fileName, false);

                // Get a reference to the main document part.
                var docPart = document.MainDocumentPart;

                // Assign a reference to the appropriate part to the stylesPart variable.
                NumberingDefinitionsPart numberingPart = null;
                numberingPart = docPart.NumberingDefinitionsPart;

                // If the part exists, read it into the XDocument.
                if (numberingPart != null)
                {
                    using (var reader = XmlNodeReader.Create(
                        numberingPart.GetStream(FileMode.Open, FileAccess.Read)))
                    {
                        // Create the XDocument.
                        numbering = XDocument.Load(reader);
                    }
                }
            }
            finally
            {
                if (document != null)
                    document.Close();
            }

            // Return the XDocument instance.
            return numbering;
        }

        // Add a NumberingDefinitionsPart to the document. Returns a reference to it.
        public static NumberingDefinitionsPart AddNumberingPartToPackage(WordprocessingDocument doc, Numbering root = null)
        {
            NumberingDefinitionsPart part = doc.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();

            root.Save(part);

            return part;
        }

        private static AbstractNum CreateAbstractNum(int id, PropriedadesNumeracao[] pNum)
        {
            AbstractNum abstractNum = new AbstractNum() { AbstractNumberId = id };
            abstractNum.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid = new Nsid() { Val = NSID };
            MultiLevelType multiLevelType = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode = new TemplateCode() { Val = NSID };

            abstractNum.Append(nsid);
            abstractNum.Append(multiLevelType);
            abstractNum.Append(templateCode);

            for (int L = 0; L < pNum.Length; L++)
            {
                // Level level = pNum[L].level;
                Level level = new Level() { LevelIndex = L };
                StartNumberingValue startNumberingValue = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat = new NumberingFormat() { Val = pNum[L].numberFormatValue };
                LevelText levelText = new LevelText() { Val = pNum[L].levelTextValue };
                LevelJustification levelJustification = new LevelJustification() { Val = pNum[L].levelJustificationValue };

                PreviousParagraphProperties previousParagraphProperties = new PreviousParagraphProperties();

                if (pNum[L].tabStopPositionValue > 0)
                {
                    Tabs tabs = new Tabs();
                    TabStop tabStop = new TabStop() { Val = TabStopValues.Number, Position = pNum[L].tabStopPositionValue };

                    tabs.Append(tabStop);

                    previousParagraphProperties.Append(tabs);
                }

                Indentation indentation = new Indentation() { Start = pNum[L].indentationStartValue.ToString(), Hanging = pNum[L].indentationHangingValue.ToString() };
                previousParagraphProperties.Append(indentation);

                NumberingSymbolRunProperties numberingSymbolRunProperties = new NumberingSymbolRunProperties();
                numberingSymbolRunProperties.Append(pNum[L].runFonts);

                if (pNum[L].fontSize > 0)
                {
                    FontSize fontSize = new FontSize() { Val = pNum[L].fontSize.ToString() };
                    numberingSymbolRunProperties.Append(fontSize);
                }

                level.Append(startNumberingValue);
                level.Append(numberingFormat);
                level.Append(levelText);
                level.Append(levelJustification);
                level.Append(previousParagraphProperties);
                level.Append(numberingSymbolRunProperties);

                abstractNum.Append(level);
            }

            return abstractNum;
        }

        public static int CreateNumberingInstance(WordprocessingDocument doc, int abstractNumId)
        {
            Numbering numbering = doc.MainDocumentPart.NumberingDefinitionsPart.Numbering;

            var numberId = numbering.Elements<NumberingInstance>().Count() + 1;

            NumberingInstance numberingInstance = new NumberingInstance() { NumberID = numberId };
            AbstractNumId absNumId = new AbstractNumId() { Val = abstractNumId };

            LevelOverride levelOverride = new LevelOverride() { LevelIndex = 0 };
            levelOverride.StartOverrideNumberingValue = new StartOverrideNumberingValue() { Val = 1 };

            numberingInstance.Append(levelOverride);

            numberingInstance.Append(absNumId);
            numbering.Append(numberingInstance);

            return numberId;
        }

        // Generates content of numberingDefinitionsPart.
        public static Numbering GenerateNumberingDefinitionsPartContent()
        {
            Numbering numbering = new Numbering() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se wp14" } };
            numbering.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            numbering.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            numbering.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            numbering.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            numbering.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            numbering.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            numbering.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            numbering.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            numbering.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            numbering.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            numbering.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            numbering.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            numbering.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            numbering.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            numbering.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            numbering.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            numbering.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            numbering.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            numbering.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            numbering.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            #region AbstractNum1
            PropriedadesNumeracao[] absnum1 = new PropriedadesNumeracao[]
            {
                // MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 1080,
                    indentationStartValue = 1080,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" }
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 1800,
                    indentationStartValue = 1800,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New" }
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "§",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 2520,
                    indentationStartValue = 2520,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" }
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 3240,
                    indentationStartValue = 3240,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" }
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 3960,
                    indentationStartValue = 3960,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New" }
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "§",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 4680,
                    indentationStartValue = 4680,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" }
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 5400,
                    indentationStartValue = 5400,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" }
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 6120,
                    indentationStartValue = 6120,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New" }
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "§",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 6840,
                    indentationStartValue = 6840,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" }
                },
            };

            AbstractNum abstractNum1 = CreateAbstractNum(0, absnum1);
            #endregion

            #region AbstractNum2
            PropriedadesNumeracao[] absnum2 = new PropriedadesNumeracao[]
            {
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Decimal,
                    levelTextValue = "%1)",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 1080,
                    indentationStartValue = 1080,
                    indentationHangingValue = 720,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.LowerLetter,
                    levelTextValue = "%2.",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 1440,
                    indentationStartValue = 1440,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.LowerRoman,
                    levelTextValue = "%3.",
                    levelJustificationValue = LevelJustificationValues.Right,
                    tabStopPositionValue = 2160,
                    indentationStartValue = 2160,
                    indentationHangingValue = 180,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Decimal,
                    levelTextValue = "%4.",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 2880,
                    indentationStartValue = 2880,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.LowerLetter,
                    levelTextValue = "%5.",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 3600,
                    indentationStartValue = 3600,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.LowerRoman,
                    levelTextValue = "%6.",
                    levelJustificationValue = LevelJustificationValues.Right,
                    tabStopPositionValue = 4320,
                    indentationStartValue = 4320,
                    indentationHangingValue = 180,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Decimal,
                    levelTextValue = "%7.",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 5040,
                    indentationStartValue = 5040,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.LowerLetter,
                    levelTextValue = "%8.",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 5760,
                    indentationStartValue = 5760,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.LowerRoman,
                    levelTextValue = "%9.",
                    levelJustificationValue = LevelJustificationValues.Right,
                    tabStopPositionValue = 6480,
                    indentationStartValue = 6480,
                    indentationHangingValue = 180,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
            };

            AbstractNum abstractNum2 = CreateAbstractNum(0, absnum2);
            #endregion

            #region AbstractNum3
            PropriedadesNumeracao[] absnum3 = new PropriedadesNumeracao[]
            {
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Decimal,
                    levelTextValue = "%1)",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 1080,
                    indentationStartValue = 1080,
                    indentationHangingValue = 720,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.LowerLetter,
                    levelTextValue = "%2.",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 1440,
                    indentationStartValue = 1440,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.LowerRoman,
                    levelTextValue = "%3.",
                    levelJustificationValue = LevelJustificationValues.Right,
                    tabStopPositionValue = 2160,
                    indentationStartValue = 2160,
                    indentationHangingValue = 180,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Decimal,
                    levelTextValue = "%4.",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 2880,
                    indentationStartValue = 2880,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.LowerLetter,
                    levelTextValue = "%5.",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 3600,
                    indentationStartValue = 3600,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.LowerRoman,
                    levelTextValue = "%6.",
                    levelJustificationValue = LevelJustificationValues.Right,
                    tabStopPositionValue = 4320,
                    indentationStartValue = 4320,
                    indentationHangingValue = 180,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Decimal,
                    levelTextValue = "%7.",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 5040,
                    indentationStartValue = 5040,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.LowerLetter,
                    levelTextValue = "%8.",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 5760,
                    indentationStartValue = 5760,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.LowerRoman,
                    levelTextValue = "%9.",
                    levelJustificationValue = LevelJustificationValues.Right,
                    tabStopPositionValue = 6480,
                    indentationStartValue = 6480,
                    indentationHangingValue = 180,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
            };

            AbstractNum abstractNum3 = CreateAbstractNum(2, absnum3);
            #endregion

            #region AbstractNum4
            PropriedadesNumeracao[] absnum4 = new PropriedadesNumeracao[]
            {
                // MultiLevelType multiLevelType4 = new MultiLevelType() { Val = MultiLevelValues.Multilevel };
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Decimal,
                    levelTextValue = "%1",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 720,
                    indentationStartValue = 400,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Decimal,
                    levelTextValue = "%1.%2",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 1440,
                    indentationStartValue = 400,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Decimal,
                    levelTextValue = "%1.%2.%3",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 2160,
                    indentationStartValue = 400,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Decimal,
                    levelTextValue = "%1.%2.%3.%4",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 2880,
                    indentationStartValue = 400,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Decimal,
                    levelTextValue = "%1.%2.%3.%4.%5",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 3600,
                    indentationStartValue = 400,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Decimal,
                    levelTextValue = "%1.%2.%3.%4.%5.%6",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 4320,
                    indentationStartValue = 400,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Decimal,
                    levelTextValue = "%1.%2.%3.%4.%5.%6.%7",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 5040,
                    indentationStartValue = 400,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Decimal,
                    levelTextValue = "%1.%2.%3.%4.%5.%6.%7.%8",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 5760,
                    indentationStartValue = 400,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Decimal,
                    levelTextValue = "%1.%2.%3.%4.%5.%6.%7.%8.%9",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 6480,
                    indentationStartValue = 400,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { ComplexScript = "Times New Roman" },
                },
            };

            AbstractNum abstractNum4 = CreateAbstractNum(3, absnum4);
            #endregion

            #region AbstractNum5
            PropriedadesNumeracao[] absnum5 = new PropriedadesNumeracao[]
            {
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "-",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 1080,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "Times New Roman" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 1800,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "§",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 2520,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 3240,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 3960,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "§",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 4680,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 5400,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 6120,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "§",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 6840,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" },
                },
            };

            AbstractNum abstractNum5 = CreateAbstractNum(4, absnum5);
            #endregion

            #region AbstractNum6
            PropriedadesNumeracao[] absnum6 = new PropriedadesNumeracao[]
            {
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 720,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 1440,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "§",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 2160,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 2880,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 3600,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "§",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 4320,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 5040,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 5760,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "§",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 6480,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" },
                },
            };

            AbstractNum abstractNum6 = CreateAbstractNum(5, absnum6);
            #endregion

            #region AbstractNum7
            PropriedadesNumeracao[] absnum7 = new PropriedadesNumeracao[]
            {
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 720,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 1440,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "§",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 2160,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 2880,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 3600,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "§",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 4320,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 5040,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 5760,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "§",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 6480,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" },
                },
            };

            AbstractNum abstractNum7 = CreateAbstractNum(6, absnum7);
            #endregion

            #region AbstractNum8
            PropriedadesNumeracao[] absnum8 = new PropriedadesNumeracao[]
            {
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 360,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 1080,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "§",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 1800,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 2520,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 3240,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "§",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 3960,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 4680,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 5400,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "§",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 6120,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" },
                },
            };

            AbstractNum abstractNum8 = CreateAbstractNum(7, absnum8);
            #endregion

            #region AbstractNum9
            PropriedadesNumeracao[] absnum9 = new PropriedadesNumeracao[]
            {
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 630,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 1440,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "§",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 2160,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 2880,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 3600,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "§",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 4320,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 5040,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "o",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 5760,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" },
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "§",
                    levelJustificationValue = LevelJustificationValues.Left,
                    indentationStartValue = 6480,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" },
                },
            };

            AbstractNum abstractNum9 = CreateAbstractNum(8, absnum9);
            #endregion

            #region AbstractNum10
            PropriedadesNumeracao[] absnum10 = new PropriedadesNumeracao[]
            {
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 720,
                    indentationStartValue = 720,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                    fontSize = 20,
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 1440,
                    indentationStartValue = 1440,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                    fontSize = 20,
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 2160,
                    indentationStartValue = 2160,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                    fontSize = 20,
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 2880,
                    indentationStartValue = 2880,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                    fontSize = 20,
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 3600,
                    indentationStartValue = 3600,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                    fontSize = 20,
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 4320,
                    indentationStartValue = 4320,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                    fontSize = 20,
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 5040,
                    indentationStartValue = 5040,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                    fontSize = 20,
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 5760,
                    indentationStartValue = 5760,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                    fontSize = 20,
                },
                new PropriedadesNumeracao{
                    numberFormatValue = NumberFormatValues.Bullet,
                    levelTextValue = "·",
                    levelJustificationValue = LevelJustificationValues.Left,
                    tabStopPositionValue = 6480,
                    indentationStartValue = 6480,
                    indentationHangingValue = 360,
                    runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" },
                    fontSize = 20,
                },
           };

            AbstractNum abstractNum10 = CreateAbstractNum(9, absnum10);
            #endregion

            numbering.Append(abstractNum1);
            numbering.Append(abstractNum2);
            numbering.Append(abstractNum3);
            numbering.Append(abstractNum4);
            numbering.Append(abstractNum5);
            numbering.Append(abstractNum6);
            numbering.Append(abstractNum7);
            numbering.Append(abstractNum8);
            numbering.Append(abstractNum9);
            numbering.Append(abstractNum10);

            NumberingInstance numberingInstance1 = new NumberingInstance() { NumberID = 1 };
            AbstractNumId abstractNumId1 = new AbstractNumId() { Val = 3 };
            // 1. 1.
            numberingInstance1.Append(abstractNumId1);

            NumberingInstance numberingInstance2 = new NumberingInstance() { NumberID = 2 };
            AbstractNumId abstractNumId2 = new AbstractNumId() { Val = 9 };
            // bullet o
            numberingInstance2.Append(abstractNumId2);

            NumberingInstance numberingInstance3 = new NumberingInstance() { NumberID = 3 };
            AbstractNumId abstractNumId3 = new AbstractNumId() { Val = 2 };
            // 1) a.
            numberingInstance3.Append(abstractNumId3);

            NumberingInstance numberingInstance4 = new NumberingInstance() { NumberID = 4 };
            AbstractNumId abstractNumId4 = new AbstractNumId() { Val = 1 };
            // 1) a.
            numberingInstance4.Append(abstractNumId4);

            NumberingInstance numberingInstance5 = new NumberingInstance() { NumberID = 5 };
            AbstractNumId abstractNumId5 = new AbstractNumId() { Val = 4 };
            // - o
            numberingInstance5.Append(abstractNumId5);

            NumberingInstance numberingInstance6 = new NumberingInstance() { NumberID = 6 };
            AbstractNumId abstractNumId6 = new AbstractNumId() { Val = 0 };
            // bullet o
            numberingInstance6.Append(abstractNumId6);

            NumberingInstance numberingInstance7 = new NumberingInstance() { NumberID = 7 };
            AbstractNumId abstractNumId7 = new AbstractNumId() { Val = 6 };
            // bullet o
            numberingInstance7.Append(abstractNumId7);

            NumberingInstance numberingInstance8 = new NumberingInstance() { NumberID = 8 };
            AbstractNumId abstractNumId8 = new AbstractNumId() { Val = 8 };
            // bullet o
            numberingInstance8.Append(abstractNumId8);

            NumberingInstance numberingInstance9 = new NumberingInstance() { NumberID = 9 };
            AbstractNumId abstractNumId9 = new AbstractNumId() { Val = 5 };
            // bullet o
            numberingInstance9.Append(abstractNumId9);

            NumberingInstance numberingInstance10 = new NumberingInstance() { NumberID = 10 };
            AbstractNumId abstractNumId10 = new AbstractNumId() { Val = 7 };
            // o o 
            numberingInstance10.Append(abstractNumId10);

            numbering.Append(numberingInstance1);
            numbering.Append(numberingInstance2);
            numbering.Append(numberingInstance3);
            numbering.Append(numberingInstance4);
            numbering.Append(numberingInstance5);
            numbering.Append(numberingInstance6);
            numbering.Append(numberingInstance7);
            numbering.Append(numberingInstance8);
            numbering.Append(numberingInstance9);
            numbering.Append(numberingInstance10);

            // numberingDefinitionsPart1.Numbering = numbering;
            return numbering;
        }
        #endregion

        #region Footnotes
        // Extract the numbering part from a 
        // word processing document as an XDocument instance.
        public static XDocument ExtractFootnotesPart(string fileName)
        {
            // Declare a variable to hold the XDocument.
            XDocument footnotes = null;

            WordprocessingDocument document = null;

            // Open the document for read access and get a reference.
            try
            {
                document = WordprocessingDocument.Open(fileName, false);

                // Get a reference to the main document part.
                var docPart = document.MainDocumentPart;

                // Assign a reference to the appropriate part to the stylesPart variable.
                FootnotesPart footnotesPart = null;
                footnotesPart = docPart.FootnotesPart;

                // If the part exists, read it into the XDocument.
                if (footnotesPart != null)
                {
                    using (var reader = XmlNodeReader.Create(
                      footnotesPart.GetStream(FileMode.Open, FileAccess.Read)))
                    {
                        // Create the XDocument.
                        footnotes = XDocument.Load(reader);
                    }
                }
            }
            finally
            {
                if (document != null)
                    document.Close();
            }

            // Return the XDocument instance.
            return footnotes;
        }

        // Add a FootnotesPart to the document. Returns a reference to it.
        public static FootnotesPart AddFootnotesPartToPackage(WordprocessingDocument doc, Footnotes root = null)
        {
            FootnotesPart part;
            part = doc.MainDocumentPart.AddNewPart<FootnotesPart>();
            if (root == null)
                root = new Footnotes();
            root.Save(part);
            return part;
        }
        #endregion

        #region Headers
        // Extract the numbering part from a 
        // word processing document as an XDocument instance.
        public static List<XDocument> ExtractHeaderPart(string fileName)
        {
            // Declare a variable to hold the XDocument.
            List<XDocument> header = new List<XDocument>();

            WordprocessingDocument document = null;

            // Open the document for read access and get a reference.
            try
            {
                document = WordprocessingDocument.Open(fileName, false);

                // Get a reference to the main document part.
                var docPart = document.MainDocumentPart;

                // Assign a reference to the appropriate part to the stylesPart variable.
                IEnumerable<HeaderPart> headerPart = docPart.HeaderParts;

                // If the part exists, read it into the XDocument.
                if (headerPart != null)
                {
                    foreach (HeaderPart part in headerPart)
                    {
                        using (var reader = XmlNodeReader.Create(
                          part.GetStream(FileMode.Open, FileAccess.Read)))
                        {
                            // Create the XDocument.
                            header.Add(XDocument.Load(reader));
                        }
                    }
                }
            }
            finally
            {
                if (document != null)
                    document.Close();
            }

            // Return the XDocument instance.
            return header;
        }

        // Add a FootnotesPart to the document. Returns a reference to it.
        public static HeaderPart AddHeaderPartToPackage(WordprocessingDocument doc, List<Header> root = null)
        {
            HeaderPart part;
            part = doc.MainDocumentPart.AddNewPart<HeaderPart>();
            if (root == null)
                root = new List<Header>();
            foreach(Header header in root)
                header.Save(part);
            return part;
        }
        #endregion

        public static string AddStyleDefinitionsPartToPackage(WordprocessingDocument doc)
        {
            string id = "";

            // Verify that the document contains a 
            // WordProcessingCommentsPart part; if not, add a new one.
            if (doc.MainDocumentPart.StyleDefinitionsPart != null)
            {
                Styles styles =
                    doc.MainDocumentPart.StyleDefinitionsPart.Styles;

                if (styles.HasChildren)
                {
                    id = styles.Descendants<Style>().Select(e => e.StyleId.Value).Max();
                }
            }
            else
            {
                // No WordprocessingCommentsPart part exists, so add one to the package.
                StyleDefinitionsPart styleDefinitionsPart =
                    doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                styleDefinitionsPart.Styles = new Styles();
            }

            return id;
        }

        public static string AddCommentPartToPackage(WordprocessingDocument doc)
        {
            string id = "0";

            // Verify that the document contains a 
            // WordProcessingCommentsPart part; if not, add a new one.
            if (doc.MainDocumentPart.WordprocessingCommentsPart != null)
            {
                Comments comments =
                    doc.MainDocumentPart.WordprocessingCommentsPart.Comments;

                if (comments.HasChildren)
                {
                    // Obtain an unused ID.
                    id = comments.Descendants<Comment>().Select(e => e.Id.Value).Max();
                }
            }
            else
            {
                // No WordprocessingCommentsPart part exists, so add one to the package.
                WordprocessingCommentsPart commentPart =
                    doc.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
                commentPart.Comments = new Comments();
            }

            return id;
        }

        public static string AddNumberingPartToPackage(WordprocessingDocument doc)
        {
            int id = 0;

            // Verify that the document contains a 
            // WordProcessingCommentsPart part; if not, add a new one.
            if (doc.MainDocumentPart.NumberingDefinitionsPart != null)
            {
                Numbering numbering =
                    doc.MainDocumentPart.NumberingDefinitionsPart.Numbering;

                if (numbering.HasChildren)
                {
                    id = numbering.Descendants<AbstractNum>().Select(e => e.AbstractNumberId.Value).Max();
                }
            }
            else
            {
                // No WordprocessingCommentsPart part exists, so add one to the package.
                NumberingDefinitionsPart numberingPart =
                    doc.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                numberingPart.Numbering = new Numbering();
            }

            return id.ToString();
        }
    }
}


/*
 */
