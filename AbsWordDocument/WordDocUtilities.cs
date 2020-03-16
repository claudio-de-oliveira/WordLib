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
