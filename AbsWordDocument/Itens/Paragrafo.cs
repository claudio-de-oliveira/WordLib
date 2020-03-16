using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AbsWordDocument.Itens
{
    public abstract class Paragrafo : WordItem
    {
        private readonly string _style;

        protected static int _numberingLevel = 0;
        protected static int _numberingId = 0;
        protected static bool _numbering = false;

        protected Paragrafo(string style = "Normal")
        {
            _style = style;
        }

        public Paragraph CreateParagraph()
        {
            Paragraph paragraph = new Paragraph();

            // If the paragraph has no ParagraphProperties object, create one.
            if (paragraph.Elements<ParagraphProperties>().Count() == 0)
                paragraph.PrependChild(new ParagraphProperties());

            // Get a reference to the ParagraphProperties object.
            ParagraphProperties pPr = paragraph.ParagraphProperties;

            // If a ParagraphStyleId object doesn't exist, create one.
            if (pPr.ParagraphStyleId == null)
                pPr.ParagraphStyleId = new ParagraphStyleId();

            // Set the style of the paragraph.
            pPr.ParagraphStyleId.Val = _style;

            return paragraph;
        }

        public static void StartNumbering(WordprocessingDocument wordDocument, int abstractId)
        {
            _numberingId = WordDocUtilities.CreateNumberingInstance(wordDocument, abstractId);
            // _numberingId = numberingId;
            _numbering = true;
            _numberingLevel = 0;
        }

        public static void IncrementNumbering()
        {
            if (_numberingLevel < 10)
                _numberingLevel++;
        }

        public static void DecrementNumbering()
        {
            if (_numberingLevel > 0)
                _numberingLevel--;
        }

        public static void EndNumbering()
        {
            _numbering = false;
        }
    }
}
