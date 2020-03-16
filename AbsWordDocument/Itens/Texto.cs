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
    public class Texto : Paragrafo
    {
        private readonly List<OpenXmlElement> _runList;

        private new readonly int _numberingLevel;
        private new readonly int _numberingId;
        private new readonly bool _numbering;

        public Texto(string style = "Normal")
            : base(style) 
        {
            _runList = new List<OpenXmlElement>();

            _numberingLevel = Paragrafo._numberingLevel;
            _numberingId = Paragrafo._numberingId;
            _numbering = Paragrafo._numbering;
        }

        public Texto Append(string text, RunProperties rPr = null)
        {
            Run run = new Run();

            if (rPr != null)
                run.AppendChild(rPr);

            run.AppendChild(new Text { Space = SpaceProcessingModeValues.Preserve, Text = text });

            _runList.Add(run);

            return this;
        }

        public void StartComment(string commentId)
        {
            Run run = new Run();

            AnnotationReferenceMark annotationReferenceMark =
                new AnnotationReferenceMark();
            run.Append(annotationReferenceMark);

            _runList.Add(run);

            CommentRangeStart commentRangeStart = new CommentRangeStart() { Id = commentId };
            _runList.Add(commentRangeStart);
        }

        public void EndComment(string commentId)
        {
            CommentRangeEnd commentRangeEnd = new CommentRangeEnd() { Id = commentId };
            _runList.Add(commentRangeEnd);

            Run commentRun = new Run();

            CommentReference commentReference = new CommentReference() { Id = commentId };
            commentRun.Append(commentReference);

            _runList.Add(commentRun);
        }

        public override void ToWordDocument(WordprocessingDocument wordDocument)
        {
            Paragraph paragraph = base.CreateParagraph();

            if (_numbering)
            {
                // Create items for paragraph properties
                var numberingLevelReference = new NumberingLevelReference() { Val = _numberingLevel };
                var numberingId = new NumberingId() { Val = _numberingId };
                // var c = new NumberingRestart() { Val = RestartNumberValues.EachPage };
                var numberingProperties = new NumberingProperties(numberingLevelReference, numberingId);

                // create paragraph properties
                var paragraphProperties = new ParagraphProperties(numberingProperties);

                paragraph.Append(paragraphProperties);
            }

            foreach (OpenXmlElement element in _runList)
                paragraph.AppendChild(element);

            wordDocument.MainDocumentPart.Document.Body.AppendChild(paragraph);
        }
    }
}
