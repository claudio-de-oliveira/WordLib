using System;
using System.Collections.Generic;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AbsWordDocument.Itens
{
    public class Comentario : Paragrafo
    {
        private readonly List<string> _runList;
        private readonly Comment _comment;
        public string Id { get; private set; }

        public Comentario(WordprocessingDocument wordDocument, string author, string initials, string style = "Normal")
            : base(style)
        {
            _runList = new List<string>();

            Id = WordDocUtilities.AddCommentPartToPackage(wordDocument);

            _comment =
                new Comment()
                {
                    Initials = initials,
                    Author = author,
                    Date = System.Xml.XmlConvert.ToDateTime(
                        DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                        System.Xml.XmlDateTimeSerializationMode.RoundtripKind
                        ),
                    Id = String.Format("{0}", Id)
                };
        }

        public Comentario Append(string text, RunProperties rPr = null)
        {
            Run run = new Run();

            if (rPr != null)
                run.AppendChild(rPr);

            run.AppendChild(new Text { Space = SpaceProcessingModeValues.Preserve, Text = text });

            _runList.Add(run.OuterXml);

            return this;
        }

        public override void ToWordDocument(WordprocessingDocument wordDocument)
        {
            Paragraph paragraph = base.CreateParagraph();

            foreach (string str in _runList)
                paragraph.Append(new Run(str));
            _comment.AppendChild(paragraph);

            wordDocument.MainDocumentPart.WordprocessingCommentsPart.Comments.AppendChild(_comment);
        }
    }
}

