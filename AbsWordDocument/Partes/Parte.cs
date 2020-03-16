using System.Collections.Generic;

using DocumentFormat.OpenXml.Packaging;

using AbsWordDocument.Itens;

namespace AbsWordDocument.Partes
{
    public abstract class Parte : WordItem
    {
        private readonly string _style;
        private readonly string _titulo;

        protected List<Parte> Children { get; set; }
        public List<Paragrafo> Paragrafos { get; private set; }

        protected Parte(string text, string style)
        {
            this._titulo = text;
            this._style = style;

            Children = new List<Parte>();
            Paragrafos = new List<Paragrafo>();

            // Termina qualquer lista pendente
            Paragrafo.EndNumbering();
        }

        public override void ToWordDocument(WordprocessingDocument wordDocument)
        {
            wordDocument.MainDocumentPart.Document.Body.AppendChild(
                WordDocUtilities.CreateParagraphWithStyle(_titulo, _style)
                );

            foreach (Paragrafo p in Paragrafos)
                p.ToWordDocument(wordDocument);
            foreach (Parte p in Children)
                p.ToWordDocument(wordDocument);
        }
    }
}
