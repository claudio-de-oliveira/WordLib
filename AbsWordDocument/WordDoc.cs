using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using System.Xml;
using System.Xml.Linq;
using AbsWordDocument.Itens;
using AbsWordDocument.Partes;

namespace AbsWordDocument
{
    public class WordDoc
    {
        public WordprocessingDocument WordDocument { get; set; }
        public MainDocumentPart MainPart { get; set; }
        public Body Body { get; set; }

        public List<Parte> Partes;
        public List<Comentario> Comentarios;

        private MemoryStream _stream;

        private readonly string _author;
        private readonly string _initials;

        protected static int _numberingLevel = 0;
        protected static int _numberingId = 0;
        protected static bool _numbering = false;

        public WordDoc(string author, string initials)
        {
            Partes = new List<Parte>();
            Comentarios = new List<Comentario>(); 

            _stream = new MemoryStream();

            WordDocument = WordprocessingDocument.Create(_stream, WordprocessingDocumentType.Document);

            // Add a main document part. 
            MainPart = WordDocument.AddMainDocumentPart();

            // Create the document structure and add some text.
            MainPart.Document = new Document();
            Body = MainPart.Document.AppendChild(new Body());

            _author = author;
            _initials = initials;

            // SectionProperties SecPro = new SectionProperties();
            // PageSize PSize = new PageSize();
            // PSize.Width = 15000;
            // PSize.Height = 11000;
            // SecPro.Append(PSize);
            // Body.Append(SecPro);

        }

        public Comentario CreateComment()
        {
            Comentario comentario = new Comentario(WordDocument, _author, _initials);

            Comentarios.Add(comentario);

            return comentario;
        }

        public int CreateAbstractNum(AbstractNum newAbstractNum)
        {
            if (WordDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.HasChildren)
            {
                AbstractNum lastAbstract = WordDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.Elements<AbstractNum>().Last();
                newAbstractNum.AbstractNumberId = lastAbstract.AbstractNumberId + 1;
                WordDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.InsertAfter(newAbstractNum, lastAbstract);
            }
            else
            {
                newAbstractNum.AbstractNumberId = 0;
                WordDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.AppendChild(newAbstractNum);
            }
            return newAbstractNum.AbstractNumberId;
        }

        public int CreateNumberingInstance(int absId)
        {
            NumberingInstance lastNumbering = WordDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.Elements<NumberingInstance>().Last();
            NumberingInstance newNumberingInstance = new NumberingInstance() { NumberID = lastNumbering.NumberID + 1 };
            newNumberingInstance.AbstractNumId = new AbstractNumId() { Val = absId };
            WordDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.InsertAfter(newNumberingInstance, lastNumbering);
            return lastNumbering.NumberID;
        }

        public void SetNumberingFromDocument(string stylesFile)
        {
            XDocument numbering = WordDocUtilities.ExtractNumberingPart(stylesFile);

            // Get the Styles part for this document.
            NumberingDefinitionsPart part =
                WordDocument.MainDocumentPart.NumberingDefinitionsPart;

            Numbering root = new Numbering(numbering.ToString());

            // If the Styles part does not exist, add it.
            if (part == null)
                WordDocUtilities.AddNumberingPartToPackage(WordDocument, root);
            else
                root.Save(part);
        }

        public void SetStylesFromDocument(string stylesFile)
        {
            XDocument styles = WordDocUtilities.ExtractStylesPart(stylesFile, false);

            // Get the Styles part for this document.
            StyleDefinitionsPart part =
                WordDocument.MainDocumentPart.StyleDefinitionsPart;

            Styles root = new Styles(styles.ToString());

            // If the Styles part does not exist, add it.
            if (part == null)
                WordDocUtilities.AddStylesPartToPackage(WordDocument, root);
            else
                root.Save(part);
        }

        public void SetHeaderFromDocument(string headerFile)
        {
            List<XDocument> headers = WordDocUtilities.ExtractHeaderPart(headerFile);

            List<Header> root = null;

            if (headers != null && headers.Count > 0)
            {
                root = new List<Header>();

                foreach (XDocument header in headers)
                    root.Add(new Header(header.ToString()));
            }

            // Get the Header part for this document.
            IEnumerable<HeaderPart> part =
                WordDocument.MainDocumentPart.HeaderParts;

            // If the Header part does not exist, add it.
            if (part == null)
            {
                WordDocUtilities.AddHeaderPartToPackage(WordDocument, root);
            }
            else
            {
                WordDocument.MainDocumentPart.DeleteParts(WordDocument.MainDocumentPart.HeaderParts);

                // foreach (Header header in root)
                //     header.Save(part);
            }
        }

        #region IDisposable
        public void Dispose()
        {
            CloseAndDisposeOfDocument();
            if (_stream != null)
            {
                ((IDisposable)_stream).Dispose();
                _stream = null;
            }
        }
        #endregion

        #region Save
        private void CloseAndDisposeOfDocument()
        {
            if (WordDocument != null)
            {
                foreach (Comentario c in Comentarios)
                    c.ToWordDocument(WordDocument);
                foreach (Parte p in Partes)
                    p.ToWordDocument(WordDocument);

                WordDocument.Close();
                WordDocument.Dispose();
                WordDocument = null;
            }
        }

        public void SaveToFile(string fileName)
        {
            if (WordDocument != null)
                CloseAndDisposeOfDocument();

            if (_stream == null)
                throw new ArgumentException("This object has already been disposed of so you cannot save it!");

            using (var fs = File.Create(fileName))
            {
                _stream.WriteTo(fs);
            }
        }
        #endregion

        public static void StartNumbering(int numberingId)
        {
            _numberingId = numberingId;
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
