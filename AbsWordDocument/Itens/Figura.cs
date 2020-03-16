using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace AbsWordDocument.Itens
{
    public class Figura : Paragrafo
    {
        private readonly string _fileName;
        private readonly double _width;

        private new readonly int _numberingLevel;
        private new readonly int _numberingId;
        private new readonly bool _numbering;

        public Figura(string fileName, double width, string style = "Normal")
            : base(style)
        {
            this._fileName = fileName;
            this._width = width;

            _numberingLevel = Paragrafo._numberingLevel;
            _numberingId = Paragrafo._numberingId;
            _numbering = Paragrafo._numbering;
        }

        public override void ToWordDocument(WordprocessingDocument wordDocument)
        {
            MainDocumentPart mainPart = wordDocument.MainDocumentPart;

            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

            try
            {
                long widthEmus = 0;
                long heightEmus = 0;

                // Recupera as dimensões da imagem
                {
                    BitmapImage img = new BitmapImage();

                    using (var fs = new FileStream(_fileName, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        img.BeginInit();
                        img.StreamSource = fs;
                        img.EndInit();
                    }

                    var widthPx = img.PixelWidth;
                    var heightPx = img.PixelHeight;
                    var horzRezDpi = img.DpiX;
                    var vertRezDpi = img.DpiY;

                    const int emusPerInch = 914400;
                    const int emusPerCm = 360000;

                    widthEmus = (long)(widthPx / horzRezDpi * emusPerInch);
                    heightEmus = (long)(heightPx / vertRezDpi * emusPerInch);

                    widthEmus = (long)(_width * widthEmus);
                    heightEmus = (long)(_width * heightEmus);

                    /***********************************************************************/
                    var maxWidthEmus = (long)(21 * emusPerCm); /***** LARGURA DA PÁGINA *****/
                    /***********************************************************************/

                    if (widthEmus > maxWidthEmus)
                    {
                        var ratio = (heightEmus * 1.0m) / widthEmus;
                        widthEmus = maxWidthEmus;
                        heightEmus = (long)(widthEmus * ratio);
                    }
                }

                using (FileStream stream = new FileStream(_fileName, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                Paragraph paragraph = base.CreateParagraph();

                if (_numbering)
                {
                    // Create items for paragraph properties
                    var numberingProperties = new NumberingProperties(new NumberingLevelReference() { Val = _numberingLevel }, new NumberingId() { Val = _numberingId });

                    // create paragraph properties
                    var paragraphProperties = new ParagraphProperties(numberingProperties);

                    paragraph.Append(paragraphProperties);
                }

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                Justification justification1 = new Justification() { Val = JustificationValues.Center };

                // Hanging = "0" => noindent 
                Indentation indentation = new Indentation { Hanging = "0" };

                paragraphProperties1.Append(justification1);
                paragraphProperties1.Append(indentation);

                string bookmarkId = WordDocUtilities.NewBookmark();
                BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = bookmarkId };

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                Drawing drawing1 = AddImageToBody(wordDocument, mainPart.GetIdOfPart(imagePart), widthEmus, heightEmus);

                run1.Append(runProperties1);
                run1.Append(drawing1);

                BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = WordDocUtilities.NewBookmark() };

                paragraph.Append(paragraphProperties1);
                paragraph.Append(bookmarkStart1);
                paragraph.Append(run1);
                paragraph.Append(bookmarkEnd1);

                // Append the reference to body, the element should be in a Run.
                wordDocument.MainDocumentPart.Document.Body.AppendChild(paragraph);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static Drawing AddImageToBody(WordprocessingDocument wordDoc, string relationshipId, long widthEmus, long heightEmus)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = widthEmus, Cy = heightEmus },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = String.Format("Imagem {0}", WordDocUtilities.NewFigureCounter())
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                        new A.Blip(
                                            new A.BlipExtensionList(
                                                new A.BlipExtension(
                                                    new A14.UseLocalDpi() { Val = false }
                                                    )
                                                { Uri = String.Format("{{{0}}}", Guid.NewGuid().ToString()) }
                                                )
                                            )
                                        {
                                            Embed = relationshipId,
                                            CompressionState = A.BlipCompressionValues.Print
                                        },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = widthEmus, Cy = heightEmus }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         AnchorId = "63B56EF4",
                         EditId = "50D07946"
                     });

            return element;
        }
    }
}
