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
    public class Celula : Paragrafo
    {
        private readonly List<OpenXmlElement> _runList;

        public bool Header { get; set; }

        public Celula(string style = "TabelaNormal")
            : base(style)
        {
            _runList = new List<OpenXmlElement>();
            Header = false;
        }

        public Celula Append(string text, RunProperties rPr = null)
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

        public TableCell ToTableCell(int width)
        {
            TableCell tableCell = new TableCell();

            TableCellProperties properties = new TableCellProperties();

            // Specify the width property of the table cell.  
            properties.Append(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = width.ToString() }
                );

            if (Header)
            {
                properties.Append(new Shading() { Val = ShadingPatternValues.Percent10 });
            }

            tableCell.Append(properties);

            Paragraph paragraph = base.CreateParagraph();

            foreach (OpenXmlElement element in _runList)
                paragraph.AppendChild(element);


            // Write some text in the cell.
            tableCell.Append(paragraph);

            return tableCell;
        }

        public override void ToWordDocument(WordprocessingDocument wordDocument)
        {
            throw new InvalidOperationException("Celula.ToWordDocument() não pode ser invocada!");
        }
    }

    public class Tabela : Paragrafo
    {
        private readonly Celula[][] _cells;
        public int[] Width;

        public Tabela(int r, int c, int width, string style = "Normal")
            : base(style)
        {
            _cells = new Celula[r][];

            int i;

            for (i = 0; i < r; i++)
            {
                _cells[i] = new Celula[c];
                for (int col = 0; col < c; col++)
                    _cells[i][col] = new Celula();
            }

            Width = new int[c];
            int w = width / c;
            for (i = 0; i < c - 1; i++)
                Width[i] = w;
            Width[i] = width - (c - 1) * w;
        }

        public Celula GetCelula(int row, int col)
        {
            return _cells[row][col];
        }

        public override void ToWordDocument(WordprocessingDocument wordDocument)
        {
            Paragraph paragraph = base.CreateParagraph();

            // Create an empty table.  
            Table table = new Table();

            // Create a TableProperties object and specify its border information.  
            TableProperties tableProperties = new TableProperties(
                new TableBorders(
                    new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Double), Size = 1 },
                    new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Double), Size = 1 },
                    new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Double), Size = 1 },
                    new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Double), Size = 1 },
                    new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 1 },
                    new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 1 }
                    )
                );

            // Append the TableProperties object to the empty table.  
            table.AppendChild<TableProperties>(tableProperties);

            for(int row = 0; row < _cells.Length; row++)
            {
                // Create a row and a cell.  
                TableRow tableRow = new TableRow();

                for (int col = 0; col < _cells[row].Length; col++)
                {
                    TableCell tableCell = _cells[row][col].ToTableCell(Width[col]);

                    // Append the cell to the row.  
                    tableRow.Append(tableCell);
                }
                table.Append(tableRow);
            }

            wordDocument.MainDocumentPart.Document.Body.AppendChild(table);
        }
    }
}



