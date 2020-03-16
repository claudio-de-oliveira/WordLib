﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AbsWordDocument.Itens
{
    public enum TipoDeAlinhamento { NENHUM, ESQUERDO, DIREITO, CENTRO }
    public enum TipoDeMerge { NENHUM, RESTART, CONTINUE }
    public enum TipoDeCelula { NORMAL, HEADER, RESUME }

    public class Celula
    {
        private readonly List<OpenXmlElement> _runList;
        public int Width { get; set; }

        public TipoDeCelula TipoDeCelula { get; set; }

        public TipoDeAlinhamento Alinhamento { get; set; }
        public TipoDeMerge Merge { get; set; }

        public Celula()
        {
            _runList = new List<OpenXmlElement>();
            Alinhamento = TipoDeAlinhamento.NENHUM;
            Merge = TipoDeMerge.NENHUM;
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

        public TableCell ToTableCell()
        {
            TableCell tableCell = new TableCell();

            TableCellProperties tableCellProperties = new TableCellProperties();

            // Specify the width property of the table cell.  
            TableCellWidth tableCellWidth = new TableCellWidth() { Width = Width.ToString(), Type = TableWidthUnitValues.Dxa };

            TableCellMargin tableCellMargin = new TableCellMargin();
            LeftMargin leftMargin = new LeftMargin() { Width = "100", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin = new RightMargin() { Width = "100", Type = TableWidthUnitValues.Dxa };

            tableCellMargin.Append(leftMargin);
            tableCellMargin.Append(rightMargin);
            TableCellVerticalAlignment tableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties.Append(tableCellWidth);
            tableCellProperties.Append(tableCellMargin);
            tableCellProperties.Append(tableCellVerticalAlignment);

            switch (TipoDeCelula)
            {
                case TipoDeCelula.HEADER:
                    tableCellProperties.Append(new Shading() { Val = ShadingPatternValues.Percent10, Color = "000000", Fill = "auto" });
                    break;
                case TipoDeCelula.RESUME:
                    tableCellProperties.Append(new Shading() { Val = ShadingPatternValues.Percent10, Color = "000000", Fill = "auto" });
                    break;
                default: // TipoDeCelula.NORMAL
                    break;
            }

            switch (Merge)
            {
                case TipoDeMerge.RESTART:
                    tableCellProperties.Append(new HorizontalMerge() { Val = MergedCellValues.Restart });
                    break;
                case TipoDeMerge.CONTINUE:
                    tableCellProperties.Append(new HorizontalMerge() { Val = MergedCellValues.Continue });
                    break;
                default: // TipoDeMerge.NENHUM
                    break;
            }

            tableCell.Append(tableCellProperties);

            Paragraph paragraph = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();

            switch (Alinhamento)
            {
                case TipoDeAlinhamento.ESQUERDO:
                    paragraphProperties.Append(new ParagraphStyleId() { Val = "LeftTextTable" });
                    break;
                case TipoDeAlinhamento.CENTRO:
                    paragraphProperties.Append(new ParagraphStyleId() { Val = "CenteredTextTable" });
                    break;
                case TipoDeAlinhamento.DIREITO:
                    paragraphProperties.Append(new ParagraphStyleId() { Val = "RightTextTable" });
                    break;
                default:  //TipoDeAlinhamento.NENHUM:
                    paragraphProperties.Append(new ParagraphStyleId() { Val = "NormalTextTable" });
                    break;
            }

            paragraph.AppendChild(paragraphProperties);

            foreach (OpenXmlElement run in _runList)
            {
                switch (TipoDeCelula)
                {
                    case TipoDeCelula.HEADER:
                        if (!run.Elements<RunProperties>().Any())
                            run.AppendChild(new RunProperties());
                        run.Elements<RunProperties>().First().AppendChild(new Bold() { Val = OnOffValue.FromBoolean(true) });
                        break;

                    case TipoDeCelula.RESUME:
                        if (!run.Elements<RunProperties>().Any())
                            run.AppendChild(new RunProperties());
                        run.Elements<RunProperties>().First().AppendChild(new Italic() { Val = OnOffValue.FromBoolean(true) });
                        run.Elements<RunProperties>().First().AppendChild(new Bold() { Val = OnOffValue.FromBoolean(true) });
                        break;
                }

                paragraph.AppendChild(run);
            }

            // Write some text in the cell.
            tableCell.Append(paragraph);

            return tableCell;
        }
    }

    public class Linha
    {
        private readonly Celula[] _celulas;

        private TipoDeCelula _tipoDeCelula;
        public TipoDeCelula TipoDeCelula {
            get {
                return _tipoDeCelula;
            }
            set {
                _tipoDeCelula = value;
                foreach (Celula cell in _celulas) cell.TipoDeCelula = _tipoDeCelula;
            }
        }

        // private bool _header;
        // public bool Header { get { return _header; } set { _header = value; foreach (Celula cell in _celulas) cell.Header = _header; } }
        // private bool _resume;
        // public bool Resume { get { return _resume; } set { _resume = value; foreach (Celula cell in _celulas) cell.Resume = _resume; } }

        public Linha(int columns)
        {
            _celulas = new Celula[columns];
            for (int i = 0; i < _celulas.Length; i++)
                _celulas[i] = new Celula();
            TipoDeCelula = TipoDeCelula.NORMAL;
            // Header = false;
            // Resume = false;
        }

        public Celula this[int i] {
            get { return _celulas[i]; }
        }
        public int Length {
            get { return _celulas.Length; }
        }

        public TableRow ToTableRow()
        {
            // Create a row and a cell.  
            TableRow tableRow = new TableRow();

            TablePropertyExceptions tablePropertyExceptions = new TablePropertyExceptions();

            TableCellMarginDefault tableCellMarginDefault = new TableCellMarginDefault();
            TopMargin topMargin = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };

            tableCellMarginDefault.Append(topMargin);
            tableCellMarginDefault.Append(bottomMargin);

            tablePropertyExceptions.Append(tableCellMarginDefault);

            // TableRowProperties tableRowProperties = new TableRowProperties();
            // TableRowHeight tableRowHeight = new TableRowHeight() { Val = (UInt32Value)236U };

            // tableRowProperties1.Append(tableRowHeight1);

            if (TipoDeCelula == TipoDeCelula.RESUME)
            {
                TableRowHeight tableRowHeight = new TableRowHeight() { Val = (UInt32Value)600U };

                TableRowProperties tableRowProperties = new TableRowProperties();

                tableRowProperties.Append(tableRowHeight);

                tableRow.Append(tableRowProperties);
            }

            tableRow.Append(tablePropertyExceptions);

            // Append the cell to the row.  
            for (int col = 0; col < _celulas.Length; col++)
                tableRow.Append(_celulas[col].ToTableCell());

            return tableRow;
        }
    }

    public class Tabela : Paragrafo
    {
        private readonly Linha[] _linhas;
        private int _width;

        public Tabela(int rows, int columns, int width, string style = "Normal")
            : base(style)
        {
            _linhas = new Linha[rows];

            int i;

            int[] CellWidth = new int[columns];
            int w = width / columns;

            for (i = 0; i < columns - 1; i++)
                CellWidth[i] = w;
            CellWidth[i] = width - (columns - 1) * w;

            for (i = 0; i < rows; i++)
                _linhas[i] = new Linha(columns);

            _width = width;
        }

        public Linha this[int row] {
            get { return _linhas[row]; }
        }

        public override void ToWordDocument(WordprocessingDocument wordDocument)
        {
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

            // TableCellMarginDefault
            TableCellMarginDefault tableCellMarginDefault = new TableCellMarginDefault();
            TableCellLeftMargin tableCellLeftMargin = new TableCellLeftMargin() { Width = 10, Type = TableWidthValues.Dxa };
            TableCellRightMargin tableCellRightMargin = new TableCellRightMargin() { Width = 10, Type = TableWidthValues.Dxa };

            tableCellMarginDefault.Append(tableCellLeftMargin);
            tableCellMarginDefault.Append(tableCellRightMargin);

            tableProperties.Append(tableCellMarginDefault);

            // TableLook
            TableLook tableLook = new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };
            tableProperties.Append(tableLook);

            // Shading
            // TableStyle
            // TablePositionProperties tablePositionProperties1 = new TablePositionProperties()
            // {
            //     LeftFromText = 141,
            //     RightFromText = 141,
            //     VerticalAnchor = VerticalAnchorValues.Text,
            //     TablePositionXAlignment = HorizontalAlignmentValues.Center,
            //     TablePositionY = 1
            // };

            // ***** NÃO ESTÁ FUNCIONANDO COM TableOverlap
            // TablePositionProperties tablePositionProperties = new TablePositionProperties();
            // tablePositionProperties.TablePositionXAlignment = HorizontalAlignmentValues.Center;
            // tablePositionProperties.VerticalAnchor = VerticalAnchorValues.Text;
            // tablePositionProperties.TablePositionY = 1;
            // tableProperties.Append(tablePositionProperties);

            // TableOverlap
            TableOverlap tableOverlap = new TableOverlap() { Val = TableOverlapValues.Never };
            tableProperties.Append(tableOverlap);

            tableProperties.Append(new Justification() { Val = JustificationValues.Center });

            // BiDiVisual

            // Make the table width 100% of the page width (50 * 100).
            TableWidth tableWidth = new TableWidth() { Width = _width.ToString(), Type = TableWidthUnitValues.Pct };
            tableProperties.Append(tableWidth);

            TableJustification tableJustification = new TableJustification() { Val = TableRowAlignmentValues.Center };
            tableProperties.Append(tableJustification);
            // TableCellSpacing
            // TableIndentation

            TableCaption tableCaption = new TableCaption() { Val = "Caption Table" };
            tableProperties.Append(tableCaption);

            // TableDescription
            // TablePropertiesChange

            TableGrid tableGrid = new TableGrid();

            for (int w = 0; w < this[0].Length; w++)
                tableGrid.Append(new GridColumn());

            table.Append(tableGrid);

            // Append the TableProperties object to the empty table.  
            table.AppendChild(tableProperties);

            // Create a row and a cell.  
            for (int row = 0; row < _linhas.Length; row++)
                table.Append(_linhas[row].ToTableRow());

            wordDocument.MainDocumentPart.Document.Body.AppendChild(table);
        }
    }
}



