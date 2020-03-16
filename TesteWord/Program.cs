using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Xml;
using System.Xml.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using AbsWordDocument;
using AbsWordDocument.Itens;
using AbsWordDocument.Partes;

namespace TesteWord
{
    class Program
    {
        static Texto TestRunProperties()
        {
            // NÃO TESTADOS
            // ******************
            // Languages,
            // LocalName,
            // NoProof,
            // RunFonts,
            // RunPropertiesChange,
            // TextEffect,
            // SpecVanish,
            // WebHidden,
            // EastAsianLayout,
            // FitText,
            // Kern,

            return new Texto()
                .Append(
                    "Bold ",
                    new RunProperties() { Bold = new Bold() { Val = OnOffValue.FromBoolean(true) } }
                    )
                .Append(
                    "Caps ",
                    new RunProperties() { Caps = new Caps() { Val = OnOffValue.FromBoolean(true) } }
                    )
                .Append(
                    "Border ",
                    new RunProperties() { Border = new Border() { Val = BorderValues.BabyPacifier } }
                    )
                .Append(
                    "Color ",
                    new RunProperties() { Color = new Color() { Val = StringValue.FromString("0000FF") } }
                    )
                .Append(
                    "DoubleStrike ",
                    new RunProperties() { DoubleStrike = new DoubleStrike() { Val = OnOffValue.FromBoolean(true) } }
                    )
                .Append(
                    "Emboss ",
                    new RunProperties() { Emboss = new Emboss() { Val = OnOffValue.FromBoolean(true) } }
                    )
                .Append(
                    "CharacterScale ",
                    new RunProperties() { CharacterScale = new CharacterScale() { Val = IntegerValue.FromInt64(75) } } // Percentual
                    )
                .Append(
                    "Emphasis ",
                    new RunProperties() { Emphasis = new Emphasis() { Val = EmphasisMarkValues.UnderDot } }
                    )
                .Append(
                    "FontSize ",
                    new RunProperties() { FontSize = new FontSize() { Val = StringValue.FromString("14pt") } }
                    )
                .Append(
                    "Imprint ",
                    new RunProperties() { Imprint = new Imprint() { Val = OnOffValue.FromBoolean(true) } }
                    )
                .Append(
                    "Italic ",
                    new RunProperties() { Italic = new Italic() { Val = OnOffValue.FromBoolean(true) } }
                    )
                .Append(
                    "Outline ",
                    new RunProperties() { Outline = new Outline() { Val = OnOffValue.FromBoolean(true) } }
                    )
                .Append(
                    "Shading ",
                    new RunProperties() { Shading = new Shading() { Val = ShadingPatternValues.HorizontalStripe } }
                    )
                .Append(
                    "Shadow ",
                    new RunProperties() { Shadow = new Shadow() { Val = OnOffValue.FromBoolean(true) } }
                    )
                .Append(
                    "SmallCaps ",
                    new RunProperties() { SmallCaps = new SmallCaps() { Val = OnOffValue.FromBoolean(true) } }
                    )
                .Append(
                    "Strike ",
                    new RunProperties() { Strike = new Strike() { Val = OnOffValue.FromBoolean(true) } }
                    )
                .Append(
                    "Underline ",
                    new RunProperties() { Underline = new Underline() { Val = UnderlineValues.DashLongHeavy } }
                    )
                .Append(
                    "Vanish ",
                    new RunProperties() { Vanish = new Vanish() { Val = OnOffValue.FromBoolean(true) } }
                    )
                .Append(
                    "Spacing ",
                    new RunProperties() { Spacing = new Spacing() { Val = 200 } }
                    )
                .Append(
                    "SpecVanish ",
                    new RunProperties() { SpecVanish = new SpecVanish() { Val = OnOffValue.FromBoolean(false) } }
                    )
                .Append(
                    "VerticalTextAlignment ",
                    new RunProperties() { VerticalTextAlignment = new VerticalTextAlignment() { Val = VerticalPositionValues.Subscript } }
                    )
                .Append(
                    "Position ",
                    new RunProperties() { Position = new Position() { Val = StringValue.FromString("6") } }
                    )
                .Append(
                    "FontSizeComplexScript ",
                    new RunProperties() { FontSizeComplexScript = new FontSizeComplexScript() { Val = "40" } }
                    );
        }

        static Texto TestComment(WordDoc wordDoc)
        {
            Comentario comentario = wordDoc.CreateComment();
            comentario.Append(
                "Descrição do comentário",
                new RunProperties() { Caps = new Caps() { Val = OnOffValue.FromBoolean(true) } }
                );

            Texto texto = new Texto();

            texto.Append("Aqui vai começar um comentário, ");
            texto.StartComment(comentario.Id);
            texto.Append(
                "Texto de comentário, ",
                new RunProperties() { Bold = new Bold() { Val = OnOffValue.FromBoolean(true) } }
                );
            texto.EndComment(comentario.Id);
            texto.Append("Texto fora do comentário,");

            return texto;
        }

        static Capitulo TestList(WordDoc wordDoc)
        {
            Capitulo capitulo = new Capitulo("Teste de Listas");

            Paragrafo.StartNumbering(wordDoc.WordDocument, 2);
            capitulo.Paragrafos.Add(
                new Texto()
                .Append("Item 1")
                );
            capitulo.Paragrafos.Add(
                new Texto()
                .Append("Item 2")
                );
            Paragrafo.IncrementNumbering();
            capitulo.Paragrafos.Add(
                new Texto()
                .Append("Item 1.1")
                );
            capitulo.Paragrafos.Add(
                new Texto()
                .Append("Item 1.2")
                );
            Paragrafo.DecrementNumbering();
            capitulo.Paragrafos.Add(
                new Texto()
                .Append("Item 3")
                );
            capitulo.Paragrafos.Add(
                new Texto()
                .Append("Item 4")
                );
            Paragrafo.EndNumbering();

            Secao secao = capitulo.NovaSecao("Listas em Seção");

            Paragrafo.StartNumbering(wordDoc.WordDocument, 2);
            secao.Paragrafos.Add(
                new Texto()
                .Append("Item 1")
                );
            secao.Paragrafos.Add(
                new Texto()
                .Append("Item 2")
                );
            Paragrafo.IncrementNumbering();
            secao.Paragrafos.Add(
                new Texto()
                .Append("Item 1.1")
                );
            secao.Paragrafos.Add(
                new Texto()
                .Append("Item 1.2")
                );
            Paragrafo.IncrementNumbering();
            secao.Paragrafos.Add(
                new Texto()
                .Append("Item 1.2.1")
                );
            secao.Paragrafos.Add(
                new Texto()
                .Append("Item 1.2.2")
                );
            Paragrafo.DecrementNumbering();
            Paragrafo.DecrementNumbering();
            secao.Paragrafos.Add(
                new Texto()
                .Append("Item 3")
                );
            secao.Paragrafos.Add(
                new Texto()
                .Append("Item 4")
                );
            Paragrafo.EndNumbering();

            return capitulo;
        }

        static Tabela TestTabel()
        {
            Tabela tabela = new Tabela(3, 3, 5000);

            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    tabela[i][j].Append($"Célula {i + 1}-{j + 1}");

                    tabela[i][j].Alinhamento = TipoDeAlinhamento.ESQUERDO;

                    if (i == 0)
                        tabela[i].TipoDeCelula = TipoDeCelula.HEADER;
                    else if (i == 2)
                        tabela[i].TipoDeCelula = TipoDeCelula.RESUME;
                }
            }

            tabela[2][0].Merge = TipoDeMerge.RESTART;
            tabela[2][1].Merge = TipoDeMerge.CONTINUE;
            tabela[2][2].Merge = TipoDeMerge.RESTART;

            return tabela;
        }

        static void Main(string[] args)
        {
            //string filepath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            // string filepath = @"C:\Users\clalu\source\repos\WordLib";
            string filepath = @"C:\Users\claudio_oliveira\Source\Repos\WordLib";

            WordDoc wordDoc = new WordDoc("Claudio de Oliveira", "CdO");

            // string stylesFile = Path.Combine(filepath, "Minuta de Mecânica (21-02-20).docx");
            // wordDoc.SetHeaderFromDocument(stylesFile);

            // Paragraph p = WordDocUtilities.CreateParagraphWithStyle("Teste do Título 2", "Ttulo2");

            // wordDoc.SetStylesFromDocument(stylesFile);
            WordDocUtilities.AddStyleDefinitionsPartToPackage(wordDoc.WordDocument);
            wordDoc.WordDocument.MainDocumentPart.StyleDefinitionsPart.Styles = WordDocUtilities.GenerateStyleDefinitionsPartContent();

            // wordDoc.SetNumberingFromDocument(stylesFile);
            WordDocUtilities.AddNumberingPartToPackage(wordDoc.WordDocument);
            wordDoc.WordDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering = WordDocUtilities.GenerateNumberingDefinitionsPartContent();

            Capitulo capitulo;
            Secao secao;
            Texto texto;
            Tabela tabela;

            capitulo = TestList(wordDoc);
            texto = TestRunProperties();
            capitulo.Paragrafos.Add(texto);
            texto = TestComment(wordDoc);
            capitulo.Paragrafos.Add(texto);
            wordDoc.Partes.Add(capitulo);
            tabela = TestTabel();
            capitulo.Paragrafos.Add(tabela);

            capitulo = new Capitulo(
                "Construção, Implantação e consolidação do Projeto Pedagógico de Curso"
                );

            capitulo.Paragrafos.Add(
                new Texto()
                    .Append(
                        "O Projeto Pedagógico do Curso (PPC) de Engenharia Civil da Universidade Tiradentes – Unit " +
                        "é resultado da construção das diretrizes organizacionais, " +
                        "estruturais e pedagógicas, com a participação do corpo docente do curso por " +
                        "meio de seus representantes no Núcleo Docente Estruturante (NDE) e Colegiado. " +
                        "Encontra-se articulado com as bases legais e a concepção de formação profissional " +
                        "que favoreça o desenvolvimento de competências e habilidades necessárias ao " +
                        "exercício da profissão, como a capacidade de observação, criticidade e " +
                        "questionamento, sintonizada com a dinâmica da sociedade nas suas demandas " +
                        "locais, regionais e nacionais, assim como com os avanços científicos e " +
                        "tecnológicos. O referido documento surge a partir da criação do curso, autorizado " +
                        "pela Portaria CONSAD nº 008 de 08/04/2010 tendo como objetivo principal o " +
                        "atendimento aos princípios e diretrizes do Projeto Pedagógico Institucional, " +
                        "Diretrizes Curriculares Nacionais, Pareceres do CNE e indicadores de qualidade " +
                        "do Inep/MEC."
                        )
                    );
            capitulo.Paragrafos.Add(
                new Texto()
                    .Append(
                        "A construção do PPC ocorre, afirmativamente, ancorada em uma ação intencional, refletida e fundamentada no coletivo de sujeitos, agentes interessados em promover a missão da Universidade de inspirar as pessoas a ampliar horizontes por meio do ensino, pesquisa e extensão, com ética e compromisso com o desenvolvimento social. Desta forma, o Projeto Pedagógico do Curso de Graduação em Engenharia Civil da Unit está em conformidade com as Diretrizes Curriculares Nacionais para os cursos de Graduação em Engenharia Civil, Projeto Pedagógico Institucional da Unit – PPI e seu Plano de Desenvolvimento Institucional – PDI, fundamentado nas necessidades socioeconômicas, políticas, educacionais, demandas do mercado de trabalho no Estado de Sergipe e as condições institucionais da IES para expansão da oferta de cursos na área da saúde."
                        )
                    );
            capitulo.Paragrafos.Add(
                new Texto()
                    .Append(
                        "Cônscia de sua responsabilidade com a sociedade e com o desenvolvimento de Sergipe e do Nordeste, a Unit mantém o curso de Engenharia Civil tendo por base os princípios preconizados na Lei nº 9.394, de 20 de dezembro de 1996, que enfatiza a importância da construção dos conhecimentos mediante políticas e planejamentos educacionais, capazes de garantir o padrão de qualidade no ensino, flexibilizando a ação educativa, valorizando a experiência do aluno, respeitando o pluralismo de ideias e princípios básicos da democracia."
                        )
                    );
            capitulo.Paragrafos.Add(
                new Texto()
                    .Append(
                        "O PPC está organizado de modo a contemplar os critérios indispensáveis à formação de um profissional dotado das competências essenciais para o exercício profissional frente ao contexto sócio-econômico – cultural e político da região e do país."
                        )
                    );
            capitulo.Paragrafos.Add(
                new Texto()
                    .Append(
                        "A proposta conceitual e metodológica é entendida como um conjunto de cenários em que há a construção do perfil do estudante a partir da aprendizagem significativa, que promove e produz sentidos.Esta proposta está em conformidade com os princípios da UNESCO, isto é, educar para fazer, para aprender, para sentir e para ser; busca-se a construção de uma visão da realidade e de situações excepcionais e singulares na qual atuará o futuro profissional com o compromisso de transformar a realidade em que vive."
                        )
                    );
            capitulo.Paragrafos.Add(
                new Texto()
                    .Append(
                        "Nesse contexto, a Unit se compromete com a oferta de um curso de relevância social que assegura a qualidade na formação acadêmica, com vistas a atender as necessidades da população tanto local como nas regiões circunvizinhas como pilares essenciais para a construção da cidadania."
                        )
                    );
            wordDoc.Partes.Add(capitulo);

            capitulo = new Capitulo("Dados de identificação");
            capitulo.NovaSecao("Identificação");
            capitulo.NovaSecao("Legislação e normas que regem o curso");
            wordDoc.Partes.Add(capitulo);

            capitulo = new Capitulo("Contextualização e Justificativa de Oferta do Curso");
            wordDoc.Partes.Add(capitulo);

            capitulo.Paragrafos.Add(
                new Figura(@"C:\Users\clalu\OneDrive\Documentos\Imagens\IMG10083.JPG", .3)
                );

            capitulo = new Capitulo("Objetivos do Curso");
            capitulo.NovaSecao("Geral");
            capitulo.NovaSecao("Específicos");
            secao = capitulo.NovaSecao("Campo de Atuação");

            Paragrafo.StartNumbering(wordDoc.WordDocument, 2);
            secao.Paragrafos.Add(
                new Texto()
                .Append("Item 1")
                );
            secao.Paragrafos.Add(
                new Texto()
                .Append("Item 2")
                );
            Paragrafo.IncrementNumbering();
            secao.Paragrafos.Add(
                new Texto()
                .Append("Item 1.1")
                );
            secao.Paragrafos.Add(
                new Texto()
                .Append("Item 1.2")
                );
            Paragrafo.IncrementNumbering();
            secao.Paragrafos.Add(
                new Texto()
                .Append("Item 1.2.1")
                );
            secao.Paragrafos.Add(
                new Texto()
                .Append("Item 1.2.2")
                );
            Paragrafo.DecrementNumbering();
            Paragrafo.DecrementNumbering();
            secao.Paragrafos.Add(
                new Texto()
                .Append("Item 3")
                );
            secao.Paragrafos.Add(
                new Texto()
                .Append("Item 4")
                );
            Paragrafo.EndNumbering();

            wordDoc.Partes.Add(capitulo);

            capitulo = new Capitulo("Perfil do Egresso");
            wordDoc.Partes.Add(capitulo);

            capitulo = new Capitulo("Forma de Acesso ao Curso");
            wordDoc.Partes.Add(capitulo);

            capitulo = new Capitulo("Sistemas de Avaliação");
            capitulo.NovaSecao("Sistema de Avaliação do Projeto do Curso");
            capitulo.NovaSecao("Atividades Complementares");
            secao = capitulo.NovaSecao("Práticas Profissionais e Estágio");
            secao.NovaSubSecao("Estágio Extracurricular");
            wordDoc.Partes.Add(capitulo);

            capitulo = new Capitulo("Corpo Docente");
            wordDoc.Partes.Add(capitulo);

            capitulo = new Capitulo("Estrutura Curricular");
            wordDoc.Partes.Add(capitulo);

            capitulo = new Capitulo("Módulos Curriculares");
            capitulo.NovaSecao("Primeiro Período");
            capitulo.NovaSecao("Segundo Período");
            capitulo.NovaSecao("Terceiro Período");
            capitulo.NovaSecao("Quarto Período");
            capitulo.NovaSecao("Quinto Período");
            capitulo.NovaSecao("Sexto Período");
            capitulo.NovaSecao("Sétimo Período");
            capitulo.NovaSecao("Oitavo Período");
            capitulo.NovaSecao("Nono Período");
            capitulo.NovaSecao("Décimo Período");
            wordDoc.Partes.Add(capitulo);

            string filename = Path.Combine(filepath, "tese.docx");
            wordDoc.SaveToFile(filename);
        }
    }
}
