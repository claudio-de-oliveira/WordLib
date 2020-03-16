using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using AbsWordDocument.Itens;

namespace AbsWordDocument.Partes
{
    public class Capitulo : Parte
    {
        public Capitulo(string text, string style = "Ttulo11")
            : base(text, style)
        {
            // Nothing more todo
        }

        public Secao NovaSecao(string text, string style = "Ttulo21")
        {
            Secao secao = new Secao(text, style);
            Children.Add(secao);
            return secao;
        }
    }
}
