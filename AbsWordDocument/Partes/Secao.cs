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
    public class Secao : Parte
    {
        public Secao(string text, string style = "Ttulo21")
            : base(text, style)
        {
            // Nothing more todo
        }

        public SubSecao NovaSubSecao(string text, string style = "Ttulo31")
        {
            SubSecao secao = new SubSecao(text, style);
            Children.Add(secao);
            return secao;
        }
    }
}
