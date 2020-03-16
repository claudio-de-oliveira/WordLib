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
    public class SubSecao : Parte
    {
        public SubSecao(string text, string style = "Ttulo31")
            : base(text, style)
        {
            // Nothing more todo
        }

        public SubSubSecao NovaSubSecao(string text, string style = "Ttulo41")
        {
            SubSubSecao secao = new SubSubSecao(text, style);
            Children.Add(secao);
            return secao;
        }
    }
}
