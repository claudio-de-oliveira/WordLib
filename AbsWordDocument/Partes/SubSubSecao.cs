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
    public class SubSubSecao : Parte
    {
        public SubSubSecao(string text, string style = "Ttulo51")
            : base(text, style)
        {
            // Nothing more todo
        }
    }
}
