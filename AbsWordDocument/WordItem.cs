using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace AbsWordDocument
{
    public abstract class WordItem
    {
        public abstract void ToWordDocument(WordprocessingDocument wordDocument);
    }
}
