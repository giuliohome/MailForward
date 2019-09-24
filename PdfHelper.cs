using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailForward
{
    public class Cpty
    {
        public string Name = "";
        public string EMail = "";
        public bool Active = true;
        public IEnumerable<FileInfo> pdfFilles;
    }

    public class PdfHelper
    {
        public static IEnumerable<Cpty> ToCpties(IEnumerable<FileInfo> pdfFilles)
        {
            return pdfFilles
                .Select(pdf => new { Name = pdf.Name.Split('_')[0], pdf})
                .GroupBy(n => n.Name, (k,lst) => 
                    new Cpty() {
                        Name = k,
                        pdfFilles = lst
                            .Select(elem => elem.pdf) });
        }
    }
}
