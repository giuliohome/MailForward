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
        public string Name { get; set; } = "";
        public string EMail { get; set; } = "";
        public bool Active { get; set; } = true;
        public IEnumerable<FileInfo> pdfFilles;
    }

    public class PdfHelper
    {
        public Cpty[] Cpties { get; set; }

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
