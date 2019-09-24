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

        private const string On = "active";
        private const string Off = "inactive";
        public static async Task Save(string SelectedArea, IEnumerable<Cpty> cpties, string csv_path, Action<string> log)
        {
            try
            {
                using (var sw = new StreamWriter(csv_path, false))
                {
                    foreach (var cpty in cpties)
                    {
                        await sw.WriteLineAsync($"{SelectedArea}\t{cpty.Name}\t{cpty.EMail}\t{(cpty.Active ? On : Off)}");
                    }
                }
            }
            catch (Exception exc)
            {
                log(exc.Message);
            }
        }
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
