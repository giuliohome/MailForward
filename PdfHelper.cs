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
        public string BusinessArea = "";
        public bool Active { get; set; } = true;
        public IEnumerable<FileInfo> pdfFilles;

        private const string On = "active";
        private const string Off = "inactive";
        public static async Task<List<Cpty>> Read(string csv_path)
        {
            var cpties = new List<Cpty>();
            if (String.IsNullOrWhiteSpace(csv_path) || !File.Exists(csv_path)) return cpties;
            using (var sr = new StreamReader(csv_path))
            {
                while (!sr.EndOfStream)
                {
                    var line = await sr.ReadLineAsync();
                    var fields = line.Split('\t');
                    if (fields.Length == 4)
                    {
                        cpties.Add(new Cpty() {
                            BusinessArea = fields[0],
                            Name = fields[1], EMail = fields[2],
                            Active = fields[3] == On });
                    }
                }
            }
            return cpties;
        }
        public static async Task Save(IEnumerable<Cpty> cpties, string csv_path, Action<string> log)
        {
            try
            {
                using (var sw = new StreamWriter(csv_path, false))
                {
                    foreach (var cpty in cpties)
                    {
                        await sw.WriteLineAsync($"{cpty.BusinessArea}\t{cpty.Name}\t{cpty.EMail}\t{(cpty.Active ? On : Off)}");
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

        public static IEnumerable<Cpty> ToCpties(string SelectedArea, IEnumerable<FileInfo> pdfFilles, IEnumerable<Cpty> savedCpties)
        {
            return pdfFilles
                .Select(pdf => new { Name = pdf.Name.Split('_')[0], pdf})
                .GroupBy(n => n.Name, (k,lst) => {
                    Cpty found = savedCpties.FirstOrDefault(c => 
                        c.Name == k && c.BusinessArea == SelectedArea);
                    return new Cpty()
                        {
                            Name = k, BusinessArea = SelectedArea,
                            EMail = found?.EMail ?? "",
                            Active = found?.Active ?? true,
                            pdfFilles = lst
                                .Select(elem => elem.pdf)
                        };
                    });
        }
    }
}
