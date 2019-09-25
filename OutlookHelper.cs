using System;
using System.Collections.Generic;
#if DEBUG
using System.Diagnostics;
#endif
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Configuration;
using System.IO;
using MVVM.ViewModel;
using Microsoft.Win32;

namespace MailForward
{
    internal class OutlookHelper : ViewModelBase
    {
        private Outlook.Application application;
        internal OutlookHelper()
        {
        }

        internal async Task StartOutlook()
        {
            await Task.Run(() => application = new Outlook.Application());
        }

        internal async Task ReadConfig()
        {
            string settingsPath = ConfigurationManager.AppSettings["settingsPath"];
            if (File.Exists(settingsPath))
            {
                using (var sr = new StreamReader(settingsPath))
                {
                    var linesMap = new Dictionary<string, List<string>>();
                    foreach (string area in Areas)
                    {
                        linesMap[area] = new List<string>();
                    }
                    while (!sr.EndOfStream)
                    {
                        string line = await sr.ReadLineAsync();
                        if (String.IsNullOrWhiteSpace(line)) continue;
                        string[] fields = line.Split('\t');
                        if (fields.Length != 3 || !Areas.Contains(fields[0]) ) continue;
                        switch (fields[1])
                        {
                            case CsvFolderPath:
                                FolderPathMap[fields[0]] = fields[2];
                                continue;
                            case CsvFolderName:
                                FolderNameMap[fields[0]] = fields[2];
                                continue;
                            case CsvAddressTo:
                                AddressToMap[fields[0]] = fields[2];
                                continue;
                            case CsvAddressCc:
                                AddressCcMap[fields[0]] = fields[2];
                                continue;
                            case CsvForwardedTxt:
                                linesMap[fields[0]].Add(fields[2]);
                                continue;
                            default:
                                continue;
                        }
                    }
                    foreach (string area in Areas)
                    {
                        ForwardedTxtMap[area] = String.Join("\n", linesMap[area]);
                    }
                }
                var folderPath = FolderPathMap.ContainsKey(SelectedArea) ? FolderPathMap[SelectedArea] : "";
                if (SelectedArea == AuthFwd && folderPath != null)
                {
                    var folderKeys = folderPath.Split('\\');
                    if (folderKeys.Length == 2)
                    {
                        folderStoreID = folderKeys[0];
                        folderEntryID = folderKeys[1];
                    }
                } else
                {
                    folderStoreID = "";
                    folderEntryID = "";
                }
                SelFolderName = FolderNameMap.ContainsKey(SelectedArea) ? FolderNameMap[SelectedArea] : "";
                AddressTo = AddressToMap.ContainsKey(SelectedArea) ? AddressToMap[SelectedArea] : "";
                AddressCc = AddressCcMap.ContainsKey(SelectedArea) ? AddressCcMap[SelectedArea] : "";
                ForwardedTxt = ForwardedTxtMap.ContainsKey(SelectedArea) ? ForwardedTxtMap[SelectedArea] : "";
            }

        }
        private const string CsvFolderPath = "Folder Path";
        private const string CsvFolderName = "Folder Name";
        private const string CsvAddressTo = "Address To";
        private const string CsvAddressCc = "Address Cc";
        private const string CsvForwardedTxt = "Forwarded Text";
        internal async Task SaveConfig()
        {
            FolderPathMap[SelectedArea] = folderStoreID + "\\" + folderEntryID;
            FolderNameMap[SelectedArea] = selFolderName;
            AddressToMap[SelectedArea] = AddressTo;
            AddressCcMap[SelectedArea] = AddressCc;
            ForwardedTxtMap[SelectedArea] = ForwardedTxt;
            using (var sw = new StreamWriter(ConfigurationManager.AppSettings["settingsPath"], false))
            {
                foreach (string area in Areas)
                {
                    if (FolderNameMap.ContainsKey(area))
                    {
                        await sw.WriteLineAsync($"{area}\t{CsvFolderName}\t{FolderNameMap[area].Replace("\t", " ")}");
                    }
                    if (FolderPathMap.ContainsKey(area))
                    {
                        await sw.WriteLineAsync($"{area}\t{CsvFolderPath}\t{FolderPathMap[area].Replace("\t", " ")}");
                    }
                    if (AddressToMap.ContainsKey(area))
                    {
                        await sw.WriteLineAsync($"{area}\t{CsvAddressTo}\t{AddressToMap[area].Replace("\t", " ")}");
                    }
                    if (AddressCcMap.ContainsKey(area))
                    {
                        if (!String.IsNullOrWhiteSpace(AddressCcMap[area]))
                        {
                            await sw.WriteLineAsync($"{area}\t{CsvAddressCc}\t{AddressCcMap[area].Replace("\t", " ")}");
                        }
                    }
                    if (ForwardedTxtMap.ContainsKey(area))
                    {
                        if (!String.IsNullOrWhiteSpace(ForwardedTxtMap[area]))
                        {
                            foreach (string line in ForwardedTxtMap[area].Replace("\t", " ").Split('\n'))
                            {
                                await sw.WriteLineAsync($"{area}\t{CsvForwardedTxt}\t{line}");
                            }

                        }
                    }
                }
            }

        }

        private const string NotSelected = "Not Selected";
        private string selFolderName = NotSelected;
        public string SelFolderName
        {
            get { return selFolderName; }
            set
            {
                selFolderName = value;
                OnPropertyChanged();
            }
        }
        private string folderEntryID = null;
        private string folderStoreID = null;
        internal async Task SelectFolder()
        {
            await ReadConfig();
            if (SelectedArea == AuthFwd)
            {
                Outlook.Folder selectedFolder = await Task.Run(() => application.Session.PickFolder() as Outlook.Folder);
                folderEntryID = selectedFolder?.EntryID;
                folderStoreID = selectedFolder?.StoreID;
                SelFolderName = selectedFolder?.FolderPath ?? NotSelected;
            } else
            {
                var dialog = new OpenFileDialog();
                dialog.Title = "select a pdf for " + SelectedArea;
                dialog.DefaultExt = ".pdf";
                var res = dialog.ShowDialog();
                if (res??false == true)
                {
                    SelFolderName = Path.GetDirectoryName(dialog.FileName);
                } else
                {
                    SelFolderName = NotSelected;
                    folderStoreID = "";
                    folderEntryID = "";
                }
            }
            await SaveConfig();
        }

        private string status = "";
        public string Status
        {
            get { return status; }
            set
            {
                status = value;
                OnPropertyChanged();
            }
        }

        private Dictionary<string, string> FolderPathMap = new Dictionary<string, string>();
        private Dictionary<string, string> FolderNameMap = new Dictionary<string, string>();

        private Dictionary<string, string> AddressToMap = new Dictionary<string, string>();
        private Dictionary<string, string> AddressCcMap = new Dictionary<string, string>();
        private Dictionary<string, string> ForwardedTxtMap = new Dictionary<string, string>();

        public string AddressTo { get; set; } = "Name Surname <Name.Surname@xxx.com>; "; 
        public string AddressCc { get; set; } = "";

        public string ForwardedTxt { get; set; } = "";
        const string PR_SMTP_ADDRESS =
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        private void ComposeMail(Outlook.MailItem newItem)
        {
            if (!String.IsNullOrWhiteSpace(AddressCc))
            {
                newItem.CC = AddressCc;
            }
            newItem.Body = ForwardedTxt + newItem.Body;
            newItem.Importance = Outlook.OlImportance.olImportanceHigh;
            newItem.Display(false);
            //newItem.Save();
        }
        private string cpty_path = ConfigurationManager.AppSettings["cptiesPath"];
        internal async Task ForwardItems()
        {
            Status = "Forwarding... pls wait!";
            Outlook.Folder folder = null;
            Cpty[] cpties = new Cpty[0];
            IEnumerable<FileInfo> pdfFilles = null;
            if (SelectedArea == AuthFwd)
            {
                folder = GetOutlookFolder();
                if (folder == null)
                {
                    return;
                }
                if (folder.Items.Count == 0)
                {
                    Status = "No folder/items";
                    return;
                }
            } else
            {
                if (!GetPdfFiles(out pdfFilles)) return;
            }
            try
            {
                if (SelectedArea != AuthFwd)
                {
                    var savedCpties = await Cpty.Read(cpty_path);
                    cpties = PdfHelper.ToCpties(SelectedArea, pdfFilles, savedCpties).ToArray();
                    var dialog = new Counterparties();
                    dialog.DataContext = new PdfHelper() { Cpties = cpties };
                    dialog.ShowDialog();
                    var allCpties = cpties.Concat(savedCpties.Where(s => s.BusinessArea != SelectedArea));
                    await Cpty.Save(allCpties, ConfigurationManager.AppSettings["cptiesPath"], txt => Status = txt);
                }
                await Task.Run(() =>
                {

                    if (SelectedArea == AuthFwd)
                    {
                        foreach (var obj in folder.Items)
                        {
                            if (obj is Outlook.MailItem mailItem)
                            {
                                #if DEBUG
                                Debug.WriteLine("forwarding mail: " + mailItem.Subject);
                                var recipientNames = new List<string>();
                                foreach (var objRecipient in mailItem.Recipients)
                                {
                                    if (objRecipient is Outlook.Recipient recipient)
                                    {
                                        Outlook.PropertyAccessor pa = recipient.PropertyAccessor;
                                        string smtpAddress =
                                            pa.GetProperty(PR_SMTP_ADDRESS).ToString();
                                        recipientNames.Add($"{recipient.Name} <{smtpAddress}>");
                                    }
                                }
                                Debug.WriteLine($"sent to {String.Join("; ", recipientNames)}");
                                #endif

                                var newItem = mailItem.Forward();
                                newItem.To = AddressTo;
                                ComposeMail(newItem);
                            }
                        }

                    } else
                    {
                        foreach (var cpty in cpties.Where(c => c.Active))
                        {
                            Outlook.MailItem mail = application.CreateItem(
                                Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                            mail.Subject = SelectedArea + " - Netting Statement: " + cpty.Name;
                            mail.To = cpty.EMail;
                            foreach (var pdf in cpty.pdfFilles)
                            {
                                mail.Attachments.Add(pdf.FullName,
                                    Outlook.OlAttachmentType.olByValue, Type.Missing,
                                    Type.Missing);
                            }
                            ComposeMail(mail);
                        }
                    }
                }
                );
                Status = "Forward done.";
            }
            catch (Exception exc)
            {
                Status = exc.Message;
            }
        }

        private Outlook.Folder GetOutlookFolder()
        {
            try
            {
                if (String.IsNullOrEmpty(folderEntryID) || String.IsNullOrEmpty(folderStoreID))
                {

                    Status = "No folder selected";
                    return null;
                }
                return
                    application.Session.GetFolderFromID(
                        folderEntryID, folderStoreID)
                        as Outlook.Folder;
            }
            catch (Exception exc)
            {
                #if DEBUG
                Debug.WriteLine(exc.ToString());
                #endif
                Status = exc.Message;
                return null;
            }
        }

        private bool GetPdfFiles(out IEnumerable<FileInfo> pdfFilles)
        {
            pdfFilles = null;
            if (SelFolderName == NotSelected || String.IsNullOrWhiteSpace(SelFolderName))
            {
                Status = NotSelected;
                return false;
            }
            if (!Directory.Exists(SelFolderName))
            {
                Status = "Dir not found!";
                return false;
            }
            else
            {
                var info = new DirectoryInfo(SelFolderName);
                pdfFilles = info.EnumerateFiles("*.pdf");
                return true;
            }
        }

        internal void DisplayFolder()
        {
            if (SelectedArea == AuthFwd)
            {
                Outlook.Folder folder = GetOutlookFolder();
                if (folder == null)
                {
                    return;
                }
                try
                {

                    var actExpl = application.ActiveExplorer();
                    if (actExpl == null)
                    {
                        Status = "Pls, open Outlook";
                        return;
                    }
                    actExpl.CurrentFolder = folder; //folderFromID.Display();
                    Status = "Selected emails: " + (folder?.Items?.Count ?? 0);
                }
                catch (Exception exc)
                {
                    #if DEBUG
                    Debug.WriteLine(exc.Message);
                    #endif
                    Status = exc.Message;
                    folder = null;
                }
            } else
            {
                if (GetPdfFiles(out IEnumerable<FileInfo> pdfFilles))
                {
                    System.Diagnostics.Process.Start("explorer.exe", SelFolderName);
                    Status = "pdf files: " + pdfFilles.Count();
                }
            }
        }


        public const string AuthFwd = "Authorize";
        public string[] Areas => new string[] { AuthFwd, Business.Area1, Business.Area2, Business.Area3 };
        private string selArea = AuthFwd;
        public string SelectedArea {
            get { return selArea; }
            set
            {
                selArea = value;
                OnPropertyChanged();
                Task.Run( async () => await ReadConfig());
            }
        } 

    }
}
