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

namespace MailForward
{
    internal class OutlookHelper : ViewModelBase
    {
        private Outlook.Application application;
        internal OutlookHelper()
        {
            application = new Outlook.Application();
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
                AddressTo = AddressToMap.ContainsKey(SelectedArea) ? AddressToMap[SelectedArea] : "";
                AddressCc = AddressCcMap.ContainsKey(SelectedArea) ? AddressCcMap[SelectedArea] : "";
                ForwardedTxt = ForwardedTxtMap.ContainsKey(SelectedArea) ? ForwardedTxtMap[SelectedArea] : "";
            }

        }
        private const string CsvAddressTo = "Address To";
        private const string CsvAddressCc = "Address Cc";
        private const string CsvForwardedTxt = "Forwarded Text";
        internal async Task SaveConfig()
        {
            AddressToMap[SelectedArea] = AddressTo;
            AddressCcMap[SelectedArea] = AddressCc;
            ForwardedTxtMap[SelectedArea] = ForwardedTxt;
            using (var sw = new StreamWriter(ConfigurationManager.AppSettings["settingsPath"], false))
            {
                foreach (string area in Areas)
                {
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

        private string selFolderName = "Not Selected";
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
            Outlook.Folder selectedFolder = await Task.Run(() => application.Session.PickFolder() as Outlook.Folder);
            folderEntryID = selectedFolder?.EntryID;
            folderStoreID = selectedFolder?.StoreID;
            SelFolderName = selectedFolder?.FolderPath ?? "Not Selected";
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

        private Dictionary<string, string> AddressToMap = new Dictionary<string, string>();
        private Dictionary<string, string> AddressCcMap = new Dictionary<string, string>();
        private Dictionary<string, string> ForwardedTxtMap = new Dictionary<string, string>();

        public string AddressTo { get; set; } = "Name Surname <Name.Surname@xxx.com>; "; 
        public string AddressCc { get; set; } = "";

        public string ForwardedTxt { get; set; } = "";
        const string PR_SMTP_ADDRESS =
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
        internal async Task ForwardItems()
        {
            Status = "Forwarding... pls wait!";
            Outlook.Folder folder = GetOutlookFolder();
            if (folder == null || folder.Items.Count == 0)
            {
                Status = "No folder/items";
                return;
            }
            try
            {

                await Task.Run(() =>
                {
                    foreach (var obj in folder.Items)
                    {
                        if (obj is Outlook.MailItem mailItem)
                        {
                            var newItem = mailItem.Forward();
                            //newItem.Recipients.Add(AddressTo);
                            newItem.To = AddressTo;
                            if (!String.IsNullOrWhiteSpace(AddressCc))
                            {
                                newItem.CC = AddressCc;
                            }
                            newItem.Body = ForwardedTxt + newItem.Body;
                            newItem.Importance = Outlook.OlImportance.olImportanceHigh;
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
                            newItem.Display(false);
                            //newItem.Save();
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
            if (folderEntryID == null || folderStoreID == null)
            {

                return null;
            }
            return 
                application.Session.GetFolderFromID(
                    folderEntryID, folderStoreID)
                    as Outlook.Folder;
        }

        internal void DisplayFolder()
        {
            Outlook.Folder folder = GetOutlookFolder();
            if (folder == null)
            {
                Status = "No Folder selected";
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
                Status = "Selected items: " + (folder?.Items?.Count ?? 0);
            }
            catch (Exception exc)
            {
                #if DEBUG
                Debug.WriteLine(exc.Message);
                #endif
                Status = exc.Message;
                folder = null;
            }
        }


        public const string AuthFwd = "Authorize";
        public string[] Areas => new string[] { AuthFwd, Business.Area1, Business.Area2, Business.Area3 };
        public string SelectedArea { get; set; } = AuthFwd;

    }
}
