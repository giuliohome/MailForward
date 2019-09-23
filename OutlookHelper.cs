﻿using System;
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

        private Outlook.Folder selectedFolder = null;
        internal async Task<string> SelectFolder()
        {
            selectedFolder = await Task.Run(() => application.Session.PickFolder() as Outlook.Folder);
            return selectedFolder?.FolderPath ?? "Not Selected";
        }

        internal int? GetItemNumber()
        {
            return selectedFolder?.Items?.Count;
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
            await Task.Run(() =>
           {
               if (GetItemNumber() == 0)
               {
                   return;
               }
               foreach (var obj in selectedFolder.Items)
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
        }

        internal void DisplayFolder()
        {
            Outlook.Folder folder = selectedFolder;
            if (folder == null)
            {
                Status = "No Folder selected";
                return;
            }
            try
            {
                Outlook.Folder folderFromID =
                    application.Session.GetFolderFromID(
                    folder.EntryID, folder.StoreID)
                    as Outlook.Folder;
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

        private void SetCurrentFolder(string folderName)
        {
            Outlook.Folder inBox = (Outlook.Folder)
                application.ActiveExplorer().Session.GetDefaultFolder
                (Outlook.OlDefaultFolders.olFolderInbox);
            try
            {
                application.ActiveExplorer().CurrentFolder = inBox.
                    Folders[folderName];
                application.ActiveExplorer().CurrentFolder.Display();
            }
            catch
            {
                #if DEBUG
                Debug.WriteLine("There is no folder named " + folderName +
                    ".", "Find Folder Name");
                #endif
                Status = "There is no folder named " + folderName;
            }
        }

        private string ShowFolderInfo()
        {
            Outlook.Folder folder =
                application.Session.PickFolder()
                as Outlook.Folder;
            if (folder != null)
            {
                //StringBuilder sb = new StringBuilder();
                //sb.AppendLine("Folder EntryID:");
                //sb.AppendLine(folder.EntryID);
                //sb.AppendLine();
                //sb.AppendLine("Folder StoreID:");
                //sb.AppendLine(folder.StoreID);
                //sb.AppendLine();
                //sb.AppendLine("Unread Item Count: "
                //    + folder.UnReadItemCount);
                //sb.AppendLine("Default MessageClass: "
                //    + folder.DefaultMessageClass);
                //sb.AppendLine("Current View: "
                //    + folder.CurrentView.Name);
                //sb.AppendLine("Folder Path: "
                //    + folder.FolderPath);
                //Debug.WriteLine(sb.ToString());
                Outlook.Folder folderFromID =
                    application.Session.GetFolderFromID(
                    folder.EntryID, folder.StoreID)
                    as Outlook.Folder;
                folderFromID.Display();
                return folder.FolderPath;
            }
            return "Not Selected";
        }


        public string[] Areas => new string[] { "Business1", "Area2", "Biz3", "Company4" };
        public string SelectedArea { get; set; } = AuthFwd;

    }
}