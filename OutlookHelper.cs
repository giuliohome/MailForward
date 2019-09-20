using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Configuration;
using System.IO;

namespace MailForward
{
    internal class OutlookHelper
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
                    var lines = new List<string>();
                    while (!sr.EndOfStream)
                    {
                        string line = await sr.ReadLineAsync();
                        if (String.IsNullOrWhiteSpace(line)) continue;
                        string[] fields = line.Split('\t');
                        if (fields.Length != 2) continue;
                        switch (fields[0])
                        {
                            case CsvAddressTo:
                                AddressTo = fields[1];
                                continue;
                            case CsvAddressCc:
                                AddressCc = fields[1];
                                continue;
                            case CsvForwardedTxt:
                                lines.Add(fields[1]);
                                continue;
                            default:
                                continue;
                        }
                    }
                    ForwardedTxt = String.Join("\n", lines);
                }
            }

        }
        private const string CsvAddressTo = "Address To";
        private const string CsvAddressCc = "Address Cc";
        private const string CsvForwardedTxt = "Forwarded Text";
        internal async Task SaveConfig()
        {
            using (var sw = new StreamWriter(ConfigurationManager.AppSettings["settingsPath"], false))
            {
                await sw.WriteLineAsync($"{CsvAddressTo}\t{AddressTo.Replace("\t"," ")}");
                if (!String.IsNullOrWhiteSpace(AddressCc))
                {
                    await sw.WriteLineAsync($"{CsvAddressCc}\t{AddressCc.Replace("\t", " ")}");
                }
                if (!String.IsNullOrWhiteSpace(ForwardedTxt))
                {
                    foreach (string line in ForwardedTxt.Replace("\t", " ").Split('\n'))
                    {
                        await sw.WriteLineAsync($"{CsvForwardedTxt}\t{line}");
                    } 
                    
                }
                
            }

        }

        private Outlook.Folder selectedFolder = null;
        internal async Task DisplayFolder()
        {
            await Task.Run(() => DisplayFolder(selectedFolder));
            DisplayFolder(selectedFolder);
        }
        internal async Task<string> SelectFolder()
        {
            selectedFolder = await Task.Run(() => application.Session.PickFolder() as Outlook.Folder);
            return selectedFolder?.FolderPath?? "Not Selected";
        }

        internal int? GetItemNumber()
        {
            return selectedFolder?.Items?.Count ?? 0;
        }
        public string AddressTo { get; set; } = ""; // Mbx Some.Grouo <Some.Group@xxx.com>; Name Surname <Name.Surname@xxx.com>;
        public string AddressCc { get; set; } = "";

        public string ForwardedTxt { get; set; } = @"

Request authorized, please follow up.


";
        const string PR_SMTP_ADDRESS =
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
        internal async Task ForwardItems()
        {
            await Task.Run( () =>
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
                        newItem.Display(false);
                        //newItem.Save();
                    }
                }
            }
            );
        }

        private void DisplayFolder(Outlook.Folder folder)
        {
            if (folder == null)
            {
                return;
            }
            try
            {
                Outlook.Folder folderFromID =
                    application.Session.GetFolderFromID(
                    folder.EntryID, folder.StoreID)
                    as Outlook.Folder;
                application.ActiveExplorer().CurrentFolder = folder;
                //folderFromID.Display();
            }
            catch (Exception exc)
            {
                Debug.WriteLine(exc.Message);
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
                Debug.WriteLine("There is no folder named " + folderName +
                    ".", "Find Folder Name");
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

    }
}
