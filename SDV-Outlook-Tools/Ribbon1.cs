using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace SDV_Outlook_Tools
{
    public partial class Ribbon1
    {
        [System.ComponentModel.Browsable(false)]
        public Microsoft.Office.Tools.Ribbon.RibbonComboBox SelectedItem { get; set; }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                dp_Mailstatus.SelectedItemIndex =0;
                dp_Mailalter.SelectedItemIndex = 2;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Fehler im Programmablauf", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_RemoveAttachments_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string Mailstatus = dp_Mailstatus.SelectedItem.ToString();
                int Mailalter = Convert.ToInt32(dp_Mailalter.SelectedItem.ToString());
                RemoveAttachments(Mailstatus, Mailalter);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Fehler im Programmablauf", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void btn_MoveAttachments_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string Mailstatus = dp_Mailstatus.SelectedItem.ToString();
                int Mailalter = Convert.ToInt32(dp_Mailalter.SelectedItem.ToString());
                MoveAttachments(Mailstatus, Mailalter);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Fehler im Programmablauf", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void RemoveAttachments(string Mailstatus,int Mailalter)
        {
            try
            {
                DialogResult result1 = MessageBox.Show("Möchten Sie wirklich die Änhange aller "+ Mailstatus + " Mails älterer als " + Mailalter.ToString() + " Tage entfernen?", "Entfernen der Änhänge", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result1 == DialogResult.Yes)
                {
                    string pathToSave = null;
                    if (pathToSave != "0")
                    {
                        EnumerateFoldersInDefaultStore(pathToSave, Mailstatus, Mailalter);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Fehler im Programmablauf", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void MoveAttachments(string Mailstatus, int Mailalter)
        {
            try
            {
                DialogResult result1 = MessageBox.Show("Möchten Sie wirklich die Änhange aller " + Mailstatus + " Mails älterer als " + Mailalter.ToString() + " Tage verschieben?", "Verschieben der Änhänge", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result1 == DialogResult.Yes)
                {
                    string pathToSave = getSaveFolder();
                    EnumerateFoldersInDefaultStore(pathToSave, Mailstatus, Mailalter);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Fehler im Programmablauf", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static string getSaveFolder()
        {
            try
            {
                FolderBrowserDialog folderDlg = new FolderBrowserDialog
                {
                    ShowNewFolderButton = true
                };
                DialogResult result = folderDlg.ShowDialog();
                if (result == DialogResult.OK)
                {
                    return folderDlg.SelectedPath;
                }
                else
                {
                    return "0";
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
                return "0";
            }
        }

        static void EnumerateFoldersInDefaultStore(string pathToSaveFile, string Mailstatus, int Mailalter)
        {
            Microsoft.Office.Interop.Outlook.Application Application = new Microsoft.Office.Interop.Outlook.Application();
        //    Microsoft.Office.Interop.Outlook.Folder root = Application.Session.DefaultStore.GetRootFolder() as Microsoft.Office.Interop.Outlook.Folder;
            Microsoft.Office.Interop.Outlook.Folder folder = Application.Session.PickFolder() as Microsoft.Office.Interop.Outlook.Folder;
            Microsoft.Office.Interop.Outlook.Folder folderFromID = Application.Session.GetFolderFromID(folder.EntryID, folder.StoreID) as Microsoft.Office.Interop.Outlook.Folder;
           EnumerateFolders(folderFromID, pathToSaveFile, Mailstatus, Mailalter);

         //   EnumerateFolders(root, pathToSaveFile, Mailstatus, Mailalter);
        }

        static void EnumerateFolders(Microsoft.Office.Interop.Outlook.Folder folder, string pathToSaveFile, string Mailstatus, int Mailalter)
        {
            Microsoft.Office.Interop.Outlook.Folders childFolders = folder.Folders;
            if (childFolders.Count > 0)
            {
                IterateMessages(folder, pathToSaveFile, Mailstatus, Mailalter);
                foreach (Microsoft.Office.Interop.Outlook.Folder childFolder in childFolders)
                {
                    // We only want Inbox folders - ignore Contacts and others
                    if (childFolder.FolderPath.Contains("Posteingang"))
                    {
                        EnumerateFolders(childFolder, pathToSaveFile, Mailstatus, Mailalter);

                    }
                }

            }
            IterateMessages(folder, pathToSaveFile, Mailstatus, Mailalter);
        }

        static void IterateMessages(Microsoft.Office.Interop.Outlook.Folder folder, string basePath, string Mailstatus, int Mailalter)
        {
            var fi=folder.Items;
            if (Mailstatus== "ungelesene")
            {
                fi=folder.Items.Restrict("[Unread] = true");
            }
            else if (Mailstatus == "gelesene")
            {
                fi = folder.Items.Restrict("[Unread] = false" );
            }
      
            if (fi != null)
            {
                foreach (Object item in fi)
                {
                    Microsoft.Office.Interop.Outlook.MailItem mi = (Microsoft.Office.Interop.Outlook.MailItem)item;
                    var attachments = mi.Attachments;
                        if (attachments.Count != 0)
                        {
                            // Create a directory to store the attachment 
                            if (basePath != null)
                            {
                                if (!Directory.Exists(basePath + folder.FolderPath))
                                {
                                    Directory.CreateDirectory(basePath + folder.FolderPath);
                                }
                            }
                                if (mi.ReceivedTime < DateTime.Now.AddDays(-Mailalter))
                                {
                                    if (mi.SenderEmailAddress != "wiki@sdv.de")
                                    {
                                    int AttachmentsCount = mi.Attachments.Count;
                                        for (int i = 1; i <= AttachmentsCount; i++)
                                        {
                                            if (basePath != null)
                                            {
                                                var fn = mi.Attachments[1].FileName.ToLower();
                                            // Create a further sub-folder for the sender
                                                if (!Directory.Exists(basePath + folder.FolderPath + @"\" + mi.Sender.Address))
                                                {
                                                    Directory.CreateDirectory(basePath + folder.FolderPath + @"\" + mi.Sender.Address);
                                                }
                                                if (!File.Exists(basePath + folder.FolderPath + @"\" + mi.Sender.Address + @"\" + mi.Attachments[1].FileName))
                                                {
                                                    mi.Attachments[1].SaveAsFile(basePath + folder.FolderPath + @"\" + mi.Sender.Address + @"\" + mi.Attachments[1].FileName);
                                                    mi.Body = mi.Body + "Anhange nach " + basePath + folder.FolderPath + @"\" + mi.Sender.Address + @"\" + mi.Attachments[1].FileName + " durch " + Environment.UserName + " verschoben.";
                                                    mi.Attachments[1].Delete();
                                                    mi.Save();
                                                }
                                                else
                                                {
                                                }
                                            }
                                            else
                                            {
                                                mi.Body = mi.Body + "Anhange " + mi.Attachments[1].FileName + " durch " + Environment.UserName + " gelöscht.";
                                                mi.Attachments[1].Delete();
                                                mi.Save();
                                            }
                                        }
                                    }
                                }

                            }
                }
            }
        }

    }
}
