using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;               // MessageBox
using System.IO;                          // File operations

namespace pstTransporter
{
  internal class PST
  {
    private String pstFile = null;
    private Outlook.Application ownerApp = null;

    #region Constructor
    public PST(String name, Outlook.Application owner)  //todo: consider changing the signature so pstfile path is not needed to be passed. Instead it must be dynamically retrieved
    {
      pstFile = name;
      ownerApp = owner;
    } 
    #endregion

    /// <summary>
    /// 
    /// </summary>
    /// <param name="folderPath"></param>
    public void Backup(string folderPath)
    {
      Outlook.Store backUpStore = null;
      Outlook.ContactItem tempItem = null;
      Outlook.Folder dstFolder = null;
      Outlook.Folder srcFolder = GetFolder(folderPath);

      try
      {
        if (srcFolder != null)
        {
          ownerApp.Session.AddStore(pstFile);   //add store. If it does not exist, it will be created
          backUpStore = GetStore();   //get the store to confirm it has been successfully attached to current session

          if (backUpStore == null)
            throw new ApplicationException("Outlook Store (.pst) -> '" + pstFile + "' could not be attached to current Outlook session!");
          else
          {
            dstFolder = GetFolder(Path.Combine(backUpStore.GetRootFolder().FolderPath, Consts.BackupFolderInsidePST)); //try to locate dstFolder in -pst

            if (dstFolder != null)   //dstFolder exists already, empty it otherwise create a new one
            {
              //todo: (hard) deleting the folder is tricky. Not possible in .NET. Would need CDO1.21 or extended MAPI(C++ or Delphi only) 
              while (dstFolder.Items.Count > 0)
                dstFolder.Items.Remove(1);        
            }
            else
            {
              backUpStore.GetRootFolder().Folders.Add(Consts.BackupFolderInsidePST, Outlook.OlDefaultFolders.olFolderContacts);   //System.Reflection.Missing.Value
            }

            //todo: this copying part is not fully stabe... investigate if its due to some garbage collection issue
            //though after foced garbage collection at the end of each loop iteration, the frequency of 'operation failed' reduced significantly from 1 in 4 to about 1 in 25 
            //copy contacts one by one to appropriate folder in added store
            foreach (object obj in srcFolder.Items)
            {
              Outlook.ContactItem item = obj as Outlook.ContactItem;

              if (item != null)  //Conversion from 'object' to 'ContactItem' succeeded
              {
                //create copy of item
                tempItem = (Outlook.ContactItem)item.Copy();
                //move this copy to the backup folder in the attached Store 
                tempItem.Move(backUpStore.GetRootFolder().Folders[Consts.BackupFolderInsidePST]);
                tempItem = null;  //todo: to make it appear stable ... use a try catch and some kinda counter to keep trying 
                //unless no exception is thrown ... be sure to roll back the mess first
                GC.Collect();
                GC.WaitForPendingFinalizers();
              }
            }
          }
        }
      }
      catch (System.Exception e)
      {
        MessageBox.Show(e.Message);
      }

      finally
      {
        if (backUpStore != null)   //remove store i.e. detach store from current profile
          ownerApp.Session.RemoveStore(backUpStore.GetRootFolder());

        //ensure cleanup
        backUpStore = null;
        tempItem = null;
        dstFolder = null;
        srcFolder = null;
        GC.Collect();
        GC.WaitForPendingFinalizers();
      }
    }

    /// <summary>
    /// Restores the previously stored contact folder from a .pst (stored by the name defined by Consts.BackupFolderInsidePST)
    /// to the contact folder in Outlook upon which it was invoked 
    /// - there must be a targetFolder with the right name in Outlook (if invoked programmatically, it should check if the folder it is operating upon exists)
    /// - there must be a valid .pst at the right location (no user involvement. Must be dynamically generated path in some Outlook folder)
    /// - there must be a folder inside .pst wth the right name (Consts.BackupFolderInsidePST)
    /// 
    /// Please noe that this method just restores the content in backup folder in .pst to the folder on whic it was called. If the user backups A and invokes 
    /// restore() on B, restore() just acts as requested -> B data overwritten with A Basically always take backup first and then invoke restore()
    /// TEST INFO:
    /// tested 3 error scenarios
    /// 1. target folder not found in Outlook. Nothing happens. Graceful exit with info! :)
    /// 2. .pst not found in expected file system location. Nothing happens. Graceful exit with info! :)
    /// 3. target folder and .pst both exist but expected backup fodler not found in .pst. Graceful exit with info! :)
    /// </summary>
    /// <param name="targetFolderPath"> Full path of the Outlook Contacts folder to operate upon</param>
    public void Restore(string targetFolderPath)
    {
      string targetFolderName = null;
      Outlook.Folder targetFolderRoot = null;
      Outlook.Store backUpStore = null;
      Outlook.Folder backUpFolder = null;
      Outlook.Folder targetFolder = GetFolder(targetFolderPath);

       try
      {
        if ((targetFolder != null) && (File.Exists(pstFile)))   // the folder to restore to and the .pst to restore from, both exist
        {
          ownerApp.Session.AddStore(pstFile);
          backUpStore = GetStore();   //get the store to confirm it has been successfully attached to current session

          if (backUpStore == null)
            throw new ApplicationException("Outlook Store (.pst) -> '" + pstFile + "' could not be attached to current Outlook session!");
          else
          {
            backUpFolder = GetFolder(Path.Combine(backUpStore.GetRootFolder().FullFolderPath, Consts.BackupFolderInsidePST));

            if (backUpFolder == null)
              throw new ApplicationException("Backup folder inside (.pst) called -> '" + Consts.BackupFolderInsidePST + "' not found!");
            else
            {
              targetFolderName = targetFolder.Name;
              targetFolderRoot = (Outlook.Folder)targetFolder.Parent;
              targetFolder.Delete();
              targetFolder = (Outlook.Folder)backUpFolder.CopyTo(targetFolderRoot);
              targetFolder.Name = targetFolderName;    // rename   
              ownerApp.ActiveExplorer().SelectFolder(targetFolder);  //set focus to the folder that was operated upon   todo: delete .. as it is just for demo purpose
            }
          }
        }
        else
        {
          throw new ApplicationException("Backup not possible because either the target folder, or the .pst to retrieve from, was not found");          
        }
      }

      catch (System.Exception e)
      {
        MessageBox.Show(e.Message);
      }

      finally
      {
        //cleanup
        if (backUpStore != null)
          ownerApp.Session.RemoveStore(backUpStore.GetRootFolder());   // detach store from current profile

        targetFolderName = null;
        targetFolderRoot = null;
        backUpStore = null;
        backUpFolder = null;
        targetFolder = null;        
        GC.Collect();
        GC.WaitForPendingFinalizers();
      }      
    }

    /// <summary>
    /// This procedure takes the full path you send, in the form Mailbox—Your Name\Inbox\Customers\Archive, and splits the path into individual folder names.
    /// It attempts to locate the root folder, Mailbox—Your Name, and, if it succeeds, loops through all the remaining folder names, each time retrieving a 
    /// reference to the corresponding folder, until it runs out of names. If it successfully reaches the last folder name, it returns the corresponding Outlook.Folder object.
    /// </summary>
    /// <param name="folderPath"></param>
    /// <returns></returns>
    private Outlook.Folder GetFolder(string folderPath)
    {
      Outlook.Folder returnFolder = null;

      try
      {
        // Remove leading "\" characters.
        folderPath = folderPath.TrimStart("\\".ToCharArray());

        // Split the folder path into individual folder names.
        String[] folders = folderPath.Split("\\".ToCharArray());

        // Retrieve a reference to the root folder.
        returnFolder = ownerApp.Session.Folders[folders[0]] as Outlook.Folder;

        // If the root folder exists, look in subfolders.
        if (returnFolder != null)
        {
          Outlook.Folders subFolders = null;
          String folderName;

          // Look through folder names, skipping the first folder, which you already retrieved.
          for (int i = 1; i < folders.Length; i++)
          {
            folderName = folders[i];
            subFolders = returnFolder.Folders;
            returnFolder = subFolders[folderName] as Outlook.Folder;
          }
        }
      }
      catch
      {
        // Do nothing at all -- just return a null reference.
        returnFolder = null;
      }

      return returnFolder;
    }

    /// <summary>
    /// Looks for a store by comparing filepath attribute/preperty
    /// </summary>
    /// <returns>the sought after Store if found, null otherwise </returns>
    private Outlook.Store GetStore()
    {
      Outlook.Store storeToReturn = null;

      foreach (Outlook.Store store in ownerApp.Session.Stores)
      {
        if (store.FilePath == pstFile)
        {
          storeToReturn = store;
          break;
        }
      }

      return storeToReturn; 
    }

    ///// <summary>
    ///// blah blah
    ///// </summary>
    ///// <param name="folder"></param>
    //private void PermanentlyDeleteAllItems(Outlook.Folder folder)
    //{
    //  string entryID = null;
    //  string storeID = folder.StoreID;
    //  Outlook.ContactItem delItem = null;

    //  foreach (object obj in folder.Items)
    //  {
    //    Outlook.ContactItem item = obj as Outlook.ContactItem;
    //    entryID = item.EntryID;  // Store item entry id
    //    item.Delete();           // Send to deleted items folder
    //    delItem = (Outlook.ContactItem)ownerApp.Session.GetItemFromID(entryID, storeID);
    //    delItem.Delete();
    //    item = null;
    //    delItem = null;
    //  }
    //}
    
    ~PST()
    {
      pstFile = null;
      ownerApp = null;
    }

    #region to remove any unnecessary stores created during debugging
    //MAPIFolder tempFolder = null;
    //try
    //{
    //  foreach (Store store in ownerApp.Session.Stores)
    //  {
    //    if (store.IsDataFileStore)
    //    {
    //      if (store.DisplayName == "Outlook Data File")
    //      {
    //        tempFolder = store.GetRootFolder();
    //      }
    //    }
    //  }
    //}

    //catch (System.Exception e)
    //{
    //  MessageBox.Show(e.Message);
    //}

    ////remove unnecessarily created store
    //try
    //{

    //  if (tempFolder != null)
    //  {
    //    ownerApp.Session.RemoveStore(tempFolder);
    //  }
    //}
    //catch (System.Exception e)
    //{
    //  MessageBox.Show(e.Message);
    //} 
    #endregion
  }
}
