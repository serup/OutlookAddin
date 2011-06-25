using System;
using System.Collections.Generic;
using System.Windows.Forms;                             //MessageBox
using Outlook = Microsoft.Office.Interop.Outlook;

namespace pstTransporter
{
  internal class ContactHandler
  {
    private Outlook.Application outlookApp;
    private Outlook.Folder folder;
    private List<object> nonContacts;  //anything that is not a contact item but found in the received fodler anyway
    private List<string> itemsToDelete;
    private List<string> itemsToMerge; 
    private List<string> itemsUnique;

    // Idea:
    // Remove genuine duplicates
    // Store unique contacts separately and remove them from the list under processing
    // auto-merge to the extent possible i.e. merge when no conflicts present and add the resultant to a sepaarte list
    // conflict resolution for the remaining problem children
    // add up all partial lists and create a new folder or whatever

    #region Constructor
    public ContactHandler(Outlook.Application owner, Outlook.Folder contactFolder)  
    {
      outlookApp = owner;
      folder = contactFolder;
      nonContacts = new List<object>();
      itemsToDelete = new List<string>();
      itemsToMerge = new List<string>();
      itemsUnique = new List<string>();
    } 
    #endregion

    #region methods	

    /// <summary>
    /// 
    /// </summary> 
    public void ProcessContacts()
    {
      folder.Items.Sort("[FirstName]", false);  //sort by 'first name'
      DeleteDuplicatesAndIdentifyUniquesAndMergeCandidates();

    }

    private void DeleteDuplicatesAndIdentifyUniquesAndMergeCandidates()
    {
      int index = 1;
      int iterationNo = 1;

      foreach (object obj in folder.Items)
      {
        index = iterationNo + 1;
        Outlook.ContactItem item = obj as Outlook.ContactItem;     //The 'as' operator is like a cast except that it yields null on conversion failure instead of raising an exception

        if (item == null) //Conversion from 'object' to 'ContactItem' failed
        {
          nonContacts.Add(obj);     //todo: what else? what to do with these if some are found somehow?
        }
        else
        {
          while (index < (folder.Items.Count + 1))
          {
            Outlook.ContactItem temp = (Outlook.ContactItem)folder.Items[index.ToString()];

            //condition to check for DUPLICATES to be deleted right away
            if (                                        
                 #region condition
		                (item.FirstName == temp.FirstName) &&   
                    (item.MobileTelephoneNumber == temp.MobileTelephoneNumber) &&
                    (item.Email1Address == temp.Email1Address) 
	               #endregion
               )
            {
              if (!(itemsToDelete.Contains(temp.EntryID)))
                itemsToDelete.Add(temp.EntryID);
            }
            
            //condition to check for MERGE CANDIDATES to be processed further
            else if (                                       
                      #region condition
                      ((item.FirstName != null) && (item.FirstName == temp.FirstName)) ||
                      ((item.MobileTelephoneNumber != null) && (item.MobileTelephoneNumber == temp.MobileTelephoneNumber)) ||
                      ((item.Email1Address != null) && (item.Email1Address == temp.Email1Address))
                      #endregion                    
                    )
            {
              //if not in itemsToMerge already, add both elements
              if ((!(itemsToDelete.Contains(item.EntryID))) && (!(itemsToMerge.Contains(item.EntryID))))    //register only if 'item' is not marked for deletion not added as a 'to be merged' already
                itemsToMerge.Add(item.EntryID);

              if ((!(itemsToDelete.Contains(temp.EntryID))) && ((!(itemsToMerge.Contains(temp.EntryID)))))   //register only if 'temp' is not marked for deletion and not added as a 'to be merged' already
                itemsToMerge.Add(temp.EntryID);            
            }

            //possibly a unique item. Could be originally unique or transformed into a unique element after duplicates identification
            else
            {
              //it is the last iteration and this item, thus far, has neither been identified as a 'to be merged' nor as a 'to be deleted'
              if ((index == folder.Items.Count) && (!(itemsToDelete.Contains(item.EntryID))) && (!(itemsToMerge.Contains(item.EntryID)))) 
                itemsUnique.Add(item.EntryID);
            }

            index++;
          }

          //last item in the loop
          if ((index == folder.Items.Count) && (!(itemsToDelete.Contains(item.EntryID))) && (!(itemsToMerge.Contains(item.EntryID)))) 
            itemsUnique.Add(item.EntryID);
        }

        iterationNo++;
      }

      //**** test results
      string todelete = String.Empty;
      string unique = String.Empty;
      string tomerge = String.Empty;

      foreach (string entryID in itemsToDelete)
        todelete += GetContactFromEntryID(entryID).FirstName + ", ";

      foreach (string entryID in itemsUnique)
        unique += GetContactFromEntryID(entryID).FirstName + ", ";

      foreach (string entryID in itemsToMerge)
        tomerge += GetContactFromEntryID(entryID).FirstName + ", ";

      MessageBox.Show("\tItems to be deleted = " + itemsToDelete.Count + "\n\t" + todelete +
                      "\n\tunique items = " + itemsUnique.Count + "\n\t" + unique +
                      "\n\tto be merged = " + itemsToMerge.Count + "\n\t" + tomerge);    

      //****
      //DeleteItemsToDelete();
    }


    //****testing only
    private Outlook.ContactItem GetContactFromEntryID(string entryID)
    {
      foreach (object obj in folder.Items)
      {
        Outlook.ContactItem item = obj as Outlook.ContactItem;
        if (item.EntryID == entryID)
          return item;
      }

      return null;
    }

    //*****

    /// <summary>
    /// Take list of entryID's one at a time, look for a matching entryID in the contact items list, delete the matching item and break the loop
    /// </summary>
    private void DeleteItemsToDelete()
    {
      foreach (string entryID in itemsToDelete)
      {
        foreach (object obj in folder.Items)
        {
          Outlook.ContactItem item = obj as Outlook.ContactItem;
          if (item.EntryID == entryID)
          {
            //item.Move(outlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems)); //todo: does not work. item stays in deleted items. how to HARD delete??? 
            item.Delete();
            break;
          }
        }
      }

      itemsToDelete.Clear();
    }

    #endregion
  }
}               