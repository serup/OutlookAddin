	using System;
	using Extensibility;
	using System.Runtime.InteropServices;
  using System.Windows.Forms;                             //MessageBox
  using Office = Microsoft.Office.Core;                   //CommandBarButton, IRibbonControl
  using pstTransporter;                                   //For backup/Restore pst
  using Outlook = Microsoft.Office.Interop.Outlook;       //To avoid long fully qualified names
  using System.Reflection;                                //Assembly
  using System.IO;                                        //StreamReader  
  using System.Collections.Generic;                       //List
  using System.Diagnostics;                               //Debug

namespace pstTransporter
{
	#region Read me for Add-in installation and setup information.
	// When run, the Add-in wizard prepared the registry for the Add-in.
	// At a later time, if the Add-in becomes unavailable for reasons such as:
	//   1) You moved this project to a computer other than which is was originally created on.
	//   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
	//   3) Registry corruption.
	// you will need to re-register the Add-in by building the pstTransporterSetup project, 
	// right click the project in the Solution Explorer, then choose install.
	#endregion
	
	/// <summary>
	///   The object for implementing an Add-in.
	/// </summary>
	/// <seealso class='IDTExtensibility2' />
	[GuidAttribute("3C4B45CA-6E0B-4ED5-9120-AA6829DB4471"), ProgId("pstTransporter.Connect")]
  public class Connect : Object, Extensibility.IDTExtensibility2,  Office.IRibbonExtensibility 
	{
    #region Instance Variables

    private Outlook.Application m_OutlookAppObj;
    private object m_AddInInstance;

    //mine
    Outlook.Explorers m_Explorers;
    private static List<OutlookExplorer> m_Windows;        // List of tracked explorer windows. 
    private static Office.IRibbonUI m_Ribbon;              // Ribbon UI reference

    #endregion

    #region IDTExtensibility2 Members
    /// <summary>
    ///		Implements the constructor for the Add-in object.
    ///		Place your initialization code within this method.
    /// </summary>
    public Connect()
    {
    }

    /// <summary>
    ///      Implements the OnConnection method of the IDTExtensibility2 interface.
    ///      Receives notification that the Add-in is being loaded.
    /// </summary>
    /// <param term='application'>
    ///      Root object of the host application.
    /// </param>
    /// <param term='connectMode'>
    ///      Describes how the Add-in is being loaded.
    /// </param>
    /// <param term='addInInst'>
    ///      Object representing this Add-in.
    /// </param>
    /// <seealso class='IDTExtensibility2' />
    public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom)
    {
      m_OutlookAppObj = (Outlook.Application)application;
      m_AddInInstance = addInInst;

      // If we are not loaded upon startup, forward to OnStartupComplete() and pass the incoming System.Array.
      if (connectMode != ext_ConnectMode.ext_cm_Startup)
      {
        OnStartupComplete(ref custom);
      }
    }

    /// <summary>
    ///     Implements the OnDisconnection method of the IDTExtensibility2 interface.
    ///     Receives notification that the Add-in is being unloaded.
    /// </summary>
    /// <param term='disconnectMode'>
    ///      Describes how the Add-in is being unloaded.
    /// </param>
    /// <param term='custom'>
    ///      Array of parameters that are host application specific.
    /// </param>
    /// <seealso class='IDTExtensibility2' />
    public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
    {
      // proper implementation of this method should test the connection mode (this time for anything other than ext_DisconnectMode.ext_dm_HostShutdown) 
      // and forward the incoming System.Array to our implementation of OnBeginShutdown()
      if (disconnectMode != ext_DisconnectMode.ext_dm_HostShutdown)
      {
        OnBeginShutdown(ref custom);
      }

      m_OutlookAppObj = null;
    }

    /// <summary>
    ///      Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
    ///      Receives notification that the collection of Add-ins has changed.
    /// </summary>
    /// <param term='custom'>
    ///      Array of parameters that are host application specific.
    /// </param>
    /// <seealso class='IDTExtensibility2' />
    public void OnAddInsUpdate(ref System.Array custom)
    {
      // The OnAddInsUpdate() method is called if the end user inserts or removed Add-ins to the host (the Application.COMAddins property
      // can be used to obtain the current list at runtime). In the case that you need to perform any special processing when the end user adds
      // or removes a new Add-in, this would be the place to do so. The auto-generated implementation is currently empty, and can remain so.
    }

    /// <summary>
    ///      Implements the OnStartupComplete method of the IDTExtensibility2 interface.
    ///      Receives notification that the host application has completed loading.
    /// </summary>
    /// <param term='custom'>
    ///      Array of parameters that are host application specific.
    /// </param>
    /// <seealso class='IDTExtensibility2' />
    public void OnStartupComplete(ref System.Array custom)
    {
      // This method is called after the host application has completed loading. At this point, all host resources are available for use 
      // by the Add-in. This is an ideal place to construct the UI of your Add-in types, as you can safely obtain the set of Explorers and Inspectors.

      m_Explorers = m_OutlookAppObj.Explorers;
      m_Windows = new List<OutlookExplorer>();

      // Wire up event handlers to handle multiple Explorer windows
      m_Explorers.NewExplorer += new Outlook.ExplorersEvents_NewExplorerEventHandler(m_Explorers_NewExplorer);

      // Add the ActiveExplorer to m_Windows
      Outlook.Explorer expl = m_OutlookAppObj.ActiveExplorer() as Outlook.Explorer;
      OutlookExplorer window = new OutlookExplorer(expl);   
      m_Windows.Add(window);

      // Hook up event handlers for window
      window.Close += new EventHandler(WrappedWindow_Close);
      window.InvalidateControl += new EventHandler<OutlookExplorer.InvalidateEventArgs>(WrappedWindow_InvalidateControl);
    }

    /// <summary>
    ///      Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
    ///      Receives notification that the host application is being unloaded.
    /// </summary>
    /// <param term='custom'>
    ///      Array of parameters that are host application specific.
    /// </param>
    /// <seealso class='IDTExtensibility2' />
    public void OnBeginShutdown(ref System.Array custom)
    {
      // OnBeginShutdown() indicates that the host is in the process of shutting down (just before the call to the OnDisconnection() method).
      // At this point you still have access to the host application, so this is an ideal place to remove any UI widgets you plugged 
      // into the active explorer.

      // Get set of command bars on active explorer.
      Office.CommandBars commandBars = m_OutlookAppObj.ActiveExplorer().CommandBars;
      try
      {
        // Find our button and kill it.
        commandBars["Standard"].Controls["Do Something!"].Delete(System.Reflection.Missing.Value);
      }
      catch (System.Exception ex)
      {
        MessageBox.Show(ex.Message);
      }
    } 
    #endregion

    #region IRibbonExtensibility Members - ribbon callbacks

    public string GetCustomUI(string ribbonID)
    {
      //return Properties.Resources.CustomUI;

      string customUI = string.Empty;
      Debug.WriteLine(ribbonID);
      
      // Return the appropriate XML markup for ribbonID.
      switch (ribbonID)
      {
        case "Microsoft.Outlook.Explorer":
          customUI = GetResourceText(Consts.CustomUiDefinitionXML);
          return customUI;

        default:
          return string.Empty;
      }
    }

    public void Ribbon_Load(Office.IRibbonUI ribbonUI)
    {
      Connect.m_Ribbon = ribbonUI;
    }

    public bool ShouldCustomItemBeVisible(Office.IRibbonControl control)
    {      
      Outlook.Folder folder = control.Context as Outlook.Folder;

      if (folder.DefaultItemType == Outlook.OlItemType.olContactItem)
      {
        return true; 
      }

      return false;
    }

    // OnMyButtonClick routine handles all button click events and displays IRibbonControl.Context in message box
    public void OnMyButtonClick(Office.IRibbonControl control)
    {
      if (control.Context is Outlook.Folder)
      {
        Outlook.Folder folder = control.Context as Outlook.Folder;

        //do something ... start contact cleaning .. or whateva
        string pstPath = Consts.BackupDir + Consts.BackupFileName;

        PST myPST = new PST(pstPath, m_OutlookAppObj);
        //myPST.Backup(folder.FolderPath);
        //MessageBox.Show("Hopefully I did not fuck up! Keep faith ;)");
        //myPST.Restore(folder.FolderPath);
        //myPST.Restore(@"\\Bilal.Taherkheli@nokia.com\Contacts\TestFold");


        ContactHandler myCH = new ContactHandler(m_OutlookAppObj, folder);
        myCH.ProcessContacts();
      }
    }
  
    #endregion

    #region Event Handlers

    /// <summary>
    /// The NewExplorer event fires whenever a new Explorer is displayed. 
    /// </summary>
    /// <param name="Explorer"></param>
    private void m_Explorers_NewExplorer(Outlook.Explorer Explorer)
    {
      try
      {
        // Check to see if this is a new window we don't already track
        OutlookExplorer existingWindow = FindOutlookExplorer(Explorer);

        //If the m_Windows collection does not have a window for this Explorer, we should add it to m_Windows
        if (existingWindow == null)
        {
          OutlookExplorer window = new OutlookExplorer(Explorer);
          window.Close += new EventHandler(WrappedWindow_Close);
          window.InvalidateControl += new EventHandler<OutlookExplorer.InvalidateEventArgs>(WrappedWindow_InvalidateControl);
          m_Windows.Add(window);
        }
      }
      catch (System.Exception ex)
      {
        Debug.WriteLine(ex.Message);
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    void WrappedWindow_InvalidateControl(object sender, OutlookExplorer.InvalidateEventArgs e)
    {
      if (m_Ribbon != null)
      {
        m_Ribbon.InvalidateControl(e.ControlID);
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    void WrappedWindow_Close(object sender, EventArgs e)
    {
      OutlookExplorer window = (OutlookExplorer)sender;
      window.Close -= new EventHandler(WrappedWindow_Close);
      m_Windows.Remove(window);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    /// Looks up the window wrapper for a given window object
    /// </summary>
    /// <param name="window">An outlook explorer window</param>
    /// <returns></returns>
    internal static OutlookExplorer FindOutlookExplorer(object window)
    {
      foreach (OutlookExplorer Explorer in m_Windows)
      {
        if (Explorer.Window == window)
        {
          return Explorer;
        }
      }
      return null;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="resourceName"></param>
    /// <returns></returns>
    private static string GetResourceText(string resourceName)
    {
      Assembly asm = Assembly.GetExecutingAssembly();
      string[] resourceNames = asm.GetManifestResourceNames();
      for (int i = 0; i < resourceNames.Length; ++i)
      {
        if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
        {
          using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
          {
            if (resourceReader != null)
            {
              return resourceReader.ReadToEnd();
            }
          }
        }
      }
      return null;
    } 

    #endregion
  }
}