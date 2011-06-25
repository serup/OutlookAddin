using System;

namespace pstTransporter
{
  public class Consts
  {
    public static string BackupFolderInsidePST = "BackupFolder"; //must have a name different from TargetContactFolder or else Outlook messes up
    public static string BackupDir = @"C:\Users\taherkhe\Documents\Visual Studio 2008\Projects\HobbyProjects\pstTransporterSln\xDUMP\";
    public static string BackupFileName = "xbackup.pst";
    //public static string ExplorerRibbonID = "Microsoft.Outlook.Explorer";
    public static string CustomUiDefinitionXML = "pstTransporter.UI.CustomUI.xml";
  }
}
