using System;
using System.Runtime.InteropServices;
using Inventor;
using Microsoft.Win32;
using System.Windows.Interop;

namespace MyAddInWithWpf
{
  /// <summary>
  /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
  /// that all Inventor AddIns are required to implement. The communication between Inventor and
  /// the AddIn is via the methods on this interface.
  /// </summary>
  [GuidAttribute("880b4435-f9c2-4c5e-b234-d543a20b5c36")]
  public class StandardAddInServer : Inventor.ApplicationAddInServer
  {

    // Inventor application object.
    private Inventor.Application m_inventorApplication;
    private Inventor.ButtonDefinition m_btnDef;

    public StandardAddInServer()
    {
    }

    #region ApplicationAddInServer Members

    public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
    {
      // This method is called by Inventor when it loads the addin.
      // The AddInSiteObject provides access to the Inventor Application object.
      // The FirstTime flag indicates if the addin is loaded for the first time.

      // Initialize AddIn members.
      m_inventorApplication = addInSiteObject.Application;

      // TODO: Add ApplicationAddInServer.Activate implementation.
      // e.g. event initialization, command creation etc.
      var cmdMgr = m_inventorApplication.CommandManager;
      m_btnDef = cmdMgr.ControlDefinitions.AddButtonDefinition(
        "ShowWpfDialog", "ShowWpfDialog", CommandTypesEnum.kQueryOnlyCmdType);
      m_btnDef.OnExecute += ctrlDef_OnExecute;
      m_btnDef.AutoAddToGUI();
    }

    void ctrlDef_OnExecute(NameValueMap Context)
    {
      var wpfWindow = new InvAddIn.MyWpfWindow();

      // Could be a good idea to set the owner for this window
      // especially if it was modeless as mentioned in this article:
      var helper = new WindowInteropHelper(wpfWindow);
      helper.Owner = new IntPtr(m_inventorApplication.MainFrameHWND);

      wpfWindow.ShowDialog();
    }

    public void Deactivate()
    {
      // This method is called by Inventor when the AddIn is unloaded.
      // The AddIn will be unloaded either manually by the user or
      // when the Inventor session is terminated

      // TODO: Add ApplicationAddInServer.Deactivate implementation

      // Release objects.
      m_inventorApplication = null;

      GC.Collect();
      GC.WaitForPendingFinalizers();
    }

    public void ExecuteCommand(int commandID)
    {
      // Note:this method is now obsolete, you should use the 
      // ControlDefinition functionality for implementing commands.
    }

    public object Automation
    {
      // This property is provided to allow the AddIn to expose an API 
      // of its own to other programs. Typically, this  would be done by
      // implementing the AddIn's API interface in a class and returning 
      // that class object through this property.

      get
      {
        // TODO: Add ApplicationAddInServer.Automation getter implementation
        return null;
      }
    }

    #endregion

  }
}
