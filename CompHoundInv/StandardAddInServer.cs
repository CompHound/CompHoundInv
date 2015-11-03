using System;
using System.Runtime.InteropServices;
using Inventor;
using Microsoft.Win32;
using Microsoft.VisualBasic.Compatibility.VB6;
using System.Drawing;
using System.Reflection;
using System.IO;
using MsgBox = System.Windows.Forms.MessageBox;
using System.Diagnostics;
using System.Collections.Generic;
using System.Collections;
using MFU = MultiFileUploader;

namespace CompHoundInv
{
  /// <summary>
  /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
  /// that all Inventor AddIns are required to implement. The communication between Inventor and
  /// the AddIn is via the methods on this interface.
  /// </summary>
  [GuidAttribute("f2cbfc1c-e37f-4330-a3dd-12b617a6458f")]
  public class StandardAddInServer : Inventor.ApplicationAddInServer
  {

    // Inventor application object.
    private Inventor.Application m_inventorApplication;
    private Inventor.ButtonDefinition m_commandButton;
    private string m_clientKey = "<replace with client key>";
    private string m_clientSecret = "<replace with secret key>";

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
      Assembly asm = Assembly.GetExecutingAssembly();
      Stream btnStream = asm.GetManifestResourceStream("CompHoundInv.Button16x16.png");
      Bitmap btnBitmap = new Bitmap(btnStream);
      Icon icon = Icon.FromHandle(btnBitmap.GetHicon());
      stdole.IPictureDisp standardIconIPictureDisp;
      standardIconIPictureDisp = (stdole.IPictureDisp)Support.IconToIPicture(icon);

      CommandManager cmdMan = m_inventorApplication.CommandManager;
      m_commandButton = cmdMan.ControlDefinitions.AddButtonDefinition(
        "CompHound Uploader",
        "CompHoundUploader",
        CommandTypesEnum.kQueryOnlyCmdType,
        "{f2cbfc1c-e37f-4330-a3dd-12b617a6458f}", // AddIn GUID
        "Uploads information from the assembly components to CompHound",
        "CompHound Uploader",
        standardIconIPictureDisp,
        standardIconIPictureDisp,
        ButtonDisplayEnum.kDisplayTextInLearningMode);

      m_commandButton.OnExecute += m_commandButton_OnExecute;

      // Add to the default Add-In's tab
      m_commandButton.AutoAddToGUI(); 
    }

    void GetAllOccurrences(dynamic occs, IList<ComponentOccurrence> occList)
    {
      foreach (ComponentOccurrence occ in occs)
      {
        occList.Add(occ); 

        // Iterate the subcomponents
        GetAllOccurrences(occ.SubOccurrences, occList); 
      }
    }

    void m_commandButton_OnExecute(NameValueMap Context)
    {
      Document doc = m_inventorApplication.ActiveDocument;

      if (!(doc is AssemblyDocument))
      {
        MsgBox.Show("This command is only for assemblies!");
        return;
      }

      // Iterate through Components of the assembly
      AssemblyDocument asm = doc as AssemblyDocument;

      var occs = new List<ComponentOccurrence>();
      GetAllOccurrences(asm.ComponentDefinition.Occurrences, occs);

      int counter = 0;
      foreach (ComponentOccurrence occ in occs)
      {
        Debug.Print(
          "CompHound uploading instance {0} of {1}.",
           ++counter, occs.Count );
        InstanceData instanceData = new InstanceData(occ, counter);

        string message;
        if (!Util.Put("instances/" + instanceData._id,
          instanceData, out message))
        {
          counter--;
          break;
        }
      }

      // Upload files with references
      ArrayList fileList = new ArrayList();
      fileList.Add(asm.FullFileName);

      foreach (Document oRefDoc in asm.AllReferencedDocuments)
      {
        fileList.Add(oRefDoc.FullDocumentName);
      }

      MFU.Util u = new MFU.Util("https://developer.api.autodesk.com");
      string TopLevelAssembly_URN = u.UploadFilesWithReferences(
        m_clientKey, m_clientSecret, "adamscomphoundbucket", fileList, null, true);

      Debug.Print(
          "CompHound uploaded {0} instance{1}.",
          counter, counter > 1 ? "s" : "");
    }

    public void Deactivate()
    {
      // This method is called by Inventor when the AddIn is unloaded.
      // The AddIn will be unloaded either manually by the user or
      // when the Inventor session is terminated

      // TODO: Add ApplicationAddInServer.Deactivate implementation

      // Release objects.
      m_inventorApplication = null;
      m_commandButton = null;

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
