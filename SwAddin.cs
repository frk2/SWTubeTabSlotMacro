using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;
using SolidWorksTools.File;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace TubeTabSlot
{
    [Guid("5c9cc456-0c5f-43d7-921f-12a884a628a8"), ComVisible(true)]
    [SwAddin(
        Description = "Tab & slot features for intersecting weldment tubes",
        Title = "Tab & Slot",
        LoadAtStartup = true
        )]
    public class SwAddin : ISwAddin
    {
        #region Local Variables
        ISldWorks iSwApp = null;
        ICommandManager iCmdMgr = null;
        int addinID = 0;
        BitmapHandler iBmp;

        public const int mainCmdGroupID = 5;
        public const int mainItemID1 = 0;

        string[] mainIcons = new string[6];
        string[] icons = new string[6];

        TabSlotPMPHandler _pmpHandler;

        public ISldWorks SwApp
        {
            get { return iSwApp; }
        }

        #endregion

        #region SolidWorks Registration
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            SwAddinAttribute SWattr = null;
            Type type = typeof(SwAddin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false))
            {
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }

            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);

                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
                MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }

        #endregion

        #region ISwAddin Implementation
        public SwAddin()
        {
        }

        public bool ConnectToSW(object ThisSW, int cookie)
        {
            iSwApp = (ISldWorks)ThisSW;
            addinID = cookie;

            // Setup callbacks
            iSwApp.SetAddinCallbackInfo(0, this, addinID);

            // Setup the Command Manager
            iCmdMgr = iSwApp.GetCommandManager(cookie);
            AddCommandMgr();

            return true;
        }

        public bool DisconnectFromSW()
        {
            RemoveCommandMgr();

            _pmpHandler = null;

            Marshal.ReleaseComObject(iCmdMgr);
            iCmdMgr = null;
            iSwApp = null;

            // The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }
        #endregion

        #region UI Methods
        public void AddCommandMgr()
        {
            if (iBmp == null)
                iBmp = new BitmapHandler();

            Assembly thisAssembly = Assembly.GetAssembly(this.GetType());

            int cmdGroupErr = 0;
            bool ignorePrevious = false;

            object registryIDs;
            bool getDataResult = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);

            int[] knownIDs = new int[] { mainItemID1 };

            if (getDataResult)
            {
                if (!CompareIDs((int[])registryIDs, knownIDs))
                    ignorePrevious = true;
            }

            ICommandGroup cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, "Tab & Slot",
                "Tab & slot features for weldment tubes", "", -1, ignorePrevious, ref cmdGroupErr);

            // Load icons
            icons[0] = iBmp.CreateFileFromResourceBitmap("TubeTabSlot.toolbar20x.png", thisAssembly);
            icons[1] = iBmp.CreateFileFromResourceBitmap("TubeTabSlot.toolbar32x.png", thisAssembly);
            icons[2] = iBmp.CreateFileFromResourceBitmap("TubeTabSlot.toolbar40x.png", thisAssembly);
            icons[3] = iBmp.CreateFileFromResourceBitmap("TubeTabSlot.toolbar64x.png", thisAssembly);
            icons[4] = iBmp.CreateFileFromResourceBitmap("TubeTabSlot.toolbar96x.png", thisAssembly);
            icons[5] = iBmp.CreateFileFromResourceBitmap("TubeTabSlot.toolbar128x.png", thisAssembly);

            mainIcons[0] = iBmp.CreateFileFromResourceBitmap("TubeTabSlot.mainicon_20.png", thisAssembly);
            mainIcons[1] = iBmp.CreateFileFromResourceBitmap("TubeTabSlot.mainicon_32.png", thisAssembly);
            mainIcons[2] = iBmp.CreateFileFromResourceBitmap("TubeTabSlot.mainicon_40.png", thisAssembly);
            mainIcons[3] = iBmp.CreateFileFromResourceBitmap("TubeTabSlot.mainicon_64.png", thisAssembly);
            mainIcons[4] = iBmp.CreateFileFromResourceBitmap("TubeTabSlot.mainicon_96.png", thisAssembly);
            mainIcons[5] = iBmp.CreateFileFromResourceBitmap("TubeTabSlot.mainicon_128.png", thisAssembly);

            cmdGroup.MainIconList = mainIcons;
            cmdGroup.IconList = icons;

            int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
            int cmdIndex0 = cmdGroup.AddCommandItem2("Tab & Slot", -1,
                "Create tab & slot features on intersecting weldment tubes",
                "Tab & Slot", 0, "ShowTabSlotPMP", "EnableTabSlotPMP", mainItemID1, menuToolbarOption);

            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();

            // Register command tab for Part documents only
            int[] docTypes = new int[] { (int)swDocumentTypes_e.swDocPART };

            foreach (int type in docTypes)
            {
                CommandTab cmdTab = iCmdMgr.GetCommandTab(type, "Tab & Slot");

                if (cmdTab != null & !getDataResult | ignorePrevious)
                {
                    iCmdMgr.RemoveCommandTab(cmdTab);
                    cmdTab = null;
                }

                if (cmdTab == null)
                {
                    cmdTab = iCmdMgr.AddCommandTab(type, "Tab & Slot");

                    CommandTabBox cmdBox = cmdTab.AddCommandTabBox();

                    int[] cmdIDs = new int[] { cmdGroup.get_CommandID(cmdIndex0) };
                    int[] TextType = new int[] { (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal };

                    cmdBox.AddCommands(cmdIDs, TextType);
                }
            }

            thisAssembly = null;
        }

        public void RemoveCommandMgr()
        {
            iBmp.Dispose();
            iCmdMgr.RemoveCommandGroup(mainCmdGroupID);
        }

        public bool CompareIDs(int[] storedIDs, int[] addinIDs)
        {
            List<int> storedList = new List<int>(storedIDs);
            List<int> addinList = new List<int>(addinIDs);

            addinList.Sort();
            storedList.Sort();

            if (addinList.Count != storedList.Count)
                return false;

            for (int i = 0; i < addinList.Count; i++)
            {
                if (addinList[i] != storedList[i])
                    return false;
            }
            return true;
        }

        #endregion

        #region UI Callbacks
        public void ShowTabSlotPMP()
        {
            if (iSwApp.ActiveDoc == null)
            {
                MessageBox.Show("No active document. Please open a part with weldment bodies.",
                    "Tab & Slot", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Validate selection before showing the PMP
            try
            {
                SelectionHelper.GetSelections((SldWorks)iSwApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Tab & Slot", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Create handler lazily on first use
            if (_pmpHandler == null)
                _pmpHandler = new TabSlotPMPHandler(this);

            _pmpHandler.Show();
        }

        public int EnableTabSlotPMP()
        {
            if (iSwApp.ActiveDoc != null)
                return 1;
            return 0;
        }
        #endregion
    }
}
