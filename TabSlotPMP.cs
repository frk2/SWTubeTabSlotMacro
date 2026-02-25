using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;

namespace TubeTabSlot
{
    public enum TabMode { Both, NearOnly, FarOnly }

    public class TabSlotOptions
    {
        public TabMode Mode       = TabMode.Both;
        public double  TabDepthMm = 10.0;
    }

    /// <summary>
    /// SolidWorks Property Manager Page for Tab &amp; Slot options.
    /// Inherits PropertyManagerPage2Handler9 with [ComVisible(true)].
    /// All callbacks are PUBLIC methods so they are accessible via IDispatch
    /// late binding (how SolidWorks dispatches add-in callbacks).
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class TabSlotPMPHandler : IPropertyManagerPage2Handler9
    {
        // -- Control IDs --
        private const int GRP_PLACEMENT = 1;
        private const int OPT_BOTH      = 2;
        private const int OPT_NEAR      = 3;
        private const int OPT_FAR       = 4;
        private const int GRP_DIMS      = 5;
        private const int NUM_LABEL     = 6;
        private const int NUM_DEPTH     = 7;

        private IPropertyManagerPage2          _page;
        private IPropertyManagerPageNumberbox  _numDepth;
        private readonly SwAddin               _addin;
        private readonly TabSlotOptions        _opts = new TabSlotOptions();
        private bool                           _runOnAfterClose;

        public TabSlotPMPHandler(SwAddin addin)
        {
            _addin = addin;
        }

        public void Show()
        {
            // Reset options for each invocation
            _opts.Mode = TabMode.Both;
            _opts.TabDepthMm = 10.0;

            // Create a fresh PMP each time (matches the macro pattern that worked)
            int errors = 0;
            _page = (IPropertyManagerPage2)((SldWorks)_addin.SwApp).CreatePropertyManagerPage(
                "Tab & Slot",
                (int)(swPropertyManagerPageOptions_e.swPropertyManagerOptions_OkayButton |
                      swPropertyManagerPageOptions_e.swPropertyManagerOptions_CancelButton),
                this,
                ref errors);

            if (_page == null || errors != 0)
                throw new InvalidOperationException(
                    "Failed to create Property Manager Page. errors=" + errors);

            BuildControls();
            _page.Show2(0);
        }

        private void BuildControls()
        {
            int visEnabled = (int)swAddControlOptions_e.swControlOptions_Visible
                           + (int)swAddControlOptions_e.swControlOptions_Enabled;

            // -- Tab placement group --
            IPropertyManagerPageGroup grpPlacement =
                (IPropertyManagerPageGroup)_page.AddGroupBox(
                    GRP_PLACEMENT, "Tab placement",
                    (int)(swAddGroupBoxOptions_e.swGroupBoxOptions_Visible |
                          swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded));

            IPropertyManagerPageOption optBoth =
                (IPropertyManagerPageOption)grpPlacement.AddControl2(
                    OPT_BOTH,
                    (short)swPropertyManagerPageControlType_e.swControlType_Option,
                    "Both tabs",
                    (short)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent,
                    visEnabled,
                    "Create tabs on both sides of the intersection");
            optBoth.Checked = true;

            grpPlacement.AddControl2(
                OPT_NEAR,
                (short)swPropertyManagerPageControlType_e.swControlType_Option,
                "Near side only",
                (short)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent,
                visEnabled,
                "Tab only on the near side of the slot tube");

            grpPlacement.AddControl2(
                OPT_FAR,
                (short)swPropertyManagerPageControlType_e.swControlType_Option,
                "Far side only",
                (short)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent,
                visEnabled,
                "Tab only on the far side of the slot tube");

            // -- Dimensions group --
            IPropertyManagerPageGroup grpDims =
                (IPropertyManagerPageGroup)_page.AddGroupBox(
                    GRP_DIMS, "Dimensions",
                    (int)(swAddGroupBoxOptions_e.swGroupBoxOptions_Visible |
                          swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded));

            grpDims.AddControl2(
                NUM_LABEL,
                (short)swPropertyManagerPageControlType_e.swControlType_Label,
                "Tab depth (mm)",
                (short)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent,
                visEnabled, "Depth");

            _numDepth = (IPropertyManagerPageNumberbox)grpDims.AddControl2(
                NUM_DEPTH,
                (short)swPropertyManagerPageControlType_e.swControlType_Numberbox,
                "Tab Depth",
                (short)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent,
                visEnabled,
                "Extrusion depth for the tab boss and slot cut");

            if (_numDepth == null)
                throw new InvalidOperationException("Failed to create numberbox control.");

            _numDepth.SetRange2(
                (int)swNumberboxUnitType_e.swNumberBox_UnitlessDouble,
                1.0, 50.0, true, 0.5, 0.5, 0.5);
            _numDepth.Value = _opts.TabDepthMm;
        }

        // -- Callbacks (PUBLIC for IDispatch access) --

        public void OnClose(int Reason)
        {
            // Don't run features here — the PMP is still closing and the
            // document is locked for editing.  Set a flag so AfterClose
            // picks it up once the page is fully torn down.
            _runOnAfterClose =
                (Reason == (int)swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Okay);
        }

        public void OnOptionCheck(int Id)
        {
            switch (Id)
            {
                case OPT_BOTH: _opts.Mode = TabMode.Both;     break;
                case OPT_NEAR: _opts.Mode = TabMode.NearOnly; break;
                case OPT_FAR:  _opts.Mode = TabMode.FarOnly;  break;
            }
        }

        public void OnNumberboxChanged(int Id, double Value)
        {
            if (Id == NUM_DEPTH) _opts.TabDepthMm = Value;
        }

        // -- Required stubs (all public for IDispatch) --
        public void AfterActivation() { }
        public void AfterClose()
        {
            if (!_runOnAfterClose) return;
            _runOnAfterClose = false;

            try
            {
                TabSlotRunner.RunFeatures((SldWorks)_addin.SwApp, _opts);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    string.Format("Error creating features:\n\n{0}\n\n{1}", ex.Message, ex.StackTrace),
                    "Tab & Slot", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void OnButtonPress(int Id) { }
        public void OnCheckboxCheck(int Id, bool Checked) { }
        public void OnComboboxEditChanged(int Id, string Text) { }
        public void OnComboboxSelectionChanged(int Id, int Item) { }
        public void OnGainedFocus(int Id) { }
        public void OnGroupCheck(int Id, bool Checked) { }
        public void OnGroupExpand(int Id, bool Expanded) { }
        public bool OnHelp() { return true; }
        public bool OnKeystroke(int Wparam, int Message, int Lparam, int Id) { return false; }
        public void OnListboxRMBUp(int Id, int PosX, int PosY) { }
        public void OnListboxSelectionChanged(int Id, int Item) { }
        public void OnLostFocus(int Id) { }
        public bool OnNextPage() { return true; }
        public void OnNumberBoxTrackingCompleted(int Id, double Value) { }
        public void OnPopupMenuItem(int Id) { }
        public void OnPopupMenuItemUpdate(int Id, ref int retval) { }
        public bool OnPreview() { return true; }
        public bool OnPreviousPage() { return true; }
        public void OnRedo() { }
        public void OnSelectionboxCalloutCreated(int Id) { }
        public void OnSelectionboxCalloutDestroyed(int Id) { }
        public void OnSelectionboxFocusChanged(int Id) { }
        public void OnSelectionboxListChanged(int Id, int Count) { }
        public void OnSliderPositionChanged(int Id, double Value) { }
        public void OnSliderTrackingCompleted(int Id, double Value) { }
        public bool OnSubmitSelection(int Id, object Selection, int SelType, ref string ItemText) { return true; }
        public bool OnTabClicked(int Id) { return true; }
        public void OnTextboxChanged(int Id, string Text) { }
        public void OnUndo() { }
        public void OnWhatsNew() { }
        public int  OnActiveXControlCreated(int Id, bool Status) { return 0; }
        public int  OnWindowFromHandleControlCreated(int Id, bool Status) { return 0; }
    }
}
