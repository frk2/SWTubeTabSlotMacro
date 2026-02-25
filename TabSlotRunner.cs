using System;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace TubeTabSlot
{
    public static class TabSlotRunner
    {
        private const double FUDGE_OFFSET = -0.005;

        public static void RunFeatures(SldWorks swApp, TabSlotOptions opts)
        {
            IModelDoc2 doc = (IModelDoc2)swApp.ActiveDoc;

            WeldmentSelectionData sel       = SelectionHelper.GetSelections(swApp);
            TabPlacementData      placement = GeometryHelper.ComputeTabPlacement(swApp, sel);

            if (placement.BasePlane == null)
            {
                MessageBox.Show("Failed to find base plane for tab placement.",
                    "Tab & Slot", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

#if DEBUG
            DebugHelper.DrawTabPlacementPoints(
                doc,
                placement.Point1,      placement.Point2,
                placement.SlotAxisPt1, placement.SlotAxisPt2,
                placement.PlaneNormal, placement.PlanePoint,
                placement.Offset1,     placement.Offset2);
#endif

            double[] intersectionPt = new double[]
            {
                (placement.SlotAxisPt1[0] + placement.SlotAxisPt2[0]) * 0.5,
                (placement.SlotAxisPt1[1] + placement.SlotAxisPt2[1]) * 0.5,
                (placement.SlotAxisPt1[2] + placement.SlotAxisPt2[2]) * 0.5
            };

            double dist1 = Math.Abs(GeometryHelper.Dot(
                GeometryHelper.Sub(placement.Point1, intersectionPt), placement.PlaneNormal));
            double dist2 = Math.Abs(GeometryHelper.Dot(
                GeometryHelper.Sub(placement.Point2, intersectionPt), placement.PlaneNormal));

            double nearOffset = placement.Offset1 + FUDGE_OFFSET;
            double farOffset  = placement.Offset2 + FUDGE_OFFSET;
            double[] nearPt   = placement.Point1;
            double[] farPt    = placement.Point2;

            if (dist1 > dist2)
            {
                double tmp   = nearOffset; nearOffset = farOffset;  farOffset  = tmp;
                double[] tpt = nearPt;     nearPt     = farPt;      farPt      = tpt;
            }

            double[] tubeCenter = TabSketchHelper.FindTubeCenterAtPoint(
                swApp, sel.TabSketchLine, nearPt);

            double tabDepthM = opts.TabDepthMm / 1000.0;

            if (opts.Mode == TabMode.Both || opts.Mode == TabMode.NearOnly)
            {
                TabSketchResult nearSketch = TabSketchHelper.DrawTabProfile(
                    swApp, doc, placement.BasePlane,
                    sel.TabOuterRadius, sel.TabWallThickness,
                    tubeCenter, placement.SlotDir, farSide: false);

                FeatureHelper.ExtrudeTab(doc, nearSketch, sel.TabBody, nearOffset, tabDepthM);

                FeatureHelper.CutSlotDirect(
                    swApp, doc, placement.BasePlane,
                    sel.TabOuterRadius, sel.TabWallThickness,
                    tubeCenter, sel.SlotBody, nearOffset,
                    placement.SlotDir, farSide: false, tabDepthM: tabDepthM);
            }

            if (opts.Mode == TabMode.Both || opts.Mode == TabMode.FarOnly)
            {
                TabSketchResult farSketch = TabSketchHelper.DrawTabProfile(
                    swApp, doc, placement.BasePlane,
                    sel.TabOuterRadius, sel.TabWallThickness,
                    tubeCenter, placement.SlotDir, farSide: true);

                FeatureHelper.ExtrudeTab(doc, farSketch, sel.TabBody, farOffset, tabDepthM);

                FeatureHelper.CutSlotDirect(
                    swApp, doc, placement.BasePlane,
                    sel.TabOuterRadius, sel.TabWallThickness,
                    tubeCenter, sel.SlotBody, farOffset,
                    placement.SlotDir, farSide: true, tabDepthM: tabDepthM);
            }

            doc.ForceRebuild3(true);

            //MessageBox.Show("Tab & Slot features created successfully!",
            //    "Tab & Slot", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
