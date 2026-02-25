using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace TubeTabSlot
{
    public static class FeatureHelper
    {
        /// <summary>Tab extrusion depth in metres (5mm).</summary>
        public const double TAB_DEPTH_M = 0.010;

        /// <summary>Slot clearance offset in metres (0.4mm).</summary>
        public const double SLOT_CLEARANCE_M = 0.0005;

        /// <summary>
        /// Extrudes the tab sketch as a boss, merging with the tab body.
        /// Uses start offset to position the extrusion at the correct distance from the base plane.
        /// </summary>
        public static IFeature ExtrudeTab(
            IModelDoc2 doc,
            TabSketchResult tabSketch,
            IBody2 tabBody,
            double startOffset,
            double tabDepthM = TAB_DEPTH_M)
        {
            doc.ClearSelection2(true);
            tabSketch.SketchFeature.Select2(false, 0);
            // Mark 8 = feature scope body for boss extrude with UseAutoSelect=false
            doc.Extension.SelectByID2(tabBody.Name, "SOLIDBODY", 0, 0, 0, true, 8, null, 0);

            // Start condition: offset from sketch plane
            int startCond = Math.Abs(startOffset) > 1e-8
                ? (int)swStartConditions_e.swStartOffset
                : (int)swStartConditions_e.swStartSketchPlane;
            double absOffset = Math.Abs(startOffset);
            bool flipStart = startOffset < 0;

            IFeature extFeat = (IFeature)doc.FeatureManager.FeatureExtrusion3(
                true,  // Sd - single direction
                false, // Flip
                false, // Dir
                (int)swEndConditions_e.swEndCondBlind, // T1 end condition
                (int)swEndConditions_e.swEndCondBlind, // T2 end condition
                tabDepthM,    // T1 depth
                0,            // T2 depth
                false, false, // T1 draft on, outward
                false, false, // T2 draft on, outward
                0, 0,         // T1, T2 draft angles
                false, false, // T1, T2 offset reverse
                false, false, // T1, T2 translate surface
                true,         // Merge
                true,         // UseFeatScope
                false,        // UseAutoSelect - locked to pre-selected tab body
                startCond,    // T0 - start condition
                absOffset,    // T0 depth
                flipStart     // FlipStartOffset
            );

            return extFeat;
        }

        /// <summary>
        /// Creates a slot cut by drawing a slightly larger tab profile (offset by clearance)
        /// on the same plane, then extruding a cut through all into the slot body.
        /// </summary>
        public static IFeature CutSlotDirect(
            SldWorks swApp,
            IModelDoc2 doc,
            IFeature tabPlane,
            double outerRadius,
            double wallThickness,
            double[] centerPoint,
            IBody2 slotBody,
            double startOffset,
            double[] slotDirModel,
            bool farSide,
            double tabDepthM = TAB_DEPTH_M)
        {
            // Draw a slightly larger tab profile (offset by clearance)
            double expandedOuterRadius = outerRadius + SLOT_CLEARANCE_M;
            double shrunkInnerRadius = (outerRadius - wallThickness) - SLOT_CLEARANCE_M;

            // Half-angle slightly expanded for clearance
            double halfAngle = (TabSketchHelper.TAB_WIDTH_M / 2.0 + SLOT_CLEARANCE_M) / expandedOuterRadius;

            doc.ClearSelection2(true);
            tabPlane.Select2(false, 0);
            doc.SketchManager.InsertSketch(true);

            // Transform center point to sketch coordinates
            ISketch activeSketch = doc.SketchManager.ActiveSketch;
            MathTransform sketchXform = activeSketch.ModelToSketchTransform;

            IMathUtility mathUtil = (IMathUtility)swApp.GetMathUtility();
            IMathPoint centerMath = (IMathPoint)mathUtil.CreatePoint(centerPoint);
            centerMath = (IMathPoint)centerMath.MultiplyTransform(sketchXform);
            double[] centerSk = (double[])centerMath.ArrayData;

            double cx = centerSk[0];
            double cy = centerSk[1];

            // Project the slot tube direction into sketch space to match tab arc orientation.
            IMathVector slotDirVec = (IMathVector)mathUtil.CreateVector(slotDirModel);
            slotDirVec = (IMathVector)slotDirVec.MultiplyTransform(sketchXform);
            double[] slotDirSk = (double[])slotDirVec.ArrayData;
            double baseAngle = Math.Atan2(slotDirSk[1], slotDirSk[0]);

            // Far side: slot profile is on the opposite side of the tube cross-section
            if (farSide)
                baseAngle += Math.PI;

            // Outer arc (expanded)
            double outerStartX = cx + expandedOuterRadius * Math.Cos(baseAngle - halfAngle);
            double outerStartY = cy + expandedOuterRadius * Math.Sin(baseAngle - halfAngle);
            double outerEndX = cx + expandedOuterRadius * Math.Cos(baseAngle + halfAngle);
            double outerEndY = cy + expandedOuterRadius * Math.Sin(baseAngle + halfAngle);

            doc.SketchManager.CreateArc(
                cx, cy, 0,
                outerStartX, outerStartY, 0,
                outerEndX, outerEndY, 0, 1);

            // Inner arc (shrunk)
            double innerStartX = cx + shrunkInnerRadius * Math.Cos(baseAngle - halfAngle);
            double innerStartY = cy + shrunkInnerRadius * Math.Sin(baseAngle - halfAngle);
            double innerEndX = cx + shrunkInnerRadius * Math.Cos(baseAngle + halfAngle);
            double innerEndY = cy + shrunkInnerRadius * Math.Sin(baseAngle + halfAngle);

            doc.SketchManager.CreateArc(
                cx, cy, 0,
                innerStartX, innerStartY, 0,
                innerEndX, innerEndY, 0, 1);

            // Radial lines
            doc.SketchManager.CreateLine(
                outerStartX, outerStartY, 0,
                innerStartX, innerStartY, 0);

            doc.SketchManager.CreateLine(
                outerEndX, outerEndY, 0,
                innerEndX, innerEndY, 0);

            ISketch slotSketch = doc.SketchManager.ActiveSketch;
            IFeature slotSketchFeat = (IFeature)slotSketch;
            doc.SketchManager.InsertSketch(true);

            // Blind cut into the slot tube, scoped to slot body only
            doc.ClearSelection2(true);
            slotSketchFeat.Select2(false, 0);
            // Mark 8 = feature scope body for cut with UseAutoSelect=false
            doc.Extension.SelectByID2(slotBody.Name, "SOLIDBODY", 0, 0, 0, true, 8, null, 0);

            // Start condition: offset from sketch plane
            int startCond = Math.Abs(startOffset) > 1e-8
                ? (int)swStartConditions_e.swStartOffset
                : (int)swStartConditions_e.swStartSketchPlane;
            double absOffset = Math.Abs(startOffset);
            bool flipStart = startOffset < 0;

            IFeature cutFeat = (IFeature)doc.FeatureManager.FeatureCut4(
                true,          // Sd - single direction
                false,         // Flip
                true,          // Dir
                (int)swEndConditions_e.swEndCondBlind,  // T1 - blind 5mm
                (int)swEndConditions_e.swEndCondBlind,  // T2 (unused, Sd=true)
                tabDepthM, 0,   // T1 depth, T2 unused
                false, false,   // T1 draft on, outward
                false, false,   // T2 draft on, outward
                0, 0,           // T1, T2 draft angles
                false, false,   // T1, T2 offset reverse
                false, false,   // T1, T2 translate surface
                false,          // NormalCut
                true,           // UseFeatScope
                false,          // UseAutoSelect - locked to pre-selected slot body
                false,          // AssemblyFeatureScope
                false,          // AutoDetermineAssemblyScope
                false,          // AssemblyAutoPropagate
                startCond,      // T0 start condition
                absOffset,      // T0 depth
                flipStart,      // FlipStartOffset
                false           // CapEnds
            );

            return cutFeat;
        }
    }
}
