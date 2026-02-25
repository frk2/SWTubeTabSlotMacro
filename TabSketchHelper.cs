using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace TubeTabSlot
{
    /// <summary>Holds the result of drawing a tab sketch on the base plane.</summary>
    public class TabSketchResult
    {
        public IFeature SketchFeature;
        public ISketch  Sketch;
    }

    public static class TabSketchHelper
    {
        /// <summary>Target arc length of the tab profile in metres (~5 mm).</summary>
        public const double TAB_WIDTH_M = 0.005;

        /// <summary>
        /// Draws a tab cross-section sketch on the base plane.
        ///
        /// Geometry:
        ///   The plane containing both tube axes (tab + slot) intersects the base plane
        ///   along a line through the tube centre in the direction of slotDirModel
        ///   projected onto the base plane.  That line crosses the tube OD circle at
        ///   two points on opposite sides of the tube:
        ///
        ///     farSide = false  (near)  ->  toward   the slotDir projection
        ///     farSide = true   (far)   ->  opposite the slotDir projection
        ///
        /// Profile: outer arc (~= TAB_WIDTH_M) + inner arc + two radial lines.
        /// </summary>
        public static TabSketchResult DrawTabProfile(
            SldWorks    swApp,
            IModelDoc2  doc,
            IFeature    plane,
            double      outerRadius,
            double      wallThickness,
            double[]    tubeCenterModel,  // Any point on the tab tube axis (model space)
            double[]    slotDirModel,     // Slot tube axis direction (model space)
            bool        farSide)
        {
            double innerRadius = outerRadius - wallThickness;
            double halfAngle   = (TAB_WIDTH_M / 2.0) / outerRadius;

            doc.ClearSelection2(true);
            plane.Select2(false, 0);
            doc.SketchManager.AddToDB = true;
            doc.SketchManager.InsertSketch(true);

            ISketch       activeSketch = doc.SketchManager.ActiveSketch;
            MathTransform sketchXform  = activeSketch.ModelToSketchTransform;
            IMathUtility  mathUtil     = (IMathUtility)swApp.GetMathUtility();

            // Tube centre in 2-D sketch coordinates
            IMathPoint centerMath = (IMathPoint)mathUtil.CreatePoint(tubeCenterModel);
            centerMath = (IMathPoint)centerMath.MultiplyTransform(sketchXform);
            double[] centerSk = (double[])centerMath.ArrayData;
            double cx = centerSk[0];
            double cy = centerSk[1];

            // Project slotDir onto the sketch plane (rotation only) to get arc orientation.
            // This is the direction of the common-plane / base-plane intersection line.
            IMathVector slotDirVec = (IMathVector)mathUtil.CreateVector(slotDirModel);
            slotDirVec = (IMathVector)slotDirVec.MultiplyTransform(sketchXform);
            double[] slotDirSk = (double[])slotDirVec.ArrayData;
            double baseAngle = Math.Atan2(slotDirSk[1], slotDirSk[0]);

            // Far side: arc midpoint is on the opposite side of the tube cross-section
            if (farSide)
                baseAngle += Math.PI;

            // Outer arc
            double outerStartX = cx + outerRadius * Math.Cos(baseAngle - halfAngle);
            double outerStartY = cy + outerRadius * Math.Sin(baseAngle - halfAngle);
            double outerEndX   = cx + outerRadius * Math.Cos(baseAngle + halfAngle);
            double outerEndY   = cy + outerRadius * Math.Sin(baseAngle + halfAngle);

            doc.SketchManager.CreateArc(
                cx, cy, 0,
                outerStartX, outerStartY, 0,
                outerEndX, outerEndY, 0, 1);

            // Inner arc
            double innerStartX = cx + innerRadius * Math.Cos(baseAngle - halfAngle);
            double innerStartY = cy + innerRadius * Math.Sin(baseAngle - halfAngle);
            double innerEndX   = cx + innerRadius * Math.Cos(baseAngle + halfAngle);
            double innerEndY   = cy + innerRadius * Math.Sin(baseAngle + halfAngle);

            // Convert all arc endpoints back to model space for 3D debug visualisation.
            // Must be done while the sketch is still active (sketchXform is still valid).
            MathTransform sketchToModel = (MathTransform)sketchXform.Inverse();
            double[] centerModel     = SketchPtToModel(mathUtil, sketchToModel, cx,          cy);
            double[] outerStartModel = SketchPtToModel(mathUtil, sketchToModel, outerStartX, outerStartY);
            double[] outerEndModel   = SketchPtToModel(mathUtil, sketchToModel, outerEndX,   outerEndY);
            double[] innerStartModel = SketchPtToModel(mathUtil, sketchToModel, innerStartX, innerStartY);
            double[] innerEndModel   = SketchPtToModel(mathUtil, sketchToModel, innerEndX,   innerEndY);

            doc.SketchManager.CreateArc(
                cx, cy, 0,
                innerStartX, innerStartY, 0,
                innerEndX, innerEndY, 0, 1);

            // Radial lines connecting arc endpoints
            doc.SketchManager.CreateLine(
                outerStartX, outerStartY, 0,
                innerStartX, innerStartY, 0);

            doc.SketchManager.CreateLine(
                outerEndX, outerEndY, 0,
                innerEndX, innerEndY, 0);

            ISketch sketch        = doc.SketchManager.ActiveSketch;
            IFeature sketchFeature = (IFeature)sketch;
            doc.SketchManager.InsertSketch(true);

            // Draw 3D debug markers now that the 2D sketch is closed.
#if DEBUG
            DebugHelper.DrawArcEndpoints(doc, centerModel,
                outerStartModel, outerEndModel,
                innerStartModel, innerEndModel);
#endif

            return new TabSketchResult { SketchFeature = sketchFeature, Sketch = sketch };
        }

        /// <summary>Transforms a 2D sketch-space point (z=0) to model space.</summary>
        private static double[] SketchPtToModel(IMathUtility mathUtil, MathTransform skToModel, double x, double y)
        {
            IMathPoint pt = (IMathPoint)mathUtil.CreatePoint(new double[] { x, y, 0 });
            pt = (IMathPoint)pt.MultiplyTransform(skToModel);
            return (double[])pt.ArrayData;
        }

        /// <summary>
        /// Projects a point in model space onto the tab tube axis (the weldment sketch line),
        /// returning the tube centre at that axial cross-section position.
        /// </summary>
        public static double[] FindTubeCenterAtPoint(
            SldWorks     swApp,
            ISketchLine  sketchLine,
            double[]     tabPoint)
        {
            IMathUtility mathUtil = (IMathUtility)swApp.GetMathUtility();
            ISketchPoint sp = (ISketchPoint)sketchLine.GetStartPoint2();
            ISketchPoint ep = (ISketchPoint)sketchLine.GetEndPoint2();

            ISketchSegment seg          = (ISketchSegment)sketchLine;
            ISketch        sketch       = seg.GetSketch();
            MathTransform  sketchToModel = (MathTransform)sketch.ModelToSketchTransform.Inverse();

            IMathPoint startMath = (IMathPoint)mathUtil.CreatePoint(new double[] { sp.X, sp.Y, sp.Z });
            startMath = (IMathPoint)startMath.MultiplyTransform(sketchToModel);
            double[] lineStart = (double[])startMath.ArrayData;

            IMathPoint endMath = (IMathPoint)mathUtil.CreatePoint(new double[] { ep.X, ep.Y, ep.Z });
            endMath = (IMathPoint)endMath.MultiplyTransform(sketchToModel);
            double[] lineEnd = (double[])endMath.ArrayData;

            double[] lineDir = GeometryHelper.Normalize(GeometryHelper.Sub(lineEnd, lineStart));
            double[] w       = GeometryHelper.Sub(tabPoint, lineStart);
            double   proj    = GeometryHelper.Dot(w, lineDir);

            return GeometryHelper.Add(lineStart, GeometryHelper.Scale(lineDir, proj));
        }
    }
}
