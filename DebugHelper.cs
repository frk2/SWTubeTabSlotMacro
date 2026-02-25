using System;
using SolidWorks.Interop.sldworks;

namespace TubeTabSlot
{
    public static class DebugHelper
    {
        private const double CROSS_LARGE  = 0.005;  // final Point1/Point2
        private const double CROSS_MEDIUM = 0.004;  // slot-axis intermediate points
        private const double CROSS_SMALL  = 0.003;  // extrusion starts
        private const double ARROW_LEN    = 0.015;  // extrusion direction arrow

        /// <summary>
        /// Draws a 3D sketch with three layers of markers:
        ///
        ///   LARGE cross  (5 mm) — Point1 / Point2
        ///                         the final placement points (slotAxisPt + tabDir shift)
        ///
        ///   MEDIUM cross (4 mm) — SlotAxisPt1 / SlotAxisPt2
        ///                         the intermediate slot-axis points BEFORE the tabDir shift
        ///
        ///   SMALL cross  (3 mm) + arrow — extrusion starts
        ///                         planePoint + planeNormal * offsetN
        ///                         where SolidWorks actually begins the boss/cut
        ///
        /// Lines connect each pair so the shifts are easy to trace.
        /// </summary>
        public static void DrawTabPlacementPoints(
            IModelDoc2 doc,
            double[]   point1,
            double[]   point2,
            double[]   slotAxisPt1,
            double[]   slotAxisPt2,
            double[]   planeNormal,
            double[]   planePoint,
            double     offset1,
            double     offset2)
        {
#if DEBUG
            doc.ClearSelection2(true);
            doc.SketchManager.Insert3DSketch(true);

            // --- Final placement points (large crosses) ---
            DrawCross(doc, point1, CROSS_LARGE);
            DrawCross(doc, point2, CROSS_LARGE);
            doc.SketchManager.CreateLine(
                point1[0], point1[1], point1[2],
                point2[0], point2[1], point2[2]);

            // --- Slot-axis intermediate points (medium crosses) ---
            DrawCross(doc, slotAxisPt1, CROSS_MEDIUM);
            DrawCross(doc, slotAxisPt2, CROSS_MEDIUM);

            // Lines from slotAxisPt to the shifted result so the tabDir shift is visible
            doc.SketchManager.CreateLine(
                slotAxisPt1[0], slotAxisPt1[1], slotAxisPt1[2],
                point1[0],      point1[1],      point1[2]);
            doc.SketchManager.CreateLine(
                slotAxisPt2[0], slotAxisPt2[1], slotAxisPt2[2],
                point2[0],      point2[1],      point2[2]);

            // --- Extrusion starts (small crosses + direction arrows) ---
            double[] extStart1 = Add(planePoint, Scale(planeNormal, offset1));
            double[] extStart2 = Add(planePoint, Scale(planeNormal, offset2));

            DrawCross(doc, extStart1, CROSS_SMALL);
            DrawCross(doc, extStart2, CROSS_SMALL);
            DrawArrow(doc, extStart1, planeNormal, ARROW_LEN);
            DrawArrow(doc, extStart2, planeNormal, ARROW_LEN);

            // Lines from final point to its extrusion start to show the projection gap
            doc.SketchManager.CreateLine(
                point1[0],      point1[1],      point1[2],
                extStart1[0],   extStart1[1],   extStart1[2]);
            doc.SketchManager.CreateLine(
                point2[0],      point2[1],      point2[2],
                extStart2[0],   extStart2[1],   extStart2[2]);

            doc.SketchManager.Insert3DSketch(true);
            doc.GraphicsRedraw2();
#endif
        }

        /// <summary>
        /// Draws a 3D sketch showing the five arc-construction points for one tab profile:
        ///   LARGE  cross  — tube centre
        ///   MEDIUM cross  — outer arc start / end
        ///   SMALL  cross  — inner arc start / end
        /// Plus lines: centre->outerStart, centre->outerEnd (true radii),
        ///             outerStart->innerStart, outerEnd->innerEnd (intended connecting lines).
        /// </summary>
        public static void DrawArcEndpoints(
            IModelDoc2 doc,
            double[]   center,
            double[]   outerStart, double[] outerEnd,
            double[]   innerStart, double[] innerEnd)
        {
#if DEBUG
            doc.ClearSelection2(true);
            doc.SketchManager.Insert3DSketch(true);

            DrawCross(doc, center,     CROSS_LARGE);   // tube centre
            DrawCross(doc, outerStart, CROSS_MEDIUM);  // outer arc start
            DrawCross(doc, outerEnd,   CROSS_MEDIUM);  // outer arc end
            DrawCross(doc, innerStart, CROSS_SMALL);   // inner arc start
            DrawCross(doc, innerEnd,   CROSS_SMALL);   // inner arc end

            // True radial lines from centre to outer endpoints
            doc.SketchManager.CreateLine(
                center[0],     center[1],     center[2],
                outerStart[0], outerStart[1], outerStart[2]);
            doc.SketchManager.CreateLine(
                center[0],    center[1],    center[2],
                outerEnd[0],  outerEnd[1],  outerEnd[2]);

            // Intended connecting lines (outer -> inner)
            doc.SketchManager.CreateLine(
                outerStart[0], outerStart[1], outerStart[2],
                innerStart[0], innerStart[1], innerStart[2]);
            doc.SketchManager.CreateLine(
                outerEnd[0], outerEnd[1], outerEnd[2],
                innerEnd[0], innerEnd[1], innerEnd[2]);

            doc.SketchManager.Insert3DSketch(true);
            doc.GraphicsRedraw2();
#endif
        }

        // -- helpers --

        private static void DrawCross(IModelDoc2 doc, double[] pt, double size)
        {
            double x = pt[0], y = pt[1], z = pt[2], s = size;
            doc.SketchManager.CreateLine(x - s, y,     z,     x + s, y,     z    );
            doc.SketchManager.CreateLine(x,     y - s, z,     x,     y + s, z    );
            doc.SketchManager.CreateLine(x,     y,     z - s, x,     y,     z + s);
        }

        private static void DrawArrow(IModelDoc2 doc, double[] origin, double[] dir, double length)
        {
            double[] tip = Add(origin, Scale(dir, length));
            doc.SketchManager.CreateLine(
                origin[0], origin[1], origin[2],
                tip[0],    tip[1],    tip[2]);
        }

        private static double[] Add(double[] a, double[] b)
        {
            return new double[] { a[0] + b[0], a[1] + b[1], a[2] + b[2] };
        }

        private static double[] Scale(double[] v, double s)
        {
            return new double[] { v[0] * s, v[1] * s, v[2] * s };
        }
    }
}
