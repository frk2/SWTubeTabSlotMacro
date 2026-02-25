using System;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace TubeTabSlot
{
    public class TabPlacementData
    {
        public double[] Point1;
        public double[] Point2;
        public IFeature BasePlane;
        public double Offset1;
        public double Offset2;
        /// <summary>Unit vector along the slot tube axis in model space.</summary>
        public double[] SlotDir;
        /// <summary>Base plane normal (unit vector, parallel to tab tube axis).</summary>
        public double[] PlaneNormal;
        /// <summary>Origin of the base plane in model space.</summary>
        public double[] PlanePoint;
        /// <summary>Intermediate point on the slot axis before the tabDir shift (debug).</summary>
        public double[] SlotAxisPt1;
        /// <summary>Intermediate point on the slot axis before the tabDir shift (debug).</summary>
        public double[] SlotAxisPt2;
    }

    public static class GeometryHelper
    {
        public static TabPlacementData ComputeTabPlacement(
            SldWorks swApp,
            WeldmentSelectionData sel)
        {
            IModelDoc2 doc = (IModelDoc2)swApp.ActiveDoc;
            IMathUtility mathUtil = (IMathUtility)swApp.GetMathUtility();
            var data = new TabPlacementData();

            // Get sketch line endpoints in MODEL coordinates
            double[] tabStart, tabEnd, slotStart, slotEnd;
            GetSketchLineEndpoints(sel.TabSketchLine, mathUtil, out tabStart, out tabEnd);
            GetSketchLineEndpoints(sel.SlotSketchLine, mathUtil, out slotStart, out slotEnd);

            double[] tabDir = Normalize(Sub(tabEnd, tabStart));
            double[] slotDir = Normalize(Sub(slotEnd, slotStart));
            data.SlotDir = slotDir;

            // Calculate the tab placement points
            data.Point1 = FindIntersectionPoint(tabStart, tabEnd, slotStart, slotEnd,
                sel.TabOuterRadius, sel.SlotOuterRadius, true,  out data.SlotAxisPt1);
            data.Point2 = FindIntersectionPoint(tabStart, tabEnd, slotStart, slotEnd,
                sel.TabOuterRadius, sel.SlotOuterRadius, false, out data.SlotAxisPt2);

            // Find the existing plane from the tab structural member's sub-tree
            IFeature basePlane = FindStructuralMemberPlane(sel.TabStructuralFeature, tabDir);

            if (basePlane == null)
                basePlane = FindDefaultPlanePerpToDir(doc, tabDir);

            if (basePlane == null)
                return data;

            data.BasePlane = basePlane;

            // Get the base plane's position to calculate offsets
            double[] planeNormal;
            double[] planePoint;
            GetPlaneInfo(basePlane, out planeNormal, out planePoint);

            data.PlaneNormal = planeNormal;
            data.PlanePoint  = planePoint;

            // Calculate signed offsets from base plane to each tab point
            double[] toPoint1 = Sub(data.Point1, planePoint);
            data.Offset1 = Dot(toPoint1, planeNormal);

            double[] toPoint2 = Sub(data.Point2, planePoint);
            data.Offset2 = Dot(toPoint2, planeNormal);

            return data;
        }

        /// <summary>
        /// Finds a reference plane in the structural member's sub-feature tree
        /// that is perpendicular to the tube axis (normal parallel to tabDir).
        /// </summary>
        private static IFeature FindStructuralMemberPlane(IFeature structFeat, double[] tabDir)
        {
            double[] normTabDir = Normalize(tabDir);

            IFeature subFeat = (IFeature)structFeat.GetFirstSubFeature();
            while (subFeat != null)
            {
                string typeName = subFeat.GetTypeName2();

                // Check for reference plane types
                if (typeName == "RefPlane" || typeName == "ReferencePlane")
                {
                    // Verify the plane is perpendicular to the tube axis
                    double[] normal;
                    double[] point;
                    GetPlaneInfo(subFeat, out normal, out point);

                    if (normal != null)
                    {
                        double dot = Math.Abs(
                            normal[0] * normTabDir[0] +
                            normal[1] * normTabDir[1] +
                            normal[2] * normTabDir[2]);

                        if (dot > 0.99)
                            return subFeat;
                    }

                    // If we can't verify direction, return it anyway
                    // (the structural member plane is likely correct)
                    return subFeat;
                }

                subFeat = (IFeature)subFeat.GetNextSubFeature();
            }

            return null;
        }

        /// <summary>
        /// Fallback: find a default plane (Front/Top/Right) whose normal is
        /// closest to parallel with the tab direction.
        /// </summary>
        private static IFeature FindDefaultPlanePerpToDir(IModelDoc2 doc, double[] tabDir)
        {
            double[] normDir = Normalize(tabDir);
            IFeature bestPlane = null;
            double bestDot = 0;

            IFeature feat = (IFeature)doc.FirstFeature();
            while (feat != null)
            {
                string typeName = feat.GetTypeName2();
                if (typeName == "RefPlane")
                {
                    double[] normal;
                    double[] point;
                    GetPlaneInfo(feat, out normal, out point);

                    if (normal != null)
                    {
                        double dot = Math.Abs(
                            normal[0] * normDir[0] +
                            normal[1] * normDir[1] +
                            normal[2] * normDir[2]);

                        if (dot > bestDot)
                        {
                            bestDot = dot;
                            bestPlane = feat;
                        }
                    }
                }

                // Only check the first few features (default planes)
                if (feat.GetTypeName2() == "OriginProfileFeature")
                    break;

                feat = (IFeature)feat.GetNextFeature();
            }

            return bestPlane;
        }

        /// <summary>
        /// Gets the normal and a root point from a reference plane feature.
        /// </summary>
        private static void GetPlaneInfo(IFeature planeFeat, out double[] normal, out double[] point)
        {
            normal = null;
            point = null;

            try
            {
                IRefPlane refPlane = (IRefPlane)planeFeat.GetSpecificFeature2();
                if (refPlane == null) return;

                MathTransform transform = refPlane.Transform;
                if (transform == null) return;

                double[] xformData = (double[])transform.ArrayData;

                // The plane normal is the Z-axis of the transform (3rd column of rotation matrix)
                normal = Normalize(new double[] { xformData[6], xformData[7], xformData[8] });

                // The plane origin is the translation part
                point = new double[] { xformData[9], xformData[10], xformData[11] };
            }
            catch { }
        }


        /// <summary>
        /// Gets sketch line endpoints in MODEL coordinates.
        /// </summary>
        public static void GetSketchLineEndpoints(
            ISketchLine line, IMathUtility mathUtil,
            out double[] start, out double[] end)
        {
            ISketchPoint sp = (ISketchPoint)line.GetStartPoint2();
            ISketchPoint ep = (ISketchPoint)line.GetEndPoint2();

            ISketchSegment seg = (ISketchSegment)line;
            ISketch sketch = seg.GetSketch();
            MathTransform sketchToModel = (MathTransform)sketch.ModelToSketchTransform.Inverse();

            IMathPoint startMath = (IMathPoint)mathUtil.CreatePoint(
                new double[] { sp.X, sp.Y, sp.Z });
            startMath = (IMathPoint)startMath.MultiplyTransform(sketchToModel);
            start = (double[])startMath.ArrayData;

            IMathPoint endMath = (IMathPoint)mathUtil.CreatePoint(
                new double[] { ep.X, ep.Y, ep.Z });
            endMath = (IMathPoint)endMath.MultiplyTransform(sketchToModel);
            end = (double[])endMath.ArrayData;
        }

        private static double[] FindIntersectionPoint(
            double[] tabStart, double[] tabEnd,
            double[] slotStart, double[] slotEnd,
            double tabRadius,
            double slotRadius,
            bool firstPoint,
            out double[] slotAxisPt)
        {
            double[] tabDir = Normalize(Sub(tabEnd, tabStart));
            double[] slotDir = Normalize(Sub(slotEnd, slotStart));

            double[] intersectionPt = ClosestPointBetweenLines(tabStart, tabDir, slotStart, slotDir);

            double sinAngle = Length(Cross(tabDir, slotDir));
            if (sinAngle < 1e-10) sinAngle = 1e-10;

            double cosAngle = sinAngle; // need the complement cos angle which is sin angle

            // Step 1: move along slot axis to the tab tube outer surface
            double halfLen  = tabRadius / sinAngle;

            // Step 2: shift along tab axis toward the tab tube by slotRadius / cosAngle.
            // This corrects the projection so Offset1/Offset2 land at the right position
            // on the tab tube axis.
            double tabShift = cosAngle > 1e-10 ? slotRadius / cosAngle : 0.0;

            double sign = firstPoint ? -1.0 : 1.0;

            slotAxisPt      = Add(intersectionPt, Scale(slotDir, sign * halfLen));
            double[] result = Add(slotAxisPt,     Scale(tabDir,  -tabShift));

#if DEBUG
            MessageBox.Show(string.Format(
                "[TubeSlot] {0}: tabRadius={1:F3}mm  slotRadius={2:F3}mm  sinAngle={3:F6}  cosAngle={4:F6}  halfLen={5:F3}mm  tabShift={6:F3}mm",
                firstPoint ? "Point1" : "Point2",
                tabRadius  * 1000.0,
                slotRadius * 1000.0,
                sinAngle,
                cosAngle,
                halfLen   * 1000.0,
                tabShift  * 1000.0));
#endif

            return result;
        }

        // --- Vector math utilities ---

        public static double[] Sub(double[] a, double[] b)
        {
            return new double[] { a[0] - b[0], a[1] - b[1], a[2] - b[2] };
        }

        public static double[] Add(double[] a, double[] b)
        {
            return new double[] { a[0] + b[0], a[1] + b[1], a[2] + b[2] };
        }

        public static double[] Scale(double[] v, double s)
        {
            return new double[] { v[0] * s, v[1] * s, v[2] * s };
        }

        public static double Dot(double[] a, double[] b)
        {
            return a[0] * b[0] + a[1] * b[1] + a[2] * b[2];
        }

        public static double[] Cross(double[] a, double[] b)
        {
            return new double[]
            {
                a[1] * b[2] - a[2] * b[1],
                a[2] * b[0] - a[0] * b[2],
                a[0] * b[1] - a[1] * b[0]
            };
        }

        public static double Length(double[] v)
        {
            return Math.Sqrt(v[0] * v[0] + v[1] * v[1] + v[2] * v[2]);
        }

        public static double[] Normalize(double[] v)
        {
            double len = Length(v);
            if (len < 1e-15) return new double[] { 0, 0, 0 };
            return new double[] { v[0] / len, v[1] / len, v[2] / len };
        }

        public static double[] ClosestPointBetweenLines(
            double[] p1, double[] d1,
            double[] p2, double[] d2)
        {
            double[] w0 = Sub(p1, p2);
            double a = Dot(d1, d1);
            double b = Dot(d1, d2);
            double c = Dot(d2, d2);
            double d = Dot(d1, w0);
            double e = Dot(d2, w0);

            double denom = a * c - b * b;
            double t, s;

            if (Math.Abs(denom) < 1e-15)
            {
                t = 0;
                s = e / c;
            }
            else
            {
                t = (b * e - c * d) / denom;
                s = (a * e - b * d) / denom;
            }

            double[] pt1 = Add(p1, Scale(d1, t));
            double[] pt2 = Add(p2, Scale(d2, s));

            return Scale(Add(pt1, pt2), 0.5);
        }
    }
}
