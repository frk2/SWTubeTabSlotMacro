using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace TubeTabSlot
{
    public class WeldmentSelectionData
    {
        public IBody2 TabBody;
        public IBody2 SlotBody;
        public ISketchLine TabSketchLine;
        public ISketchLine SlotSketchLine;
        public double TabOuterRadius;
        public double TabWallThickness;
        public double SlotOuterRadius;
        public double SlotWallThickness;
        public IFeature TabStructuralFeature;
        public IFeature SlotStructuralFeature;
    }

    public static class SelectionHelper
    {
        public static WeldmentSelectionData GetSelections(SldWorks swApp)
        {
            IModelDoc2 doc = (IModelDoc2)swApp.ActiveDoc;
            if (doc == null)
                throw new InvalidOperationException("No active document.");

            ISelectionMgr selMgr = (ISelectionMgr)doc.SelectionManager;
            int selCount = selMgr.GetSelectedObjectCount2(-1);
            if (selCount != 2)
                throw new InvalidOperationException(
                    string.Format("Please select exactly 2 bodies (tab first, slot second). Currently {0} selected.", selCount));

            IBody2 tabBody = GetBodyFromSelection(selMgr, 1);
            IBody2 slotBody = GetBodyFromSelection(selMgr, 2);

            if (tabBody == null || slotBody == null)
                throw new InvalidOperationException(
                    "Could not obtain bodies from selections. Select faces on the two weldment tubes.");

            var data = new WeldmentSelectionData
            {
                TabBody = tabBody,
                SlotBody = slotBody
            };

            FindStructuralMemberInfo(swApp, doc, tabBody,
                out data.TabStructuralFeature, out data.TabSketchLine,
                out data.TabOuterRadius, out data.TabWallThickness);

            FindStructuralMemberInfo(swApp, doc, slotBody,
                out data.SlotStructuralFeature, out data.SlotSketchLine,
                out data.SlotOuterRadius, out data.SlotWallThickness);

            // Sanity check: tab and slot must not resolve to the same sketch line
            ValidateDistinctSketchLines(data.TabSketchLine, data.SlotSketchLine);

            return data;
        }

        private static IBody2 GetBodyFromSelection(ISelectionMgr selMgr, int index)
        {
            int selType = selMgr.GetSelectedObjectType3(index, -1);

            if (selType == (int)swSelectType_e.swSelSOLIDBODIES)
                return (IBody2)selMgr.GetSelectedObject6(index, -1);

            if (selType == (int)swSelectType_e.swSelFACES)
            {
                IFace2 face = (IFace2)selMgr.GetSelectedObject6(index, -1);
                if (face != null)
                    return (IBody2)face.GetBody();
                return null;
            }

            object obj = selMgr.GetSelectedObject6(index, -1);
            if (obj is IBody2)
                return (IBody2)obj;
            if (obj is IFace2)
                return (IBody2)((IFace2)obj).GetBody();

            return null;
        }

        /// <summary>
        /// Find the structural member feature, its sketch line, and profile dimensions.
        /// Matches the sketch line to the body by comparing the line direction
        /// to the body's cylindrical axis direction.
        /// </summary>
        private static void FindStructuralMemberInfo(
            SldWorks swApp,
            IModelDoc2 doc,
            IBody2 targetBody,
            out IFeature structFeat,
            out ISketchLine sketchLine,
            out double outerRadius,
            out double wallThickness)
        {
            structFeat = null;
            sketchLine = null;
            outerRadius = 0;
            wallThickness = 0;

            // First, get the body's cylindrical axis direction and a point on that axis
            double[] bodyAxisOrigin;
            double[] bodyAxisDir = GetBodyCylinderAxis(targetBody, out bodyAxisOrigin);

            IFeature feat = (IFeature)doc.FirstFeature();
            while (feat != null)
            {
                string typeName = feat.GetTypeName2();

                if (typeName == "WeldMemberFeat" || typeName == "StructuralMember")
                {
                    if (FeatureOwnsBody(feat, targetBody))
                    {
                        structFeat = feat;

                        // Primary: sketch line is a sub-feature of the structural member
                        string foundBy = null;
                        sketchLine = FindSketchLineInSubFeatures(swApp, feat, bodyAxisDir, bodyAxisOrigin);
                        if (sketchLine != null) foundBy = "FindSketchLineInSubFeatures";

                        // Secondary: search all preceding sketches with direction + position
                        if (sketchLine == null)
                        {
                            sketchLine = FindMatchingSketchLine(swApp, doc, feat, bodyAxisDir, bodyAxisOrigin);
                            if (sketchLine != null) foundBy = "FindMatchingSketchLine";
                        }

                        // Log what was found
                        if (sketchLine != null)
                        {
                            ISketchSegment seg = (ISketchSegment)sketchLine;
                            ISketch sk = seg.GetSketch() as ISketch;
                            string sketchName = sk != null ? ((IFeature)sk).Name : "?";
                            int[] idArr = seg.GetID() as int[];
                            int segId = (idArr != null && idArr.Length > 0) ? idArr[0] : -1;
#if DEBUG
                            MessageBox.Show(
                                string.Format(
                                    "Body: {0}\nStructural member: {1}\nFound by: {2}\nSketch: {3}  Segment ID: {4}",
                                    targetBody.Name, feat.Name, foundBy, sketchName, segId),
                                "Sketch line found", MessageBoxButtons.OK, MessageBoxIcon.Information);
#endif
                        }
                        else
                        {
                            throw new InvalidOperationException(string.Format(
                                "Could not find a matching sketch line for body '{0}' " +
                                "(structural member '{1}'). No sketch line passed both direction " +
                                "and proximity checks.",
                                targetBody.Name, feat.Name));
                        }

                        ExtractTubeProfile(targetBody, out outerRadius, out wallThickness);
                        return;
                    }
                }

                feat = (IFeature)feat.GetNextFeature();
            }

            if (structFeat == null)
                throw new InvalidOperationException(
                    "Could not find a structural-member feature for one of the selected bodies.");
        }

        /// <summary>
        /// Gets the axis direction and a point on the axis from the body's first cylindrical face.
        /// CylinderParams layout: [ox, oy, oz, dx, dy, dz, radius]
        /// </summary>
        private static double[] GetBodyCylinderAxis(IBody2 body, out double[] axisOrigin)
        {
            axisOrigin = null;
            object[] faces = (object[])body.GetFaces();
            if (faces == null) return null;

            double[] bestOrigin = null;
            double[] bestDir = null;
            double bestArea = 0;

            foreach (object f in faces)
            {
                IFace2 face = (IFace2)f;
                ISurface surf = (ISurface)face.GetSurface();
                if (surf.IsCylinder())
                {
                    double area = face.GetArea();
                    if (area > bestArea)
                    {
                        double[] cylParams = (double[])surf.CylinderParams;
                        bestOrigin = new double[] { cylParams[0], cylParams[1], cylParams[2] };
                        bestDir = new double[] { cylParams[3], cylParams[4], cylParams[5] };
                        bestArea = area;
                    }
                }
            }

            axisOrigin = bestOrigin;
            return bestDir;
        }

        /// <summary>
        /// Searches all sketches preceding the structural member for a line whose direction
        /// matches the body's cylindrical axis AND whose model-space position passes near
        /// the cylinder axis origin. Both checks are required to distinguish parallel tubes.
        /// </summary>
        private static ISketchLine FindMatchingSketchLine(
            SldWorks swApp, IModelDoc2 doc, IFeature structFeat,
            double[] bodyAxisDir, double[] bodyAxisOrigin)
        {
            if (bodyAxisDir == null)
                return FindPathSketchLine(doc, structFeat);

            double[] normAxis = Normalize(bodyAxisDir);
            IMathUtility mathUtil = (IMathUtility)swApp.GetMathUtility();

            IFeature feat = (IFeature)doc.FirstFeature();
            List<IFeature> sketchFeatures = new List<IFeature>();

            while (feat != null)
            {
                if (feat.Name == structFeat.Name)
                    break;

                string typeName = feat.GetTypeName2();
                if (typeName == "ProfileFeature" || typeName == "3DProfileFeature")
                    sketchFeatures.Add(feat);

                feat = (IFeature)feat.GetNextFeature();
            }

            ISketchLine directionOnlyMatch = null;
            var log = new System.Text.StringBuilder();
            log.AppendFormat("FindMatchingSketchLine for body axis ({0:F4},{1:F4},{2:F4})\n",
                normAxis[0], normAxis[1], normAxis[2]);
            log.AppendFormat("Axis origin: {0}\n",
                bodyAxisOrigin != null
                    ? string.Format("({0:F4},{1:F4},{2:F4})", bodyAxisOrigin[0], bodyAxisOrigin[1], bodyAxisOrigin[2])
                    : "null");
            log.AppendFormat("Sketches to search: {0}\n\n", sketchFeatures.Count);

            for (int i = sketchFeatures.Count - 1; i >= 0; i--)
            {
                ISketch sketch = TryGetSketch(sketchFeatures[i]);
                string skName = sketchFeatures[i].Name;
                if (sketch == null) { log.AppendFormat("  [{0}] skipped — TryGetSketch returned null\n", skName); continue; }

                object[] segs = (object[])sketch.GetSketchSegments();
                if (segs == null) { log.AppendFormat("  [{0}] skipped — no segments\n", skName); continue; }

                MathTransform skToModel = (MathTransform)sketch.ModelToSketchTransform.Inverse();
                MathTransform modelToSketch = sketch.ModelToSketchTransform;

                // Bring body axis into sketch space for direction comparison (API-correct, no manual matrix math)
                IMathVector axisVecSk = (IMathVector)mathUtil.CreateVector(normAxis);
                axisVecSk = (IMathVector)axisVecSk.MultiplyTransform(modelToSketch);
                double[] axisInSketch = Normalize((double[])axisVecSk.ArrayData);

                int lineIdx = 0;

                foreach (object seg in segs)
                {
                    ISketchLine testLine = null;
                    try { testLine = seg as ISketchLine; } catch { }
                    if (testLine == null) continue;

                    lineIdx++;
                    ISketchPoint sp = (ISketchPoint)testLine.GetStartPoint2();
                    ISketchPoint ep = (ISketchPoint)testLine.GetEndPoint2();

                    double[] sketchDir = Normalize(new double[] {
                        ep.X - sp.X, ep.Y - sp.Y, ep.Z - sp.Z });

                    // Compare in sketch space — avoids all inverse/convention issues
                    double dot = Math.Abs(
                        axisInSketch[0]*sketchDir[0] + axisInSketch[1]*sketchDir[1] + axisInSketch[2]*sketchDir[2]);

                    string dirInfo = string.Format(
                        "sketchDir=({0:F3},{1:F3},{2:F3}) axisInSketch=({3:F3},{4:F3},{5:F3})",
                        sketchDir[0], sketchDir[1], sketchDir[2],
                        axisInSketch[0], axisInSketch[1], axisInSketch[2]);

                    if (dot <= 0.99)
                    {
                        log.AppendFormat("  [{0}] line{1}: FAIL direction  dot={2:F4}  {3}\n", skName, lineIdx, dot, dirInfo);
                        continue;
                    }

                    if (bodyAxisOrigin == null)
                    {
                        log.AppendFormat("  [{0}] line{1}: direction OK (dot={2:F4}) but no axis origin — cannot proximity check  {3}\n", skName, lineIdx, dot, dirInfo);
                    }
                    else
                    {
                        // Get model-space line direction via API (not manual matrix)
                        IMathVector lineDirVec = (IMathVector)mathUtil.CreateVector(sketchDir);
                        lineDirVec = (IMathVector)lineDirVec.MultiplyTransform(skToModel);
                        double[] modelDir = Normalize((double[])lineDirVec.ArrayData);

                        double dist = ComputeLinePointDistance(mathUtil, skToModel,
                            new double[] { sp.X, sp.Y, sp.Z }, modelDir, bodyAxisOrigin);
                        if (dist < 0.001)
                        {
                            log.AppendFormat("  [{0}] line{1}: MATCH  dot={2:F4}  dist={3:F4}m  {4}\n", skName, lineIdx, dot, dist, dirInfo);
#if DEBUG
                            MessageBox.Show(log.ToString(), "FindMatchingSketchLine", MessageBoxButtons.OK, MessageBoxIcon.Information);
#endif
                            return testLine;
                        }
                        log.AppendFormat("  [{0}] line{1}: direction OK (dot={2:F4}) FAIL proximity  dist={3:F4}m  {4}\n", skName, lineIdx, dot, dist, dirInfo);
                    }

                    if (directionOnlyMatch == null)
                        directionOnlyMatch = testLine;
                }
            }

#if DEBUG
            MessageBox.Show(log.ToString(), "FindMatchingSketchLine — no match", MessageBoxButtons.OK, MessageBoxIcon.Warning);
#endif
            return null;
        }

        /// <summary>
        /// Returns the distance from targetModelPt to the infinite line defined by
        /// sketchSpaceStart (transformed to model space) and modelDir.
        /// </summary>
        private static double ComputeLinePointDistance(
            IMathUtility mathUtil,
            MathTransform skToModel,
            double[] sketchSpaceStart,
            double[] modelDir,
            double[] targetModelPt)
        {
            IMathPoint startMath = (IMathPoint)mathUtil.CreatePoint(sketchSpaceStart);
            startMath = (IMathPoint)startMath.MultiplyTransform(skToModel);
            double[] startModel = (double[])startMath.ArrayData;

            double[] w = new double[] {
                targetModelPt[0] - startModel[0],
                targetModelPt[1] - startModel[1],
                targetModelPt[2] - startModel[2] };

            double[] cross = new double[] {
                w[1]*modelDir[2] - w[2]*modelDir[1],
                w[2]*modelDir[0] - w[0]*modelDir[2],
                w[0]*modelDir[1] - w[1]*modelDir[0] };

            return Math.Sqrt(cross[0]*cross[0] + cross[1]*cross[1] + cross[2]*cross[2]);
        }

        /// <summary>
        /// Transforms a direction vector by a MathTransform (rotation only, ignores translation).
        /// </summary>
        private static double[] TransformDirection(double[] dir, MathTransform xform)
        {
            double[] xformData = (double[])xform.ArrayData;
            // The 3x3 rotation part is in elements [0..8] (row-major)
            double rx = xformData[0] * dir[0] + xformData[1] * dir[1] + xformData[2] * dir[2];
            double ry = xformData[3] * dir[0] + xformData[4] * dir[1] + xformData[5] * dir[2];
            double rz = xformData[6] * dir[0] + xformData[7] * dir[1] + xformData[8] * dir[2];
            return new double[] { rx, ry, rz };
        }

        /// <summary>
        /// Returns true if the infinite line (defined by sketchSpaceStart + modelDir, with
        /// sketchSpaceStart transformed to model space via IMathUtility) passes within 1 mm
        /// of targetModelPt.  Uses IMathUtility for the point transform so it is guaranteed
        /// to match SolidWorks' own convention.
        /// </summary>
        private static bool LinePassesNearPoint(
            IMathUtility mathUtil,
            MathTransform skToModel,
            double[] sketchSpaceStart,
            double[] modelDir,
            double[] targetModelPt)
        {
            return ComputeLinePointDistance(mathUtil, skToModel,
                sketchSpaceStart, modelDir, targetModelPt) < 0.001;
        }

        private static bool FeatureOwnsBody(IFeature feat, IBody2 targetBody)
        {
            object[] faces = (object[])feat.GetFaces();
            if (faces == null) return false;

            foreach (object f in faces)
            {
                IFace2 face = (IFace2)f;
                IBody2 body = (IBody2)face.GetBody();
                if (body != null && body.Name == targetBody.Name)
                    return true;
            }
            return false;
        }

        private static ISketch TryGetSketch(IFeature feature)
        {
            try
            {
                object specific = feature.GetSpecificFeature2();
                if (specific == null) return null;
                return specific as ISketch;
            }
            catch { return null; }
        }

        private static ISketchLine ExtractSketchLine(ISketch sketch)
        {
            object[] segs = (object[])sketch.GetSketchSegments();
            if (segs == null) return null;

            foreach (object seg in segs)
            {
                try
                {
                    ISketchLine testLine = seg as ISketchLine;
                    if (testLine != null)
                        return testLine;
                }
                catch { }
            }

            return null;
        }

        private static ISketchLine FindPathSketchLine(IModelDoc2 doc, IFeature structFeat)
        {
            IFeature feat = (IFeature)doc.FirstFeature();
            List<IFeature> sketchFeatures = new List<IFeature>();

            while (feat != null)
            {
                if (feat.Name == structFeat.Name)
                    break;

                string typeName = feat.GetTypeName2();
                if (typeName == "ProfileFeature" || typeName == "3DProfileFeature")
                    sketchFeatures.Add(feat);

                feat = (IFeature)feat.GetNextFeature();
            }

            for (int i = sketchFeatures.Count - 1; i >= 0; i--)
            {
                ISketch sketch = TryGetSketch(sketchFeatures[i]);
                if (sketch != null)
                {
                    ISketchLine line = ExtractSketchLine(sketch);
                    if (line != null)
                        return line;
                }
            }

            return null;
        }

        /// <summary>
        /// Searches the structural member's sub-feature tree for the sketch line whose
        /// direction and spatial position match this body's cylinder axis.
        /// Falls back to a direction-only match if no origin is available.
        /// </summary>
        private static ISketchLine FindSketchLineInSubFeatures(
            SldWorks swApp, IFeature structFeat, double[] bodyAxisDir, double[] bodyAxisOrigin)
        {
            double[] normAxis = bodyAxisDir != null ? Normalize(bodyAxisDir) : null;
            IMathUtility mathUtil = (IMathUtility)swApp.GetMathUtility();
            ISketchLine directionOnlyMatch = null;

            var log = new System.Text.StringBuilder();
            log.AppendFormat("FindSketchLineInSubFeatures — structural member: {0}\n\n", structFeat.Name);

            IFeature subFeat = (IFeature)structFeat.GetFirstSubFeature();
            while (subFeat != null)
            {
                string subType = subFeat.GetTypeName2();
                log.AppendFormat("  SubFeature: '{0}'  type: {1}\n", subFeat.Name, subType);

                if (subType == "ProfileFeature" || subType == "3DProfileFeature")
                {
                    ISketch sketch = TryGetSketch(subFeat);
                    if (sketch == null) { log.AppendLine("    -> TryGetSketch returned null"); }
                    else
                    {
                        object[] segs = (object[])sketch.GetSketchSegments();
                        int segCount = segs != null ? segs.Length : 0;
                        log.AppendFormat("    -> sketch has {0} segments\n", segCount);

                        if (segs != null)
                        {
                            MathTransform skToModel =
                                (MathTransform)sketch.ModelToSketchTransform.Inverse();
                            MathTransform modelToSketch = sketch.ModelToSketchTransform;

                            // Bring body axis into sketch space for direction comparison (API-correct)
                            double[] axisInSketch = null;
                            if (normAxis != null)
                            {
                                IMathVector axisVecSk = (IMathVector)mathUtil.CreateVector(normAxis);
                                axisVecSk = (IMathVector)axisVecSk.MultiplyTransform(modelToSketch);
                                axisInSketch = Normalize((double[])axisVecSk.ArrayData);
                            }

                            int lineIdx = 0;

                            foreach (object seg in segs)
                            {
                                ISketchLine testLine = null;
                                try { testLine = seg as ISketchLine; } catch { }
                                if (testLine == null) { log.AppendFormat("    seg {0}: not an ISketchLine ({1})\n", lineIdx+1, seg != null ? seg.GetType().Name : "null"); lineIdx++; continue; }

                                lineIdx++;
                                ISketchPoint sp = (ISketchPoint)testLine.GetStartPoint2();
                                ISketchPoint ep = (ISketchPoint)testLine.GetEndPoint2();

                                if (axisInSketch == null)
                                {
                                    log.AppendLine("    -> normAxis null, returning first line");
#if DEBUG
                                    MessageBox.Show(log.ToString(), "FindSketchLineInSubFeatures", MessageBoxButtons.OK, MessageBoxIcon.Information);
#endif
                                    return testLine;
                                }

                                double[] skDir = Normalize(new double[] {
                                    ep.X - sp.X, ep.Y - sp.Y, ep.Z - sp.Z });

                                // Compare in sketch space — avoids inverse/convention issues
                                double dot = Math.Abs(
                                    axisInSketch[0]*skDir[0] +
                                    axisInSketch[1]*skDir[1] +
                                    axisInSketch[2]*skDir[2]);

                                if (dot <= 0.99) { log.AppendFormat("    line{0}: FAIL dir dot={1:F4}  sk=({2:F3},{3:F3},{4:F3}) axisInSk=({5:F3},{6:F3},{7:F3})\n", lineIdx, dot, skDir[0], skDir[1], skDir[2], axisInSketch[0], axisInSketch[1], axisInSketch[2]); continue; }

                                if (directionOnlyMatch == null)
                                    directionOnlyMatch = testLine;

                                if (bodyAxisOrigin != null)
                                {
                                    // Get model-space line direction via API (not manual matrix)
                                    IMathVector lineDirVec = (IMathVector)mathUtil.CreateVector(skDir);
                                    lineDirVec = (IMathVector)lineDirVec.MultiplyTransform(skToModel);
                                    double[] modelDir = Normalize((double[])lineDirVec.ArrayData);

                                    double dist = ComputeLinePointDistance(mathUtil, skToModel,
                                        new double[] { sp.X, sp.Y, sp.Z }, modelDir, bodyAxisOrigin);
                                    if (dist < 0.001)
                                    {
                                        log.AppendFormat("    line{0}: MATCH dot={1:F4} dist={2:F4}m\n", lineIdx, dot, dist);
#if DEBUG
                                        MessageBox.Show(log.ToString(), "FindSketchLineInSubFeatures", MessageBoxButtons.OK, MessageBoxIcon.Information);
#endif
                                        return testLine;
                                    }
                                    log.AppendFormat("    line{0}: dir OK dot={1:F4} FAIL proximity dist={2:F4}m\n", lineIdx, dot, dist);
                                }
                            }
                        }
                    }
                }
                subFeat = (IFeature)subFeat.GetNextSubFeature();
            }

#if DEBUG
            MessageBox.Show(log.ToString(), "FindSketchLineInSubFeatures — no match", MessageBoxButtons.OK, MessageBoxIcon.Warning);
#endif
            return directionOnlyMatch;
        }

        private static ISketchLine TryGetPathFromDefinition(IModelDoc2 doc, IFeature structFeat)
        {
            try
            {
                object defObj = structFeat.GetDefinition();
                if (defObj == null) return null;

                dynamic smData = defObj;
                smData.AccessSelections((ModelDoc2)doc, null);

                try
                {
                    int groupCount = (int)smData.GroupCount;
                    for (int g = 0; g < groupCount; g++)
                    {
                        object segArray = smData.GetSegments(g);
                        if (segArray == null) continue;

                        object[] segments = (object[])segArray;
                        foreach (object seg in segments)
                        {
                            try
                            {
                                ISketchLine line = seg as ISketchLine;
                                if (line != null)
                                {
                                    smData.ReleaseSelectionAccess();
                                    return line;
                                }
                            }
                            catch { }
                        }
                    }
                }
                finally
                {
                    try { smData.ReleaseSelectionAccess(); } catch { }
                }
            }
            catch { }

            return null;
        }

        private static void ExtractTubeProfile(
            IBody2 targetBody,
            out double outerRadius,
            out double wallThickness)
        {
            outerRadius = 0;
            wallThickness = 0;

            object[] faces = (object[])targetBody.GetFaces();
            if (faces == null) return;

            double maxRadius = 0;
            double minRadius = double.MaxValue;

            foreach (object f in faces)
            {
                IFace2   face = (IFace2)f;
                ISurface surf = (ISurface)face.GetSurface();

                if (surf.IsCylinder())
                {
                    double[] cylParams = (double[])surf.CylinderParams;
                    if (cylParams != null && cylParams.Length >= 7)
                    {
                        double radius = cylParams[6];
                        if (radius > maxRadius) maxRadius = radius;
                        if (radius < minRadius) minRadius = radius;
                    }
                }
            }

            if (maxRadius > 0 && minRadius < double.MaxValue && minRadius > 0)
            {
                outerRadius = maxRadius;
                wallThickness = maxRadius - minRadius;
            }
            else if (maxRadius > 0)
            {
                outerRadius = maxRadius;
                wallThickness = 0.002;
            }
        }

        private static void ValidateDistinctSketchLines(ISketchLine tabLine, ISketchLine slotLine)
        {
            ISketchSegment tabSeg  = (ISketchSegment)tabLine;
            ISketchSegment slotSeg = (ISketchSegment)slotLine;

            ISketch tabSketch  = tabSeg.GetSketch() as ISketch;
            ISketch slotSketch = slotSeg.GetSketch() as ISketch;

            string tabSkName  = tabSketch  != null ? ((IFeature)tabSketch).Name  : "";
            string slotSkName = slotSketch != null ? ((IFeature)slotSketch).Name : "";

            int[] tabIdArr  = tabSeg.GetID()  as int[];
            int[] slotIdArr = slotSeg.GetID() as int[];
            int tabId  = (tabIdArr  != null && tabIdArr.Length  > 0) ? tabIdArr[0]  : -1;
            int slotId = (slotIdArr != null && slotIdArr.Length > 0) ? slotIdArr[0] : -2;

            if (tabSkName == slotSkName && tabId == slotId)
                throw new InvalidOperationException(string.Format(
                    "Tab and slot resolved to the same sketch line (sketch '{0}', segment ID {1}). " +
                    "Make sure you selected two different weldment bodies.",
                    tabSkName, tabId));
        }

        private static double[] Normalize(double[] v)
        {
            double len = Math.Sqrt(v[0] * v[0] + v[1] * v[1] + v[2] * v[2]);
            if (len < 1e-15) return new double[] { 0, 0, 0 };
            return new double[] { v[0] / len, v[1] / len, v[2] / len };
        }
    }
}
