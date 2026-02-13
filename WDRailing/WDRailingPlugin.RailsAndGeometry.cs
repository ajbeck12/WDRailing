using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Reflection;

using Tekla.Structures;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;
using Tekla.Structures.Model.Operations;
using Tekla.Structures.Model.UI;
using Tekla.Structures.Plugins;


namespace WDRailing
{
    public partial class WDRailingPlugin
    {
        private void CreateRails(
            Point startOnLine, Point endOnLine,
            Vector dirUnit, Vector leftUnit,
            double postLineLateralMm,
            double halfPostWidthMm,
            double firstPostTopZ, double lastPostTopZ,
            Part anyHostForSide,
            double railStartOffsetMm, double railEndOffsetMm,
            double railFromTopMm,
            int railCount, double railSpacingMm)
        {
            const string railProfile = "TS1-1/2X1-1/2X.188";
            const string railMaterial = "A53";
            const string railClass = "1";
            const string railName = "RAIL";

            double maxLenMm = InchesToMm(240.0); // 20'-0"

            // half rail width (OD 1.5" fallback)
            double halfRailWidthMm = InchesToMm(1.5) * 0.5;
            if (TryGetOutsideDimMm(railProfile, out var railOutsideMm))
                halfRailWidthMm = railOutsideMm * 0.5;

            int sideSign = DetermineConnectionSideSign(leftUnit, startOnLine, anyHostForSide);
            if (sideSign == 0) sideSign = +1;

            // move to the post face on connection side, then out by half rail
            double railLateralMm = postLineLateralMm + sideSign * (halfPostWidthMm + halfRailWidthMm);

            // positive offsets EXTEND outward from both ends
            Point sLine = new Point(
                startOnLine.X - dirUnit.X * railStartOffsetMm,
                startOnLine.Y - dirUnit.Y * railStartOffsetMm,
                startOnLine.Z - dirUnit.Z * railStartOffsetMm);

            Point eLine = new Point(
                endOnLine.X + dirUnit.X * railEndOffsetMm,
                endOnLine.Y + dirUnit.Y * railEndOffsetMm,
                endOnLine.Z + dirUnit.Z * railEndOffsetMm);

            // lateral shift to connection-side face
            Point s = new Point(sLine.X + leftUnit.X * railLateralMm, sLine.Y + leftUnit.Y * railLateralMm, sLine.Z);
            Point e = new Point(eLine.X + leftUnit.X * railLateralMm, eLine.Y + leftUnit.Y * railLateralMm, eLine.Z);

            for (int r = 0; r < railCount; r++)
            {
                double zStart = firstPostTopZ - railFromTopMm - (r * railSpacingMm);
                double zEnd = lastPostTopZ - railFromTopMm - (r * railSpacingMm);

                Point a = new Point(s.X, s.Y, zStart);
                Point b = new Point(e.X, e.Y, zEnd);

                CreateRailPieces(a, b, maxLenMm, railProfile, railMaterial, railClass, railName);
            }
        }


        /// <summary>
        /// Builds rails for all sides of a multi-sided pick with automatic corner behavior:
        /// - Each interior corner: next side butts into previous side (no clipping).
        /// - Non-butt ends receive end caps.
        /// - Adds a corner seat angle (slots only, no pilot holes) at each interior corner and rail row.
        /// </summary>
        private void CreateCornerAwareRailsForPolyline(
            List<RailRunSegment> segments,
            double railStartOffsetMm,
            double railEndOffsetMm,
            double railFromTopMm,
            int railCount,
            double railSpacingMm,
            double seatHoleLineFromBendIn,
            double seatSlotC2CIn,
            double seatSlotSizeIn,
            string seatSlotStandard,
            double seatSlotCutLengthIn,
            bool seatSlotSpecialFirstLayer)
        {
            if (segments == null || segments.Count == 0 || railCount <= 0) return;

            const string railProfile = "TS1-1/2X1-1/2X.188";
            const string railMaterial = "A53";
            const string railClass = "1";
            const string railName = "RAIL";

            double maxLenMm = InchesToMm(240.0); // 20'-0"

            double railOutsideMm = InchesToMm(1.5);
            if (TryGetOutsideDimMm(railProfile, out var railOutsideMmParsed))
                railOutsideMm = railOutsideMmParsed;
            double halfRailWidthMm = railOutsideMm * 0.5;

            int sideCount = segments.Count;

            var startPts = new Point[sideCount, railCount];
            var endPts = new Point[sideCount, railCount];

            // Build base centerlines per side/row before corner trimming.
            for (int s = 0; s < sideCount; s++)
            {
                RailRunSegment seg = segments[s];

                int sideSign = DetermineConnectionSideSign(seg.Left, seg.FirstStationOnLine, seg.AnyHost);
                if (sideSign == 0) sideSign = +1;

                double railLateralMm = seg.LateralOffsetMm + sideSign * (seg.HalfPostWidthMm + halfRailWidthMm);

                Point sLine = new Point(
                    seg.FirstStationOnLine.X - seg.Dir.X * railStartOffsetMm,
                    seg.FirstStationOnLine.Y - seg.Dir.Y * railStartOffsetMm,
                    seg.FirstStationOnLine.Z - seg.Dir.Z * railStartOffsetMm);

                Point eLine = new Point(
                    seg.LastStationOnLine.X + seg.Dir.X * railEndOffsetMm,
                    seg.LastStationOnLine.Y + seg.Dir.Y * railEndOffsetMm,
                    seg.LastStationOnLine.Z + seg.Dir.Z * railEndOffsetMm);

                Point sBase = new Point(
                    sLine.X + seg.Left.X * railLateralMm,
                    sLine.Y + seg.Left.Y * railLateralMm,
                    sLine.Z);

                Point eBase = new Point(
                    eLine.X + seg.Left.X * railLateralMm,
                    eLine.Y + seg.Left.Y * railLateralMm,
                    eLine.Z);

                for (int r = 0; r < railCount; r++)
                {
                    double zStart = seg.FirstPostTopZ - railFromTopMm - (r * railSpacingMm);
                    double zEnd = seg.LastPostTopZ - railFromTopMm - (r * railSpacingMm);

                    startPts[s, r] = new Point(sBase.X, sBase.Y, zStart);
                    endPts[s, r] = new Point(eBase.X, eBase.Y, zEnd);
                }
            }

            // Cap plan: non-butt ends only.
            // Convention: at every interior corner, next side butts into previous side.
            var capAtStart = new bool[sideCount, railCount];
            var capAtEnd = new bool[sideCount, railCount];

            // Open start/end are always non-butt.
            for (int r = 0; r < railCount; r++)
            {
                capAtStart[0, r] = true;
                capAtEnd[sideCount - 1, r] = true;
            }

            // Corner seat records
            var cornerSeats = new List<CornerSeatSpec>();

            for (int c = 0; c < sideCount - 1; c++)
            {
                int a = c;
                int b = c + 1;

                for (int r = 0; r < railCount; r++)
                {
                    Point aStart = startPts[a, r];
                    Point aEnd = endPts[a, r];
                    Point bStart = startPts[b, r];
                    Point bEnd = endPts[b, r];

                    Vector dirA = UnitVector(aStart, aEnd);
                    Vector dirB = UnitVector(bStart, bEnd);

                    // If nearly colinear, no corner treatment here.
                    double turn = Math.Abs(dirA.X * dirB.X + dirA.Y * dirB.Y + dirA.Z * dirB.Z);
                    if (turn > 0.999)
                        continue;

                    // Butt side B into side A and trim/extend B start so it lands on A's outer face plane.
                    Point newBStart = ComputeButtStartToSideFace(aEnd, dirA, bStart, dirB, halfRailWidthMm);
                    startPts[b, r] = newBStart;

                    // Non-butt end at this corner is side A end -> cap it.
                    capAtEnd[a, r] = true;

                    // Add corner seat at interior corner (slots only).
                    cornerSeats.Add(new CornerSeatSpec
                    {
                        ContactAEnd = aEnd,
                        ButtBStart = newBStart,
                        DirA = dirA,
                        DirB = dirB
                    });
                }
            }

            // Build rails and keep first/last piece handle so we can cap start/end pieces.
            var firstPiece = new Beam[sideCount, railCount];
            var lastPiece = new Beam[sideCount, railCount];

            for (int s = 0; s < sideCount; s++)
            {
                for (int r = 0; r < railCount; r++)
                {
                    Point a = startPts[s, r];
                    Point b = endPts[s, r];

                    if (Distance3D(a, b) < 1.0) continue;

                    List<Beam> pieces = CreateRailPiecesCollect(a, b, maxLenMm, railProfile, railMaterial, railClass, railName);
                    if (pieces.Count == 0) continue;

                    firstPiece[s, r] = pieces[0];
                    lastPiece[s, r] = pieces[pieces.Count - 1];
                }
            }

            // Apply non-butt end caps.
            for (int s = 0; s < sideCount; s++)
            {
                for (int r = 0; r < railCount; r++)
                {
                    Beam f = firstPiece[s, r];
                    Beam l = lastPiece[s, r];

                    if (capAtStart[s, r] && f != null)
                        CreateRailEndCap(f, atStart: true, outsideMm: railOutsideMm);

                    if (capAtEnd[s, r] && l != null)
                        CreateRailEndCap(l, atStart: false, outsideMm: railOutsideMm);
                }
            }

            // Add interior corner seat angles (slots only, no pilot holes).
            for (int i = 0; i < cornerSeats.Count; i++)
            {
                CornerSeatSpec cs = cornerSeats[i];
                CreateCornerSeatAngleSlotsOnly(
                    cs.ContactAEnd,
                    cs.ButtBStart,
                    cs.DirA,
                    cs.DirB,
                    seatHoleLineFromBendIn,
                    seatSlotC2CIn,
                    seatSlotSizeIn,
                    seatSlotStandard,
                    seatSlotCutLengthIn,
                    seatSlotSpecialFirstLayer);
            }
        }


        private sealed class CornerSeatSpec
        {
            public Point ContactAEnd;
            public Point ButtBStart;
            public Vector DirA;
            public Vector DirB;
        }


        private static Point ComputeButtStartToSideFace(
            Point aEnd,
            Vector dirA,
            Point bStart,
            Vector dirB,
            double halfRailWidthMm)
        {
            // Direction from A centerline toward B side near the corner.
            Vector toB = new Vector(bStart.X - aEnd.X, bStart.Y - aEnd.Y, bStart.Z - aEnd.Z);

            double dotAB = toB.X * dirA.X + toB.Y * dirA.Y + toB.Z * dirA.Z;
            Vector n = new Vector(
                toB.X - dotAB * dirA.X,
                toB.Y - dotAB * dirA.Y,
                toB.Z - dotAB * dirA.Z);

            double nLen = Math.Sqrt(n.X * n.X + n.Y * n.Y + n.Z * n.Z);
            if (nLen < 1e-6)
            {
                // Fallback normal if geometry is very tight/noisy.
                Vector up = new Vector(0.0, 0.0, 1.0);
                n = Cross(up, dirA);
                nLen = Math.Sqrt(n.X * n.X + n.Y * n.Y + n.Z * n.Z);
                if (nLen < 1e-6)
                    n = new Vector(1.0, 0.0, 0.0);
                else
                    n = new Vector(n.X / nLen, n.Y / nLen, n.Z / nLen);
            }
            else
            {
                n = new Vector(n.X / nLen, n.Y / nLen, n.Z / nLen);
            }

            Point planePoint = new Point(
                aEnd.X + n.X * halfRailWidthMm,
                aEnd.Y + n.Y * halfRailWidthMm,
                aEnd.Z + n.Z * halfRailWidthMm);

            Vector bp = new Vector(
                planePoint.X - bStart.X,
                planePoint.Y - bStart.Y,
                planePoint.Z - bStart.Z);

            double denom = dirB.X * n.X + dirB.Y * n.Y + dirB.Z * n.Z;

            double t;
            if (Math.Abs(denom) < 1e-6)
            {
                t = halfRailWidthMm;
            }
            else
            {
                t = (bp.X * n.X + bp.Y * n.Y + bp.Z * n.Z) / denom;
            }

            // Clamp to avoid pathological long jumps on odd geometry.
            double maxShift = halfRailWidthMm * 4.0;
            if (t > maxShift) t = maxShift;
            if (t < -maxShift) t = -maxShift;

            return new Point(
                bStart.X + dirB.X * t,
                bStart.Y + dirB.Y * t,
                bStart.Z + dirB.Z * t);
        }


        private static Vector UnitVector(Point a, Point b)
        {
            Vector v = new Vector(b.X - a.X, b.Y - a.Y, b.Z - a.Z);
            double len = Math.Sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z);
            if (len < 1e-9) return new Vector(1.0, 0.0, 0.0);
            return new Vector(v.X / len, v.Y / len, v.Z / len);
        }


        private static double Distance3D(Point a, Point b)
        {
            double dx = a.X - b.X;
            double dy = a.Y - b.Y;
            double dz = a.Z - b.Z;
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }


        private static Vector Cross(Vector a, Vector b)
        {
            return new Vector(
                a.Y * b.Z - a.Z * b.Y,
                a.Z * b.X - a.X * b.Z,
                a.X * b.Y - a.Y * b.X);
        }


        private static ContourPlate CreateRailEndCap(Beam rail, bool atStart, double outsideMm)
        {
            try
            {
                if (rail == null) return null;

                if (outsideMm < 0.1) outsideMm = InchesToMm(1.5);
                double half = outsideMm * 0.5;

                double tMm = InchesToMm(0.125); // 1/8"

                Point p0 = rail.StartPoint;
                Point p1 = rail.EndPoint;

                Vector u = UnitVector(p0, p1);
                if (atStart) u = new Vector(-u.X, -u.Y, -u.Z);

                Vector refV = (Math.Abs(u.Z) < 0.95)
                    ? new Vector(0.0, 0.0, 1.0)
                    : new Vector(1.0, 0.0, 0.0);

                Vector v = Cross(refV, u);
                double vLen = Math.Sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z);
                if (vLen < 1e-9) return null;
                v = new Vector(v.X / vLen, v.Y / vLen, v.Z / vLen);

                Vector w = Cross(u, v);
                double wLen = Math.Sqrt(w.X * w.X + w.Y * w.Y + w.Z * w.Z);
                if (wLen < 1e-9) return null;
                w = new Vector(w.X / wLen, w.Y / wLen, w.Z / wLen);

                Point end = atStart ? rail.StartPoint : rail.EndPoint;

                // Push slightly outward so cap sits on the exposed end face.
                Point c = new Point(
                    end.X + u.X * (tMm * 0.5),
                    end.Y + u.Y * (tMm * 0.5),
                    end.Z + u.Z * (tMm * 0.5));

                Point q1 = new Point(c.X + v.X * half + w.X * half, c.Y + v.Y * half + w.Y * half, c.Z + v.Z * half + w.Z * half);
                Point q2 = new Point(c.X - v.X * half + w.X * half, c.Y - v.Y * half + w.Y * half, c.Z - v.Z * half + w.Z * half);
                Point q3 = new Point(c.X - v.X * half - w.X * half, c.Y - v.Y * half - w.Y * half, c.Z - v.Z * half - w.Z * half);
                Point q4 = new Point(c.X + v.X * half - w.X * half, c.Y + v.Y * half - w.Y * half, c.Z + v.Z * half - w.Z * half);

                var cap = new ContourPlate();
                cap.Name = "RAIL CAP";
                cap.Class = "4";
                cap.Material.MaterialString = "A36";
                cap.Profile.ProfileString = "PL3.175";

                cap.Position.Plane = Position.PlaneEnum.MIDDLE;
                cap.Position.Rotation = Position.RotationEnum.TOP;
                cap.Position.Depth = Position.DepthEnum.MIDDLE;

                cap.Contour.AddContourPoint(new ContourPoint(q1, null));
                cap.Contour.AddContourPoint(new ContourPoint(q2, null));
                cap.Contour.AddContourPoint(new ContourPoint(q3, null));
                cap.Contour.AddContourPoint(new ContourPoint(q4, null));

                if (cap.Insert()) return cap;
            }
            catch
            {
                // Best effort only.
            }

            return null;
        }


        private static void CreateRailPieces(
            Point a, Point b, double maxLenMm,
            string profile, string material, string cls, string name)
        {
            Vector v = new Vector(b.X - a.X, b.Y - a.Y, b.Z - a.Z);
            double total = Math.Sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z);
            if (total < 1.0) return;

            int pieces = Math.Max(1, (int)Math.Ceiling(total / maxLenMm));

            for (int i = 0; i < pieces; i++)
            {
                double t0 = (double)i / pieces;
                double t1 = (double)(i + 1) / pieces;

                Point p0 = new Point(a.X + v.X * t0, a.Y + v.Y * t0, a.Z + v.Z * t0);
                Point p1 = new Point(a.X + v.X * t1, a.Y + v.Y * t1, a.Z + v.Z * t1);

                var rail = new Beam(p0, p1);
                rail.Name = name;
                rail.Class = cls;
                rail.Profile.ProfileString = profile;
                rail.Material.MaterialString = material;

                rail.Position.Plane = Position.PlaneEnum.MIDDLE;
                rail.Position.Rotation = Position.RotationEnum.TOP;
                rail.Position.Depth = Position.DepthEnum.MIDDLE;

                rail.Insert();
            }
        }


        private static List<Beam> CreateRailPiecesCollect(
            Point a, Point b, double maxLenMm,
            string profile, string material, string cls, string name)
        {
            var created = new List<Beam>();

            Vector v = new Vector(b.X - a.X, b.Y - a.Y, b.Z - a.Z);
            double total = Math.Sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z);
            if (total < 1.0) return created;

            int pieces = Math.Max(1, (int)Math.Ceiling(total / maxLenMm));

            for (int i = 0; i < pieces; i++)
            {
                double t0 = (double)i / pieces;
                double t1 = (double)(i + 1) / pieces;

                Point p0 = new Point(a.X + v.X * t0, a.Y + v.Y * t0, a.Z + v.Z * t0);
                Point p1 = new Point(a.X + v.X * t1, a.Y + v.Y * t1, a.Z + v.Z * t1);

                var rail = new Beam(p0, p1);
                rail.Name = name;
                rail.Class = cls;
                rail.Profile.ProfileString = profile;
                rail.Material.MaterialString = material;

                rail.Position.Plane = Position.PlaneEnum.MIDDLE;
                rail.Position.Rotation = Position.RotationEnum.TOP;
                rail.Position.Depth = Position.DepthEnum.MIDDLE;

                if (rail.Insert())
                    created.Add(rail);
            }

            return created;
        }


        private static int DetermineConnectionSideSign(Vector leftUnit, Point refPointOnLine, Part host)
        {
            if (host == null || refPointOnLine == null) return 0;

            try
            {
                Solid s = host.GetSolid();
                if (s == null) return 0;

                double hx = 0.5 * (s.MinimumPoint.X + s.MaximumPoint.X);
                double hy = 0.5 * (s.MinimumPoint.Y + s.MaximumPoint.Y);

                double vx = hx - refPointOnLine.X;
                double vy = hy - refPointOnLine.Y;

                double dot = vx * leftUnit.X + vy * leftUnit.Y;
                return (dot >= 0.0) ? +1 : -1;
            }
            catch { return 0; }
        }



        // ---------------- Post creation ----------------


        private static Beam CreatePost(Point start, Point end, string profile, string material, string cls, string name, Position.RotationEnum rotation)
        {
            var post = new Beam(start, end);
            post.Name = name;
            post.Class = cls;
            post.Profile.ProfileString = profile;
            post.Material.MaterialString = material;

            post.Position.Plane = Position.PlaneEnum.MIDDLE;
            post.Position.Rotation = rotation;
            post.Position.Depth = Position.DepthEnum.MIDDLE;

            if (!post.Insert())
                return null;

            return post;
        }


        private void RunOneSegment(
    Point p1, Point p2,
    bool skipFirstCornerPost,
    List<Part> hostParts,
    double spacingMm, double postHeightMm,
    double startOffsetMm, double endOffsetMm,
    double baseOffsetMm, double deckEdgeMm, string lineRef,
    string profile, string material, string postClass, string postName,
    bool connEnabled, string connName, string connAttr,
    bool railEnabled, int railCount,
    double railStartOffsetMm, double railEndOffsetMm,
    double railFromTopMm, double railSpacingMm,
    ref int inserted, ref int failed, ref int connOk, ref int connFail)
        {
            Vector run = new Vector(p2.X - p1.X, p2.Y - p1.Y, p2.Z - p1.Z);
            double runLen = Math.Sqrt(run.X * run.X + run.Y * run.Y + run.Z * run.Z);
            if (runLen < 1.0) return;

            Vector dir = new Vector(run.X / runLen, run.Y / runLen, run.Z / runLen);
            Vector left = GetLeftVectorXY(dir);
            Position.RotationEnum postRotation = GetPostRotationFromRun(dir);

            // Half-width used for face-based offsets (best effort)
            double halfPostWidthMm = 0.0;
            if (TryGetOutsideDimMm(profile, out double outsideMm)) halfPostWidthMm = outsideMm * 0.5;

            double lateralOffsetMm = ComputeLateralOffsetMm(lineRef, deckEdgeMm, halfPostWidthMm);

            double usableLen = runLen - startOffsetMm - endOffsetMm;
            if (usableLen <= 1.0) return;

            int segments = Math.Max(1, (int)Math.Ceiling(usableLen / spacingMm));
            double actualSpacingMm = usableLen / segments;

            Point firstStationOnLine = null;
            Point lastStationOnLine = null;
            double firstPostTopZ = 0.0;
            double lastPostTopZ = 0.0;
            Part firstHost = null;
            Part lastHost = null;

            for (int i = 0; i <= segments; i++)
            {
                if (skipFirstCornerPost && i == 0) continue;

                double d = startOffsetMm + i * actualSpacingMm;

                Point stationOnLine = new Point(
                    p1.X + dir.X * d,
                    p1.Y + dir.Y * d,
                    p1.Z + dir.Z * d);

                Point station = new Point(
                    stationOnLine.X + left.X * lateralOffsetMm,
                    stationOnLine.Y + left.Y * lateralOffsetMm,
                    stationOnLine.Z);

                Part nearestHost = null;
                if (hostParts != null && hostParts.Count > 0)
                    nearestHost = FindBestHostPartByXY(stationOnLine, hostParts);

                double baseZ = stationOnLine.Z + baseOffsetMm;
                if (nearestHost != null && TryGetPartTopZ(nearestHost, out double topZ))
                    baseZ = topZ + baseOffsetMm;

                Point postStart = new Point(station.X, station.Y, baseZ);
                Point postEnd = new Point(station.X, station.Y, baseZ + postHeightMm);

                Beam postObj = CreatePost(postStart, postEnd, profile, material, postClass, postName, postRotation);
                if (postObj == null) { failed++; continue; }

                inserted++;
                postObj.Modify();
                CreatePostCap(postObj);

                if (firstStationOnLine == null)
                {
                    firstStationOnLine = stationOnLine;
                    firstPostTopZ = postEnd.Z;
                    firstHost = nearestHost;
                }
                lastStationOnLine = stationOnLine;
                lastPostTopZ = postEnd.Z;
                lastHost = nearestHost;

                if (connEnabled && nearestHost != null)
                {
                    if (TryCreatePostConnection(nearestHost, postObj, connName, connAttr)) connOk++;
                    else connFail++;
                }
            }

            if (railEnabled && railCount > 0 && firstStationOnLine != null && lastStationOnLine != null)
            {
                CreateRails(
                    firstStationOnLine, lastStationOnLine,
                    dir, left,
                    lateralOffsetMm,
                    halfPostWidthMm,
                    firstPostTopZ, lastPostTopZ,
                    (firstHost ?? lastHost),
                    railStartOffsetMm, railEndOffsetMm,
                    railFromTopMm,
                    railCount,
                    railSpacingMm
                );
            }
        }


        // ---------------- Lateral offset helpers ----------------


        private static double ComputeLateralOffsetMm(string lineRef, double deckEdgeMm, double halfPostWidthMm)
        {
            int sideSign = 0; // +1 = LEFT, -1 = RIGHT, 0 = MIDDLE
            if (lineRef == "LEFT") sideSign = +1;
            else if (lineRef == "RIGHT") sideSign = -1;

            if (Math.Abs(deckEdgeMm) > 0.0001)
            {
                // Deck edge uses a side; if user left it at MIDDLE, default to LEFT.
                if (sideSign == 0) sideSign = +1;
                return sideSign * (deckEdgeMm + halfPostWidthMm);
            }

            return sideSign * halfPostWidthMm;
        }


        private static Vector GetLeftVectorXY(Vector dirUnit)
        {
            Vector left = new Vector(-dirUnit.Y, dirUnit.X, 0.0);
            double len = Math.Sqrt(left.X * left.X + left.Y * left.Y);
            if (len < 1e-9) return new Vector(1.0, 0.0, 0.0);
            return new Vector(left.X / len, left.Y / len, 0.0);
        }


        private static Position.RotationEnum GetPostRotationFromRun(Vector dirUnit)
        {
            // Choose dominant axis in XY
            if (Math.Abs(dirUnit.X) >= Math.Abs(dirUnit.Y))
            {
                // X-dominant
                return (dirUnit.X >= 0.0)
                    ? Position.RotationEnum.TOP
                    : Position.RotationEnum.BELOW;
            }
            else
            {
                // Y-dominant
                return (dirUnit.Y >= 0.0)
                    ? Position.RotationEnum.BACK
                    : Position.RotationEnum.FRONT;
            }
        }

        // ---------------- Profile parsing ----------------

    }
}
