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
        private struct SideRailInfo
        {
            public Point StartOnLine;
            public Point EndOnLine;
            public Vector Dir;
            public Vector Left;
            public double LateralOffsetMm;
            public double HalfPostWidthMm;
            public double FirstPostTopZ;
            public double LastPostTopZ;
            public Part Host;
            public double RailStartOffsetMm;
            public double RailEndOffsetMm;
            public bool CapStart;
            public bool CapEnd;
        }

        private struct CornerInfo
        {
            public Point CornerOnLine;
            public Vector IncomingDir;
            public Vector OutgoingDir;
            public Vector Left;
            public double LateralOffsetMm;
            public double HalfPostWidthMm;
            public int SideSign;
            public double CornerTopZ;
            public Part CornerPost;
        }

        private bool TryBuildButtCorner(
            SideRailInfo incoming,
            SideRailInfo outgoing,
            out SideRailInfo updatedIncoming,
            out SideRailInfo updatedOutgoing,
            out CornerInfo corner)
        {
            updatedIncoming = incoming;
            updatedOutgoing = outgoing;
            corner = new CornerInfo();

            if (incoming.EndOnLine == null || outgoing.StartOnLine == null)
                return false;

            if (!AreSamePoint(incoming.EndOnLine, outgoing.StartOnLine, 1.0))
                return false;

            Point cp = incoming.EndOnLine;

            Vector inDir = new Vector(incoming.Dir.X, incoming.Dir.Y, 0.0);
            Vector outDir = new Vector(outgoing.Dir.X, outgoing.Dir.Y, 0.0);
            double inLen = Math.Sqrt(inDir.X * inDir.X + inDir.Y * inDir.Y);
            double outLen = Math.Sqrt(outDir.X * outDir.X + outDir.Y * outDir.Y);
            if (inLen < 1e-9 || outLen < 1e-9) return false;
            inDir = new Vector(inDir.X / inLen, inDir.Y / inLen, 0.0);
            outDir = new Vector(outDir.X / outLen, outDir.Y / outLen, 0.0);

            // Skip near-colinear joints.
            double cross = (inDir.X * outDir.Y) - (inDir.Y * outDir.X);
            if (Math.Abs(cross) < 0.02) return false;

            int sideSign = DetermineConnectionSideSign(incoming.Left, cp, incoming.Host);
            if (sideSign == 0) sideSign = +1;

            double halfRailWidthMm = GetHalfRailWidthMm();
            double trimMm = halfRailWidthMm + incoming.HalfPostWidthMm;

            // Butt corner: trim incoming, extend/retract outgoing to the corner node.
            updatedIncoming.EndOnLine = new Point(
                cp.X - inDir.X * trimMm,
                cp.Y - inDir.Y * trimMm,
                cp.Z);
            updatedIncoming.RailEndOffsetMm = 0.0;
            updatedIncoming.CapEnd = false;
            updatedIncoming.CapStart = true;

            updatedOutgoing.StartOnLine = cp;
            updatedOutgoing.RailStartOffsetMm = 0.0;
            updatedOutgoing.CapStart = false;
            updatedOutgoing.CapEnd = true;

            corner = new CornerInfo
            {
                CornerOnLine = cp,
                IncomingDir = inDir,
                OutgoingDir = outDir,
                Left = incoming.Left,
                LateralOffsetMm = incoming.LateralOffsetMm,
                HalfPostWidthMm = incoming.HalfPostWidthMm,
                SideSign = sideSign,
                CornerTopZ = incoming.LastPostTopZ,
                CornerPost = null
            };

            return true;
        }

        private void CreateRails(
            Point startOnLine, Point endOnLine,
            Vector dirUnit, Vector leftUnit,
            double postLineLateralMm,
            double halfPostWidthMm,
            double firstPostTopZ, double lastPostTopZ,
            Part anyHostForSide,
            double railStartOffsetMm, double railEndOffsetMm,
            double railFromTopMm,
            int railCount, double railSpacingMm,
            bool capStart, bool capEnd)
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

                List<Beam> pieces = CreateRailPieces(a, b, maxLenMm, railProfile, railMaterial, railClass, railName);
                if (pieces.Count == 0) continue;

                if (capStart)
                    CreateRailEndCap(pieces[0], atStart: true);

                if (capEnd)
                    CreateRailEndCap(pieces[pieces.Count - 1], atStart: false);
            }
        }

        private static ContourPlate CreateRailEndCap(Beam rail, bool atStart)
        {
            try
            {
                if (rail == null) return null;
                Solid s = rail.GetSolid();
                if (s == null) return null;

                double minX = s.MinimumPoint.X;
                double minY = s.MinimumPoint.Y;
                double minZ = s.MinimumPoint.Z;
                double maxX = s.MaximumPoint.X;
                double maxY = s.MaximumPoint.Y;
                double maxZ = s.MaximumPoint.Z;

                Point end = atStart ? rail.StartPoint : rail.EndPoint;
                Point other = atStart ? rail.EndPoint : rail.StartPoint;
                Vector dir = new Vector(end.X - other.X, end.Y - other.Y, end.Z - other.Z);
                double len = Math.Sqrt(dir.X * dir.X + dir.Y * dir.Y + dir.Z * dir.Z);
                if (len < 1.0) return null;
                dir = new Vector(dir.X / len, dir.Y / len, dir.Z / len);

                double tMm = InchesToMm(0.125);
                double cx = end.X + dir.X * (tMm * 0.5);
                double cy = end.Y + dir.Y * (tMm * 0.5);
                double cz = end.Z + dir.Z * (tMm * 0.5);

                var cap = new ContourPlate();
                cap.Profile.ProfileString = "PL3.175";
                cap.Material.MaterialString = "A36";
                cap.Class = "4";
                cap.Name = "RAIL CAP";

                cap.Position.Plane = Position.PlaneEnum.MIDDLE;
                cap.Position.Rotation = Position.RotationEnum.TOP;
                cap.Position.Depth = Position.DepthEnum.MIDDLE;

                // Conservative rectangular cap oriented to global axes around end point.
                double hx = Math.Max(1.0, (maxX - minX) * 0.5);
                double hy = Math.Max(1.0, (maxY - minY) * 0.5);

                cap.Contour.AddContourPoint(new ContourPoint(new Point(cx - hx, cy - hy, cz), null));
                cap.Contour.AddContourPoint(new ContourPoint(new Point(cx + hx, cy - hy, cz), null));
                cap.Contour.AddContourPoint(new ContourPoint(new Point(cx + hx, cy + hy, cz), null));
                cap.Contour.AddContourPoint(new ContourPoint(new Point(cx - hx, cy + hy, cz), null));

                return cap.Insert() ? cap : null;
            }
            catch
            {
                return null;
            }
        }


        private static List<Beam> CreateRailPieces(
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
                    railSpacingMm,
                    capStart: true,
                    capEnd: true
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
