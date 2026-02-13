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


        // Per-side rail build data captured during post creation.
        private sealed class RailSideSpec
        {
            public Point StartOnLine;
            public Point EndOnLine;
            public Vector Dir;
            public Vector Left;
            public double PostLineLateralMm;
            public double HalfPostWidthMm;
            public double RailLateralMm;
            public double FirstPostTopZ;
            public double LastPostTopZ;
            public Part AnyHost;
        }

        private sealed class CornerSeatSpec
        {
            public Point CornerRailPoint;
            public Vector PrevDir;
            public Vector NextDir;
            public bool IsInsideCorner;
            public double RailCenterZ;
            public double HalfRailDepthMm;
        }

        private sealed class CornerFitSpec
        {
            public int SideIndex;
            public bool AtStart;
            public Point FacePoint;
            public Vector FaceNormal;
        }

        // Multi-side rail builder:
        //  - avoids corner clipping by trimming/extending the butt side start to the previous side face.
        //  - caps non-butt ends.
        //  - inserts corner seat angle (slots only, no pilot holes) per row/per corner.
        private void CreateCornerAwareRailsForPolyline(
            List<RailSideSpec> sides,
            bool isClosed,
            double railStartOffsetMm,
            double railEndOffsetMm,
            double railFromTopMm,
            int railCount,
            double railSpacingMm,
            double seatHoleLineIn,
            double seatSlotC2CIn,
            double seatSlotSizeIn,
            string seatSlotStandard,
            double seatSlotCutLengthIn,
            bool seatSlotSpecialFirstLayer)
        {
            if (sides == null || sides.Count == 0 || railCount <= 0) return;

            const string railProfile = "TS1-1/2X1-1/2X.188";
            const string railMaterial = "A53";
            const string railClass = "1";
            const string railName = "RAIL";
            double maxLenMm = InchesToMm(240.0); // 20'-0"

            double halfRailWidthMm = InchesToMm(1.5) * 0.5;
            if (TryGetOutsideDimMm(railProfile, out var railOutsideMm))
                halfRailWidthMm = railOutsideMm * 0.5;

            int n = sides.Count;
            if (n == 1)
            {
                var only = sides[0];
                CreateRails(
                    only.StartOnLine, only.EndOnLine,
                    only.Dir, only.Left,
                    only.PostLineLateralMm,
                    only.HalfPostWidthMm,
                    only.FirstPostTopZ, only.LastPostTopZ,
                    only.AnyHost,
                    railStartOffsetMm, railEndOffsetMm,
                    railFromTopMm,
                    railCount, railSpacingMm);
                return;
            }

            for (int r = 0; r < railCount; r++)
            {
                var starts = new Point[n];
                var ends = new Point[n];

                for (int i = 0; i < n; i++)
                {
                    RailSideSpec s = sides[i];

                    Point sLine = new Point(
                        s.StartOnLine.X - s.Dir.X * railStartOffsetMm,
                        s.StartOnLine.Y - s.Dir.Y * railStartOffsetMm,
                        s.StartOnLine.Z - s.Dir.Z * railStartOffsetMm);

                    Point eLine = new Point(
                        s.EndOnLine.X + s.Dir.X * railEndOffsetMm,
                        s.EndOnLine.Y + s.Dir.Y * railEndOffsetMm,
                        s.EndOnLine.Z + s.Dir.Z * railEndOffsetMm);

                    double zStart = s.FirstPostTopZ - railFromTopMm - (r * railSpacingMm);
                    double zEnd = s.LastPostTopZ - railFromTopMm - (r * railSpacingMm);

                    starts[i] = new Point(
                        sLine.X + s.Left.X * s.RailLateralMm,
                        sLine.Y + s.Left.Y * s.RailLateralMm,
                        zStart);

                    ends[i] = new Point(
                        eLine.X + s.Left.X * s.RailLateralMm,
                        eLine.Y + s.Left.Y * s.RailLateralMm,
                        zEnd);
                }

                var capStart = new bool[n];
                var capEnd = new bool[n];
                var cornerSeats = new List<CornerSeatSpec>();
                var fitSpecs = new List<CornerFitSpec>();

                // Keep unmodified row centerlines for stable corner-center math.
                var baseStarts = new Point[n];
                var baseEnds = new Point[n];
                for (int i = 0; i < n; i++)
                {
                    baseStarts[i] = new Point(starts[i].X, starts[i].Y, starts[i].Z);
                    baseEnds[i] = new Point(ends[i].X, ends[i].Y, ends[i].Z);
                }

                if (!isClosed)
                {
                    capStart[0] = true;          // open polyline start
                    capEnd[n - 1] = true;        // open polyline end
                }

                int cornerCount = isClosed ? n : (n - 1);
                for (int c = 0; c < cornerCount; c++)
                {
                    int prev = c;
                    int next = (c + 1) % n;

                    Vector prevDir = UnitVector(sides[prev].Dir);
                    Vector nextDir = UnitVector(sides[next].Dir);

                    // Option A: PREV is non-butt (cap at end), NEXT is butt side.
                    bool optPrevCapOk = ComputeButtStartToSideFace(
                        ends[prev],          // fixed side corner point
                        prevDir,             // fixed side direction
                        sides[prev].Left,    // fixed side left
                        starts[next],        // moving butt point
                        nextDir,             // moving butt direction
                        halfRailWidthMm,
                        out Point optNextStart,
                        out double moveNextButtMm,
                        out Point optPrevCapFacePoint,
                        out Vector optPrevCapFaceNormal);

                    // Option B: NEXT is non-butt (cap at start), PREV is butt side.
                    bool optNextCapOk = ComputeButtStartToSideFace(
                        starts[next],        // fixed side corner point
                        nextDir,             // fixed side direction
                        sides[next].Left,    // fixed side left
                        ends[prev],          // moving butt point
                        prevDir,             // moving butt direction
                        halfRailWidthMm,
                        out Point optPrevEnd,
                        out double movePrevButtMm,
                        out Point optNextCapFacePoint,
                        out Vector optNextCapFaceNormal);

                    if (optPrevCapOk)
                    {
                        if (Distance3D(optNextStart, ends[next]) < 1.0)
                            optPrevCapOk = false;
                    }

                    if (optNextCapOk)
                    {
                        if (Distance3D(starts[prev], optPrevEnd) < 1.0)
                            optNextCapOk = false;
                    }

                    bool choosePrevCap = false;
                    bool chooseNextCap = false;

                    if (optPrevCapOk && optNextCapOk)
                    {
                        choosePrevCap = Math.Abs(moveNextButtMm) <= Math.Abs(movePrevButtMm);
                        chooseNextCap = !choosePrevCap;
                    }
                    else if (optPrevCapOk)
                    {
                        choosePrevCap = true;
                    }
                    else if (optNextCapOk)
                    {
                        chooseNextCap = true;
                    }

                    if (choosePrevCap)
                    {
                        // PREV is non-butt + capped. NEXT butts into PREV side face.
                        starts[next] = optNextStart;
                        capEnd[prev] = true;

                        fitSpecs.Add(new CornerFitSpec
                        {
                            SideIndex = next,
                            AtStart = true,
                            FacePoint = optPrevCapFacePoint,
                            FaceNormal = optPrevCapFaceNormal
                        });
                    }
                    else if (chooseNextCap)
                    {
                        // NEXT is non-butt + capped. PREV butts into NEXT side face.
                        ends[prev] = optPrevEnd;
                        capStart[next] = true;

                        fitSpecs.Add(new CornerFitSpec
                        {
                            SideIndex = prev,
                            AtStart = false,
                            FacePoint = optNextCapFacePoint,
                            FaceNormal = optNextCapFaceNormal
                        });
                    }
                    else
                    {
                        // Fallback: cap previous side end.
                        capEnd[prev] = true;
                    }

                    // Inside/outside corner classification based on path turn + rail offset side.
                    double turn = CrossZ(prevDir, nextDir);
                    double lateral = 0.5 * (sides[prev].RailLateralMm + sides[next].RailLateralMm);
                    if (Math.Abs(lateral) < 1e-6)
                        lateral = sides[prev].RailLateralMm;
                    if (Math.Abs(lateral) < 1e-6)
                        lateral = 1.0; // stable fallback

                    bool isInside = (turn * lateral) < 0.0;

                    // Corner reference from centerline intersection.
                    Point cornerPt;
                    if (TryIntersectLines2D(baseEnds[prev], prevDir, baseStarts[next], nextDir, out Point xpt))
                    {
                        double z = 0.5 * (ends[prev].Z + starts[next].Z);
                        cornerPt = new Point(xpt.X, xpt.Y, z);
                    }
                    else
                    {
                        cornerPt = new Point(
                            0.5 * (ends[prev].X + starts[next].X),
                            0.5 * (ends[prev].Y + starts[next].Y),
                            0.5 * (ends[prev].Z + starts[next].Z));
                    }

                    cornerSeats.Add(new CornerSeatSpec
                    {
                        CornerRailPoint = cornerPt,
                        PrevDir = prevDir,
                        NextDir = nextDir,
                        IsInsideCorner = isInside,
                        RailCenterZ = cornerPt.Z,
                        HalfRailDepthMm = halfRailWidthMm
                    });
                }

                // Build rails + caps
                var firstPieceBySide = new Beam[n];
                var lastPieceBySide = new Beam[n];

                for (int i = 0; i < n; i++)
                {
                    if (Distance3D(starts[i], ends[i]) < 1.0) continue;

                    var pieces = CreateRailPiecesCollect(starts[i], ends[i], maxLenMm, railProfile, railMaterial, railClass, railName);
                    if (pieces.Count > 0)
                    {
                        firstPieceBySide[i] = pieces[0];
                        lastPieceBySide[i] = pieces[pieces.Count - 1];
                    }

                    Vector d = UnitVector(new Vector(
                        ends[i].X - starts[i].X,
                        ends[i].Y - starts[i].Y,
                        ends[i].Z - starts[i].Z));

                    if (capStart[i]) CreateRailEndCap(starts[i], new Vector(-d.X, -d.Y, -d.Z), halfRailWidthMm);
                    if (capEnd[i]) CreateRailEndCap(ends[i], d, halfRailWidthMm);
                }

                // Apply end fittings on butt sides so they truly terminate on side-face planes.
                foreach (var fit in fitSpecs)
                {
                    if (fit == null) continue;
                    if (fit.SideIndex < 0 || fit.SideIndex >= n) continue;

                    Beam target = fit.AtStart ? firstPieceBySide[fit.SideIndex] : lastPieceBySide[fit.SideIndex];
                    TryApplyEndFitting(target, fit.FacePoint, fit.FaceNormal);
                }

                // Corner seat angle per corner/row (slots only, no pilot holes)
                foreach (var cs in cornerSeats)
                {
                    CreateCornerSeatAngleSlotsOnly(
                        cs.CornerRailPoint,
                        cs.PrevDir,
                        cs.NextDir,
                        cs.IsInsideCorner,
                        cs.RailCenterZ,
                        cs.HalfRailDepthMm,
                        seatHoleLineIn,
                        seatSlotC2CIn,
                        seatSlotSizeIn,
                        seatSlotStandard,
                        seatSlotCutLengthIn,
                        seatSlotSpecialFirstLayer
                    );
                }
            }
        }


        // Finds where a moving rail point should be so it butts to the side face of a fixed rail.
        // Returns adjusted point + signed move distance along movingDir + chosen face plane data.
        private static bool ComputeButtStartToSideFace(
            Point fixedCornerPoint,
            Vector fixedDir,
            Vector fixedLeft,
            Point movingPoint,
            Vector movingDir,
            double halfRailWidthMm,
            out Point adjustedPoint,
            out double moveAlongMm,
            out Point chosenFacePoint,
            out Vector chosenFaceNormal)
        {
            adjustedPoint = movingPoint;
            moveAlongMm = 0.0;
            chosenFacePoint = fixedCornerPoint;
            chosenFaceNormal = new Vector(1.0, 0.0, 0.0);

            Vector dMove = UnitVector(movingDir);
            Vector left = UnitVector(fixedLeft);

            if (LengthXY(dMove) < 1e-9 || LengthXY(left) < 1e-9)
                return false;

            bool found = false;
            double bestScore = double.MaxValue;
            Point best = movingPoint;
            double bestT = 0.0;
            Point bestFacePoint = fixedCornerPoint;
            Vector bestFaceNormal = left;

            foreach (int sgn in new[] { +1, -1 })
            {
                Vector n = new Vector(left.X * sgn, left.Y * sgn, 0.0);
                n = UnitVector(n);

                Point q = new Point(
                    fixedCornerPoint.X + n.X * halfRailWidthMm,
                    fixedCornerPoint.Y + n.Y * halfRailWidthMm,
                    fixedCornerPoint.Z);

                double denom = Dot2D(n, dMove);
                if (Math.Abs(denom) < 1e-9) continue;

                double num = (q.X - movingPoint.X) * n.X + (q.Y - movingPoint.Y) * n.Y;
                double t = num / denom;

                Point cand = new Point(
                    movingPoint.X + dMove.X * t,
                    movingPoint.Y + dMove.Y * t,
                    movingPoint.Z + dMove.Z * t);

                // Score = shortest valid move, with tie-break favoring closer to fixed corner.
                double cornerDist = Distance3D(cand, fixedCornerPoint);
                double score = Math.Abs(t) + (cornerDist * 1e-4);

                if (score < bestScore)
                {
                    bestScore = score;
                    best = cand;
                    bestT = t;
                    bestFacePoint = q;
                    bestFaceNormal = n;
                    found = true;
                }
            }

            if (!found)
                return false;

            adjustedPoint = best;
            moveAlongMm = bestT;
            chosenFacePoint = bestFacePoint;
            chosenFaceNormal = bestFaceNormal;
            return true;
        }

        private static bool TryIntersectLines2D(Point p, Vector dp, Point q, Vector dq, out Point intersection)
        {
            intersection = new Point(0.0, 0.0, 0.0);

            Vector a = UnitVector(new Vector(dp.X, dp.Y, 0.0));
            Vector b = UnitVector(new Vector(dq.X, dq.Y, 0.0));

            double den = CrossZ(a, b);
            if (Math.Abs(den) < 1e-9)
                return false;

            Vector qp = new Vector(q.X - p.X, q.Y - p.Y, 0.0);
            double t = CrossZ(qp, b) / den;

            intersection = new Point(
                p.X + a.X * t,
                p.Y + a.Y * t,
                0.0);

            return true;
        }

        private static ContourPlate CreateRailEndCap(Point endCenter, Vector endDir, double halfRailWidthMm)
        {
            try
            {
                if (endCenter == null) return null;
                Vector d = UnitVector(endDir);
                if (LengthXY(d) < 1e-9) d = new Vector(1.0, 0.0, 0.0);

                Vector side = GetLeftVectorXY(d);
                Vector up = new Vector(0.0, 0.0, 1.0);

                // 1/8" cap plate
                double tMm = InchesToMm(0.125);
                Point c = new Point(
                    endCenter.X + d.X * (tMm * 0.5),
                    endCenter.Y + d.Y * (tMm * 0.5),
                    endCenter.Z + d.Z * (tMm * 0.5));

                Point p1 = new Point(c.X + side.X * halfRailWidthMm + up.X * halfRailWidthMm, c.Y + side.Y * halfRailWidthMm + up.Y * halfRailWidthMm, c.Z + side.Z * halfRailWidthMm + up.Z * halfRailWidthMm);
                Point p2 = new Point(c.X - side.X * halfRailWidthMm + up.X * halfRailWidthMm, c.Y - side.Y * halfRailWidthMm + up.Y * halfRailWidthMm, c.Z - side.Z * halfRailWidthMm + up.Z * halfRailWidthMm);
                Point p3 = new Point(c.X - side.X * halfRailWidthMm - up.X * halfRailWidthMm, c.Y - side.Y * halfRailWidthMm - up.Y * halfRailWidthMm, c.Z - side.Z * halfRailWidthMm - up.Z * halfRailWidthMm);
                Point p4 = new Point(c.X + side.X * halfRailWidthMm - up.X * halfRailWidthMm, c.Y + side.Y * halfRailWidthMm - up.Y * halfRailWidthMm, c.Z + side.Z * halfRailWidthMm - up.Z * halfRailWidthMm);

                var cap = new ContourPlate();
                cap.Profile.ProfileString = "PL3.175";
                cap.Material.MaterialString = "A36";
                cap.Class = "4";
                cap.Name = "RAIL CAP";

                cap.Position.Plane = Position.PlaneEnum.MIDDLE;
                cap.Position.Rotation = Position.RotationEnum.TOP;
                cap.Position.Depth = Position.DepthEnum.MIDDLE;

                cap.Contour.AddContourPoint(new ContourPoint(p1, null));
                cap.Contour.AddContourPoint(new ContourPoint(p2, null));
                cap.Contour.AddContourPoint(new ContourPoint(p3, null));
                cap.Contour.AddContourPoint(new ContourPoint(p4, null));

                if (cap.Insert()) return cap;
            }
            catch
            {
                // cap is best-effort
            }

            return null;
        }


        private static List<Beam> CreateRailPiecesCollect(
            Point a, Point b, double maxLenMm,
            string profile, string material, string cls, string name)
        {
            var outPieces = new List<Beam>();

            Vector v = new Vector(b.X - a.X, b.Y - a.Y, b.Z - a.Z);
            double total = Math.Sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z);
            if (total < 1.0) return outPieces;

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
                    outPieces.Add(rail);
            }

            return outPieces;
        }

        private static bool TryApplyEndFitting(Beam target, Point facePoint, Vector faceNormal)
        {
            try
            {
                if (target == null || facePoint == null || faceNormal == null) return false;

                Vector n = UnitVector(faceNormal);
                double nLen = Math.Sqrt(n.X * n.X + n.Y * n.Y + n.Z * n.Z);
                if (nLen < 1e-9) return false;

                Vector axisX = new Vector(-n.Y, n.X, 0.0);
                if (LengthXY(axisX) < 1e-9)
                    axisX = new Vector(1.0, 0.0, 0.0);
                axisX = UnitVector(axisX);

                Vector axisY = new Vector(
                    n.Y * axisX.Z - n.Z * axisX.Y,
                    n.Z * axisX.X - n.X * axisX.Z,
                    n.X * axisX.Y - n.Y * axisX.X);
                axisY = UnitVector(axisY);

                var pl = new Plane
                {
                    Origin = new Point(facePoint.X, facePoint.Y, facePoint.Z),
                    AxisX = axisX,
                    AxisY = axisY
                };

                var fit = new Fitting
                {
                    Father = target,
                    Plane = pl
                };

                return fit.Insert();
            }
            catch
            {
                return false;
            }
        }

        private static double Distance3D(Point a, Point b)
        {
            if (a == null || b == null) return 0.0;
            double dx = a.X - b.X;
            double dy = a.Y - b.Y;
            double dz = a.Z - b.Z;
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }

        private static Vector UnitVector(Vector v)
        {
            if (v == null) return new Vector(1.0, 0.0, 0.0);
            double len = Math.Sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z);
            if (len < 1e-12) return new Vector(1.0, 0.0, 0.0);
            return new Vector(v.X / len, v.Y / len, v.Z / len);
        }

        private static double LengthXY(Vector v)
        {
            if (v == null) return 0.0;
            return Math.Sqrt(v.X * v.X + v.Y * v.Y);
        }

        private static double Dot2D(Vector a, Vector b)
        {
            return a.X * b.X + a.Y * b.Y;
        }

        private static double CrossZ(Vector a, Vector b)
        {
            return a.X * b.Y - a.Y * b.X;
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
