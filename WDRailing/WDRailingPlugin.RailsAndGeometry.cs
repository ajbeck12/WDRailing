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
            public string CornerClass;
        }

        private sealed class CornerFitSpec
        {
            public int SideIndex;
            public bool AtStart;
            public Point FacePoint;
            public Vector FaceNormal;
        }
        private void CreateSquareRailPiecesWithSplices(
    Point start,
    Point end,
    double maxLenMm)
        {
            if (start == null || end == null)
                return;

            Vector v = new Vector(end.X - start.X, end.Y - start.Y, end.Z - start.Z);
            double len = Math.Sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z);
            if (len < 1.0)
                return;

            if (maxLenMm <= 1.0 || len <= maxLenMm)
            {
                CreateSingleSquareRailPiece(start, end);
                return;
            }

            Vector dir = new Vector(v.X / len, v.Y / len, v.Z / len);
            int pieceCount = Math.Max(1, (int)Math.Ceiling(len / maxLenMm));
            double pieceLen = len / pieceCount;

            Point pieceStart = start;

            for (int i = 1; i <= pieceCount; i++)
            {
                bool isLast = (i == pieceCount);

                Point pieceEnd;
                if (isLast)
                {
                    pieceEnd = end;
                }
                else
                {
                    pieceEnd = new Point(
                        start.X + dir.X * (pieceLen * i),
                        start.Y + dir.Y * (pieceLen * i),
                        start.Z + dir.Z * (pieceLen * i));
                }

                CreateSingleSquareRailPiece(pieceStart, pieceEnd);

                if (!isLast)
                    CreateInnerRailSplice(pieceEnd, dir, InchesToMm(4.0));

                pieceStart = pieceEnd;
            }
        }

        private Beam CreateSingleSquareRailPiece(Point start, Point end)
        {
            Beam rail = new Beam(start, end);

            rail.Profile.ProfileString = "TS1-1/2X1-1/2X.188";
            rail.Material.MaterialString = "A50";
            rail.Class = "1";
            rail.Name = "CL11 HORIZ";

            rail.Position.Plane = Position.PlaneEnum.MIDDLE;
            rail.Position.Depth = Position.DepthEnum.MIDDLE;
            rail.Position.Rotation = Position.RotationEnum.TOP;

            if (!rail.Insert())
                return null;

            rail.Modify();
            return rail;
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
    bool seatSlotSpecialFirstLayer,
    bool startLoopEnabled,
    bool[] startLoopMask,
    bool endLoopEnabled,
    bool[] endLoopMask,
    double railMaxLenMm)
        {
            if (sides == null || sides.Count == 0 || railCount <= 0) return;

            const string railProfile = "TS1-1/2X1-1/2X.188";
            const string railMaterial = "A53";
            const string railClass = "1";
            const string railName = "RAIL";
            double maxLenMm = railMaxLenMm;
            if (maxLenMm <= 1.0)
                maxLenMm = InchesToMm(240.0); // fallback

            double halfRailWidthMm = InchesToMm(1.5) * 0.5;
            if (TryGetOutsideDimMm(railProfile, out var railOutsideMm))
                halfRailWidthMm = railOutsideMm * 0.5;

            double railOutsideDiaMm = 2.0 * halfRailWidthMm;

            int n = sides.Count;

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

                var baseStarts = new Point[n];
                var baseEnds = new Point[n];
                for (int i = 0; i < n; i++)
                {
                    baseStarts[i] = new Point(starts[i].X, starts[i].Y, starts[i].Z);
                    baseEnds[i] = new Point(ends[i].X, ends[i].Y, ends[i].Z);
                }

                if (!isClosed)
                {
                    capStart[0] = true;
                    capEnd[n - 1] = true;
                }

                int cornerCount = isClosed ? n : (n - 1);
                for (int c = 0; c < cornerCount; c++)
                {
                    int prev = c;
                    int next = (c + 1) % n;

                    Vector prevDir = UnitVector(sides[prev].Dir);
                    Vector nextDir = UnitVector(sides[next].Dir);

                    Point cornerPt;
                    if (TryIntersectLines2D(baseEnds[prev], prevDir, baseStarts[next], nextDir, out Point xpt))
                    {
                        double z = 0.5 * (ends[prev].Z + starts[next].Z);
                        cornerPt = new Point(xpt.X, xpt.Y, z);
                    }
                    else
                    {
                        cornerPt = new Point(
                            0.5 * (baseEnds[prev].X + baseStarts[next].X),
                            0.5 * (baseEnds[prev].Y + baseStarts[next].Y),
                            0.5 * (ends[prev].Z + starts[next].Z));
                    }

                    double turnForOffset = CrossZ(prevDir, nextDir);
                    double lateralForOffset = 0.5 * (sides[prev].RailLateralMm + sides[next].RailLateralMm);
                    if (Math.Abs(lateralForOffset) < 1e-6) lateralForOffset = sides[prev].RailLateralMm;
                    if (Math.Abs(lateralForOffset) < 1e-6) lateralForOffset = 1.0;
                    bool isInsideForOffset = (turnForOffset * lateralForOffset) < 0.0;

                    double cornerButtOffsetMm = halfRailWidthMm;

                    bool optPrevCapOk = ComputeButtStartToSideFace(
                        ends[prev], prevDir, sides[prev].Left,
                        starts[next], nextDir,
                        isInsideForOffset,
                        cornerButtOffsetMm,
                        out Point optNextStart, out double moveNextButtMm,
                        out Point optPrevCapFacePoint, out Vector optPrevCapFaceNormal);

                    bool optNextCapOk = ComputeButtStartToSideFace(
                        starts[next], nextDir, sides[next].Left,
                        ends[prev], prevDir,
                        isInsideForOffset,
                        cornerButtOffsetMm,
                        out Point optPrevEnd, out double movePrevButtMm,
                        out Point optNextCapFacePoint, out Vector optNextCapFaceNormal);

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
                        starts[next] = optNextStart;
                        capEnd[prev] = true;

                        ends[prev] = new Point(
                            cornerPt.X + prevDir.X * cornerButtOffsetMm,
                            cornerPt.Y + prevDir.Y * cornerButtOffsetMm,
                            ends[prev].Z);

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
                        ends[prev] = optPrevEnd;
                        capStart[next] = true;

                        starts[next] = new Point(
                            cornerPt.X - nextDir.X * cornerButtOffsetMm,
                            cornerPt.Y - nextDir.Y * cornerButtOffsetMm,
                            starts[next].Z);

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
                        capEnd[prev] = true;
                    }

                    double turn = CrossZ(prevDir, nextDir);
                    double lateral = 0.5 * (sides[prev].RailLateralMm + sides[next].RailLateralMm);
                    if (Math.Abs(lateral) < 1e-6) lateral = sides[prev].RailLateralMm;
                    if (Math.Abs(lateral) < 1e-6) lateral = 1.0;
                    bool isInside = (turn * lateral) < 0.0;
                    string cornerClass = GetCornerDebugClassByRunDirection(prevDir, nextDir, isInside);

                    cornerSeats.Add(new CornerSeatSpec
                    {
                        CornerRailPoint = cornerPt,
                        PrevDir = prevDir,
                        NextDir = nextDir,
                        IsInsideCorner = isInside,
                        RailCenterZ = cornerPt.Z,
                        HalfRailDepthMm = halfRailWidthMm,
                        CornerClass = cornerClass
                    });
                }

                // OPEN-END LOOP RULES
                if (!isClosed)
                {
                    // START
                    if (startLoopEnabled && IsSelectedLoopRow(startLoopMask, railCount, r))
                    {
                        bool isEndRow = IsTerminalLoopEndRow(startLoopMask, railCount, r);
                        bool isMidRow = IsMiddleLoopRow(startLoopMask, railCount, r);

                        // Top/bottom loop rows KEEP their caps.
                        // Middle loop rows trim back to the inside face and get NO cap.
                        capStart[0] = isEndRow;

                        if (isMidRow)
                        {
                            Vector d0 = UnitVector(sides[0].Dir);
                            starts[0] = new Point(
                                starts[0].X + d0.X * railOutsideDiaMm,
                                starts[0].Y + d0.Y * railOutsideDiaMm,
                                starts[0].Z);
                        }
                    }

                    // END
                    if (endLoopEnabled && IsSelectedLoopRow(endLoopMask, railCount, r))
                    {
                        bool isEndRow = IsTerminalLoopEndRow(endLoopMask, railCount, r);
                        bool isMidRow = IsMiddleLoopRow(endLoopMask, railCount, r);

                        capEnd[n - 1] = isEndRow;

                        if (isMidRow)
                        {
                            Vector dn = UnitVector(sides[n - 1].Dir);
                            ends[n - 1] = new Point(
                                ends[n - 1].X - dn.X * railOutsideDiaMm,
                                ends[n - 1].Y - dn.Y * railOutsideDiaMm,
                                ends[n - 1].Z);
                        }
                    }
                }

                var firstPieceBySide = new Beam[n];
                var lastPieceBySide = new Beam[n];

                for (int i = 0; i < n; i++)
                {
                    if (Distance3D(starts[i], ends[i]) < 1.0) continue;

                    var pieces = CreateRailPiecesCollect(
                        starts[i],
                        ends[i],
                        maxLenMm,
                        railProfile,
                        railMaterial,
                        railClass,
                        railName,
                        true);
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

                foreach (var fit in fitSpecs)
                {
                    if (fit == null) continue;
                    if (fit.SideIndex < 0 || fit.SideIndex >= n) continue;

                    Beam moving = fit.AtStart ? firstPieceBySide[fit.SideIndex] : lastPieceBySide[fit.SideIndex];
                    if (moving == null) continue;

                    TryApplyEndFitting(moving, fit.FacePoint, fit.FaceNormal);
                    TryForceBeamEndToPlane(moving, fit.AtStart, fit.FacePoint, fit.FaceNormal);
                }

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
                        seatSlotSpecialFirstLayer,
                        cs.CornerClass
                    );
                }

                // END LOOP SEAT-ANGLE HOOKS
                if (!isClosed)
                {
                    // Start terminal seat angle for CURRENT row only
                    if (startLoopEnabled && IsSelectedLoopRow(startLoopMask, railCount, r))
                    {
                        Point startInsidePoint = GetTerminalLoopInsidePoint(
                            sides[0],
                            atStart: true,
                            railTerminalOffsetMm: railStartOffsetMm,
                            railOutsideDiaMm: railOutsideDiaMm);

                        double railZ = sides[0].FirstPostTopZ - railFromTopMm - (r * railSpacingMm);

                        CreateTerminalLoopSeatAngleHook(
                            startInsidePoint,
                            UnitVector(sides[0].Dir),
                            railZ,
                            halfRailWidthMm,
                            seatHoleLineIn,
                            seatSlotC2CIn,
                            seatSlotSizeIn,
                            seatSlotStandard,
                            seatSlotCutLengthIn,
                            seatSlotSpecialFirstLayer,
                            atStart: true,
                            isBottomLoopRow: IsBottomLoopEndRow(startLoopMask, railCount, r));
                    }

                    // End terminal seat angle for CURRENT row only
                    if (endLoopEnabled && IsSelectedLoopRow(endLoopMask, railCount, r))
                    {
                        Point endInsidePoint = GetTerminalLoopInsidePoint(
                            sides[n - 1],
                            atStart: false,
                            railTerminalOffsetMm: railEndOffsetMm,
                            railOutsideDiaMm: railOutsideDiaMm);

                        double railZ = sides[n - 1].LastPostTopZ - railFromTopMm - (r * railSpacingMm);

                        CreateTerminalLoopSeatAngleHook(
                            endInsidePoint,
                            UnitVector(sides[n - 1].Dir),
                            railZ,
                            halfRailWidthMm,
                            seatHoleLineIn,
                            seatSlotC2CIn,
                            seatSlotSizeIn,
                            seatSlotStandard,
                            seatSlotCutLengthIn,
                            seatSlotSpecialFirstLayer,
                            atStart: false,
                            isBottomLoopRow: IsBottomLoopEndRow(endLoopMask, railCount, r));
                    }
                }
            }
        }

        private static bool IsSelectedLoopRow(bool[] rowMask, int railCount, int rowIndex)
        {
            if (rowMask == null) return false;
            if (rowIndex < 0 || rowIndex >= railCount) return false;
            if (rowIndex >= rowMask.Length) return false;
            return rowMask[rowIndex];
        }

        private static bool TryGetSelectedRowRange(bool[] rowMask, int railCount, out int firstRow, out int lastRow)
        {
            firstRow = -1;
            lastRow = -1;

            if (rowMask == null || railCount <= 0)
                return false;

            int max = Math.Min(rowMask.Length, railCount);
            for (int i = 0; i < max; i++)
            {
                if (!rowMask[i]) continue;

                if (firstRow < 0)
                    firstRow = i;

                lastRow = i;
            }

            return firstRow >= 0 && lastRow >= 0;
        }

        private static bool IsMiddleLoopRow(bool[] rowMask, int railCount, int rowIndex)
        {
            if (!TryGetSelectedRowRange(rowMask, railCount, out int firstRow, out int lastRow))
                return false;

            if (rowIndex < 0 || rowIndex >= railCount) return false;
            if (rowIndex >= rowMask.Length) return false;

            return rowMask[rowIndex] && rowIndex > firstRow && rowIndex < lastRow;
        }

        private static bool IsTerminalLoopEndRow(bool[] rowMask, int railCount, int rowIndex)
        {
            if (!TryGetSelectedRowRange(rowMask, railCount, out int firstRow, out int lastRow))
                return false;

            return rowIndex == firstRow || rowIndex == lastRow;
        }

        private Point GetTerminalLoopInsidePoint(
            RailSideSpec side,
            bool atStart,
            double railTerminalOffsetMm,
            double railOutsideDiaMm)
        {
            Vector dir = UnitVector(side.Dir);
            Point onLine = atStart ? side.StartOnLine : side.EndOnLine;

            // Terminal plane of horizontal rail outside face:
            // start = StartOnLine - Dir * railStartOffset
            // end   = EndOnLine   + Dir * railEndOffset
            //
            // Inside face is one rail OD inward from that plane.
            double along = atStart
                ? (-railTerminalOffsetMm + railOutsideDiaMm)
                : (railTerminalOffsetMm - railOutsideDiaMm);

            return new Point(
                onLine.X + side.Left.X * side.RailLateralMm + dir.X * along,
                onLine.Y + side.Left.Y * side.RailLateralMm + dir.Y * along,
                0.0);
        }

        private void CreateTerminalLoopSeatAngleHook(
    Point insidePointXY,
    Vector railDir,
    double railZ,
    double halfRailDepthMm,
    double seatHoleLineIn,
    double seatSlotC2CIn,
    double seatSlotSizeIn,
    string seatSlotStandard,
    double seatSlotCutLengthIn,
    bool seatSlotSpecialFirstLayer,
    bool atStart,
    bool isBottomLoopRow)
        {
            try
            {
                Vector dirXY = GetDirXYUnit(railDir);

                // Beam axis runs across the rail end face.
                Vector faceAxis = GetLeftVectorXY(dirXY);

                // Start side needs mirrored beam direction so the seat faces correctly.
                if (atStart)
                    faceAxis = new Vector(-faceAxis.X, -faceAxis.Y, -faceAxis.Z);

                double L = GetSeatAngleLengthMm();

                // Normal loop seats sit under the rail.
                // Bottom loop seat must sit on TOP of the bottom rail so it ties the rail + loop.
                double zSeat = isBottomLoopRow
                    ? (railZ + halfRailDepthMm + SEAT_Z_ADJUST_MM)
                    : (railZ - halfRailDepthMm + SEAT_Z_ADJUST_MM);

                Point a = new Point(
                    insidePointXY.X - faceAxis.X * (L * 0.5),
                    insidePointXY.Y - faceAxis.Y * (L * 0.5),
                    zSeat);

                Point b = new Point(
                    insidePointXY.X + faceAxis.X * (L * 0.5),
                    insidePointXY.Y + faceAxis.Y * (L * 0.5),
                    zSeat);

                var seat = new Beam(a, b);
                seat.Name = SEAT_ANGLE_NAME;
                seat.Profile.ProfileString = SEAT_ANGLE_PROFILE;
                seat.Material.MaterialString = SEAT_ANGLE_MATERIAL;
                seat.Class = SEAT_ANGLE_CLASS;

                seat.Position.Plane = Position.PlaneEnum.LEFT;

                if (isBottomLoopRow)
                {
                    seat.Position.Depth = Position.DepthEnum.FRONT;
                    seat.Position.Rotation = Position.RotationEnum.FRONT;
                }
                else
                {
                    seat.Position.Depth = Position.DepthEnum.BEHIND;
                    seat.Position.Rotation = Position.RotationEnum.BELOW;
                }

                if (!seat.Insert())
                    return;

                TryAddTerminalLoopSlotsOnly(
                    seat,
                    faceAxis,
                    seatHoleLineIn,
                    seatSlotC2CIn,
                    seatSlotSizeIn,
                    seatSlotStandard,
                    seatSlotCutLengthIn,
                    seatSlotSpecialFirstLayer,
                    isBottomLoopRow);
            }
            catch
            {
                // best effort only
            }
        }

        private static void TryAddTerminalLoopSlotsOnly(
    Beam seat,
    Vector slotAxisDir,
    double holeLineFromBendIn,
    double slotC2CIn,
    double slotSizeIn,
    string slotStandard,
    double slotCutLengthIn,
    bool slotSpecialFirstLayer,
    bool isBottomLoopRow)
        {
            if (seat == null) return;

            try
            {
                var cs = seat.GetCoordinateSystem();

                Vector axisX = NormalizeVectorOrFallback(cs.AxisX, new Vector(1.0, 0.0, 0.0));
                Vector axisY = NormalizeVectorOrFallback(cs.AxisY, GetDirXYUnit(slotAxisDir));

                Vector axisZ = new Vector(
                    axisX.Y * axisY.Z - axisX.Z * axisY.Y,
                    axisX.Z * axisY.X - axisX.X * axisY.Z,
                    axisX.X * axisY.Y - axisX.Y * axisY.X);
                axisZ = NormalizeVectorOrFallback(axisZ, new Vector(0.0, 0.0, 1.0));

                Point mid = new Point(
                    (seat.StartPoint.X + seat.EndPoint.X) * 0.5,
                    (seat.StartPoint.Y + seat.EndPoint.Y) * 0.5,
                    (seat.StartPoint.Z + seat.EndPoint.Z) * 0.5);

                double bendOffsetMm = InchesToMm(holeLineFromBendIn);
                double legThicknessHalfMm = InchesToMm(0.125 * 0.5);
                double orientLenMm = Math.Max(5.0, InchesToMm(0.25));

                Point slot1 = new Point(
                    mid.X + axisY.X * bendOffsetMm + axisZ.X * legThicknessHalfMm,
                    mid.Y + axisY.Y * bendOffsetMm + axisZ.Y * legThicknessHalfMm,
                    mid.Z + axisY.Z * bendOffsetMm + axisZ.Z * legThicknessHalfMm);

                Point slot2 = new Point(
                    mid.X + axisZ.X * bendOffsetMm + axisY.X * legThicknessHalfMm,
                    mid.Y + axisZ.Y * bendOffsetMm + axisY.Y * legThicknessHalfMm,
                    mid.Z + axisZ.Z * bendOffsetMm + axisY.Z * legThicknessHalfMm);

                Position.RotationEnum slot1Rotation = isBottomLoopRow
                    ? Position.RotationEnum.FRONT
                    : Position.RotationEnum.BELOW;

                Position.RotationEnum slot2Rotation = isBottomLoopRow
                    ? Position.RotationEnum.TOP
                    : Position.RotationEnum.FRONT;

                TryInsertSeatSlotHole(
                    seat,
                    slot1,
                    axisX,
                    orientLenMm,
                    slotStandard,
                    slotCutLengthIn,
                    slotSizeIn,
                    slotC2CIn,
                    slotSpecialFirstLayer,
                    slot1Rotation,
                    Position.DepthEnum.MIDDLE);

                TryInsertSeatSlotHole(
                    seat,
                    slot2,
                    axisX,
                    orientLenMm,
                    slotStandard,
                    slotCutLengthIn,
                    slotSizeIn,
                    slotC2CIn,
                    slotSpecialFirstLayer,
                    slot2Rotation,
                    Position.DepthEnum.MIDDLE);
            }
            catch
            {
                // best effort only
            }
        }

        private static bool IsBottomLoopEndRow(bool[] rowMask, int railCount, int rowIndex)
        {
            if (!TryGetSelectedRowRange(rowMask, railCount, out int firstRow, out int lastRow))
                return false;

            return rowIndex == lastRow;
        }

        // Finds where a moving rail point should be so it butts to the side face of a fixed rail.
        // Returns adjusted point + signed move distance along movingDir + chosen face plane data.
        private static bool ComputeButtStartToSideFace(
            Point fixedCornerPoint,
            Vector fixedDir,
            Vector fixedLeft,
            Point movingPoint,
            Vector movingDir,
            bool isInsideCorner,
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
            Vector dFix = UnitVector(new Vector(fixedDir.X, fixedDir.Y, 0.0));
            Vector left = UnitVector(new Vector(fixedLeft.X, fixedLeft.Y, 0.0));

            if (LengthXY(dMove) < 1e-9 || LengthXY(left) < 1e-9 || LengthXY(dFix) < 1e-9)
                return false;

            // Determine on which side of fixed centerline the moving rail currently sits.
            Vector fromCorner = new Vector(
                movingPoint.X - fixedCornerPoint.X,
                movingPoint.Y - fixedCornerPoint.Y,
                0.0);

            int sideSign = 0;
            double side = Dot2D(left, fromCorner);
            if (Math.Abs(side) > 1e-6)
                sideSign = (side >= 0.0) ? +1 : -1;

            if (sideSign == 0)
            {
                // Fallback when the moving corner lies almost exactly on fixed centerline.
                double turn = CrossZ(dFix, new Vector(dMove.X, dMove.Y, 0.0));
                sideSign = (turn >= 0.0) ? +1 : -1;
            }

            // INSIDE corners use the side the moving rail is on.
            // OUTSIDE corners must use the opposite face (fixes one-rail-width overshoot on outside corners).
            int faceSign = isInsideCorner ? sideSign : -sideSign;
            if (faceSign == 0) faceSign = +1;

            double offsetMm = Math.Abs(halfRailWidthMm);
            Vector n = UnitVector(new Vector(left.X * faceSign, left.Y * faceSign, 0.0));

            // If chosen face is nearly parallel to moving rail, flip to opposite face.
            double denom = Dot2D(n, new Vector(dMove.X, dMove.Y, 0.0));
            if (Math.Abs(denom) < 1e-9)
            {
                n = new Vector(-n.X, -n.Y, 0.0);
                denom = Dot2D(n, new Vector(dMove.X, dMove.Y, 0.0));
                if (Math.Abs(denom) < 1e-9)
                    return false;
            }

            Point q = new Point(
                fixedCornerPoint.X + n.X * offsetMm,
                fixedCornerPoint.Y + n.Y * offsetMm,
                fixedCornerPoint.Z);

            double num = (q.X - movingPoint.X) * n.X + (q.Y - movingPoint.Y) * n.Y;
            double t = num / denom;

            adjustedPoint = new Point(
                movingPoint.X + dMove.X * t,
                movingPoint.Y + dMove.Y * t,
                movingPoint.Z + dMove.Z * t);

            moveAlongMm = t;
            chosenFacePoint = new Point(adjustedPoint.X, adjustedPoint.Y, adjustedPoint.Z);

            // Stable normal orientation for fallback use.
            Vector nFit = n;
            Vector toMoving = new Vector(adjustedPoint.X - q.X, adjustedPoint.Y - q.Y, 0.0);
            if (Dot2D(nFit, toMoving) < 0.0)
                nFit = new Vector(-nFit.X, -nFit.Y, -nFit.Z);

            chosenFaceNormal = nFit;
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

        private Beam CreateRoundCornerCl11(Point start, Point end)
        {
            if (start == null || end == null)
                return null;

            Beam rail = new Beam(start, end);

            rail.Profile.ProfileString = "TS1-1/2X1-1/2X.188";
            rail.Material.MaterialString = "A50";
            rail.Class = "1";
            rail.Name = "ROUND CORNER CL11";

            rail.Position.Plane = Position.PlaneEnum.MIDDLE;
            rail.Position.Depth = Position.DepthEnum.MIDDLE;

            // Square tube, so rotation usually won't matter much geometrically,
            // but FRONT is a reasonable "turned on side" choice.
            rail.Position.Rotation = Position.RotationEnum.FRONT;

            if (!rail.Insert())
                return null;

            rail.Modify();
            return rail;
        }

        private void CreateRoundEndLoopsForPolyline(
    List<RailSideSpec> sides,
    RoundTerminalRailRefs terminalRefs,
    double railStartOffsetMm,
    double railEndOffsetMm,
    double railFromTopMm,
    int railCount,
    double railSpacingMm,
    bool startLoopEnabled,
    bool endLoopEnabled)
        {
            if (sides == null || sides.Count == 0 || railCount < 2)
                return;

            if (startLoopEnabled)
                CreateSingleRoundTerminalEndLoop(
                    sides[0],
                    terminalRefs != null ? terminalRefs.StartRailByRow : null,
                    atStart: true,
                    railStartOffsetMm: railStartOffsetMm,
                    railEndOffsetMm: railEndOffsetMm,
                    railFromTopMm: railFromTopMm,
                    railCount: railCount,
                    railSpacingMm: railSpacingMm);

            if (endLoopEnabled)
                CreateSingleRoundTerminalEndLoop(
                    sides[sides.Count - 1],
                    terminalRefs != null ? terminalRefs.EndRailByRow : null,
                    atStart: false,
                    railStartOffsetMm: railStartOffsetMm,
                    railEndOffsetMm: railEndOffsetMm,
                    railFromTopMm: railFromTopMm,
                    railCount: railCount,
                    railSpacingMm: railSpacingMm);
        }

        private Beam CreateRailSplice(
    Point center,
    Vector axisDir,
    double totalLenMm)
        {
            if (center == null || axisDir == null || totalLenMm <= 0.0)
                return null;

            Vector u;
            if (Math.Abs(axisDir.Z) > 0.999)
                u = new Vector(0.0, 0.0, Math.Sign(axisDir.Z) == 0 ? 1.0 : Math.Sign(axisDir.Z));
            else
                u = GetDirXYUnit(axisDir);

            double halfLen = totalLenMm * 0.5;

            Point start = new Point(
                center.X - u.X * halfLen,
                center.Y - u.Y * halfLen,
                center.Z - u.Z * halfLen);

            Point end = new Point(
                center.X + u.X * halfLen,
                center.Y + u.Y * halfLen,
                center.Z + u.Z * halfLen);

            Beam splice = new Beam(start, end);

            splice.Profile.ProfileString = "PD47.625*3.175";
            splice.Material.MaterialString = "A36";
            splice.Class = "1";
            splice.Name = "RAIL SPLICE";

            splice.Position.Plane = Position.PlaneEnum.MIDDLE;
            splice.Position.Depth = Position.DepthEnum.MIDDLE;
            splice.Position.Rotation = Position.RotationEnum.TOP;

            if (!splice.Insert())
                return null;

            splice.Modify();
            return splice;
        }

        private void CreateSingleRoundTerminalEndLoop(
    RailSideSpec side,
    Beam[] straightRailsByRow,
    bool atStart,
    double railStartOffsetMm,
    double railEndOffsetMm,
    double railFromTopMm,
    int railCount,
    double railSpacingMm)
        {
            if (side == null || railCount < 2)
                return;

            double cl11LegMm = InchesToMm(10.25);
            double spliceLenMm = InchesToMm(4.0);
            double teeVerticalSpliceLenMm = InchesToMm(8.0);

            // RAIL SPLICE OD = 1-7/8
            double spliceOdMm = InchesToMm(1.875);

            Vector railDir = GetDirXYUnit(side.Dir);
            if (railDir == null)
                return;

            Vector inwardDir = atStart
                ? new Vector(railDir.X, railDir.Y, 0.0)
                : new Vector(-railDir.X, -railDir.Y, 0.0);

            Vector upDir = new Vector(0.0, 0.0, 1.0);

            double along = atStart ? -railStartOffsetMm : railEndOffsetMm;
            Point onLine = atStart ? side.StartOnLine : side.EndOnLine;
            double lateral = side.PostLineLateralMm;

            Point topCorner = GetRoundTerminalRowPoint(
                side, onLine, along, lateral, railFromTopMm, railSpacingMm, 0);

            Point botCorner = GetRoundTerminalRowPoint(
                side, onLine, along, lateral, railFromTopMm, railSpacingMm, railCount - 1);

            Point topRailJoint = new Point(
                topCorner.X + inwardDir.X * cl11LegMm,
                topCorner.Y + inwardDir.Y * cl11LegMm,
                topCorner.Z);

            Point botRailJoint = new Point(
                botCorner.X + inwardDir.X * cl11LegMm,
                botCorner.Y + inwardDir.Y * cl11LegMm,
                botCorner.Z);

            double topVerticalDropMm = (railCount >= 3) ? railSpacingMm : cl11LegMm;
            double botVerticalRiseMm = (railCount >= 3) ? railSpacingMm : cl11LegMm;

            Point topVerticalEnd = new Point(
                topCorner.X,
                topCorner.Y,
                topCorner.Z - topVerticalDropMm);

            Point botVerticalEnd = new Point(
                botCorner.X,
                botCorner.Y,
                botCorner.Z + botVerticalRiseMm);

            PolyBeam topLoop = CreateCl11PolyBeam(topRailJoint, topCorner, topVerticalEnd);
            PolyBeam botLoop = CreateCl11PolyBeam(botRailJoint, botCorner, botVerticalEnd);

            Beam topSplice = CreateRailSplice(topRailJoint, inwardDir, spliceLenMm);
            Beam botSplice = CreateRailSplice(botRailJoint, inwardDir, spliceLenMm);

            if (topLoop != null &&
                topSplice != null &&
                straightRailsByRow != null &&
                straightRailsByRow.Length > 0 &&
                straightRailsByRow[0] != null)
            {
                TryCreateRoundSpliceBolts(
                    straightRailsByRow[0],
                    topSplice,
                    topLoop,
                    topRailJoint);
            }

            if (botLoop != null &&
                botSplice != null &&
                straightRailsByRow != null &&
                straightRailsByRow.Length > (railCount - 1) &&
                straightRailsByRow[railCount - 1] != null)
            {
                TryCreateRoundSpliceBolts(
                    straightRailsByRow[railCount - 1],
                    botSplice,
                    botLoop,
                    botRailJoint);
            }

            if (railCount == 2)
            {
                Point mid = new Point(
                    topCorner.X,
                    topCorner.Y,
                    (topVerticalEnd.Z + botVerticalEnd.Z) * 0.5);

                Beam midSplice = CreateRailSplice(mid, upDir, spliceLenMm);

                if (midSplice != null && topLoop != null && botLoop != null)
                {
                    TryCreateRoundTSpliceBolts(
                        topLoop,
                        midSplice,
                        botLoop,
                        mid);
                }

                return;
            }

            var middleTeeNodes = new List<Point>();
            var verticalSpliceByRow = new Beam[railCount];
            var verticalLoopRails = new List<Beam>();

            for (int row = 1; row < railCount - 1; row++)
            {
                Point teeNode = GetRoundTerminalRowPoint(
                    side, onLine, along, lateral, railFromTopMm, railSpacingMm, row);

                middleTeeNodes.Add(teeNode);

                // Vertical tee splice
                Beam verticalSplice = CreateRailSplice(teeNode, upDir, teeVerticalSpliceLenMm);
                verticalSpliceByRow[row] = verticalSplice;

                Beam rowRail = null;
                if (straightRailsByRow != null && row >= 0 && row < straightRailsByRow.Length)
                    rowRail = straightRailsByRow[row];

                // Horizontal pipe + end splice
                CreateRoundTeeHorizontalPipeAndEndSplice(
                    teeNode,
                    inwardDir,
                    spliceLenMm,
                    teeVerticalSpliceLenMm,
                    spliceOdMm,
                    verticalSplice,
                    rowRail);
            }

            // For 4+ rails, add the vertical tube segments between the T splice nodes.
            if (railCount >= 4)
            {
                for (int i = 0; i < middleTeeNodes.Count - 1; i++)
                {
                    Beam vert = CreateRoundVerticalLoopRail(
                        middleTeeNodes[i],
                        middleTeeNodes[i + 1]);

                    verticalLoopRails.Add(vert);
                }
            }

            // Add bolts to the vertical T splices
            for (int row = 1; row < railCount - 1; row++)
            {
                Beam verticalSplice = verticalSpliceByRow[row];
                if (verticalSplice == null)
                    continue;

                Part upperPart;
                if (row == 1)
                    upperPart = topLoop;
                else
                    upperPart = verticalLoopRails[row - 2];

                Part lowerPart;
                if (row == railCount - 2)
                    lowerPart = botLoop;
                else
                    lowerPart = verticalLoopRails[row - 1];

                if (upperPart != null && lowerPart != null)
                {
                    TryCreateRoundTSpliceBolts(
                        upperPart,
                        verticalSplice,
                        lowerPart,
                        middleTeeNodes[row - 1]);
                }
            }
        }

        private Beam CreateRoundVerticalLoopRail(Point start, Point end)
        {
            if (start == null || end == null)
                return null;

            Beam rail = new Beam(start, end);

            rail.Profile.ProfileString = "PIPE1-1/2X14GA";
            rail.Material.MaterialString = "A53";
            rail.Class = "1";
            rail.Name = "CL11 VERT";

            rail.Position.Plane = Position.PlaneEnum.MIDDLE;
            rail.Position.Depth = Position.DepthEnum.MIDDLE;
            rail.Position.Rotation = Position.RotationEnum.TOP;

            if (!rail.Insert())
                return null;

            rail.Modify();
            return rail;
        }

        private void CreateRoundTeeHorizontalPipeAndEndSplice(
            Point teeNode,
            Vector inwardDir,
            double endSpliceLenMm,
            double copeCutLenMm,
            double verticalSpliceOdMm,
            Beam verticalSplice,
            Beam straightRailPart)
        {
            if (teeNode == null || inwardDir == null)
                return;

            Vector u = GetDirXYUnit(inwardDir);
            if (u == null)
                return;

            // Pipe starts 5/8" inward from vertical splice center
            double pipeStartOffsetMm = InchesToMm(0.625);

            // Pipe ends 10-1/4" inward from vertical splice center
            double pipeEndOffsetMm = InchesToMm(10.25);

            Point pipeStart = new Point(
                teeNode.X + u.X * pipeStartOffsetMm,
                teeNode.Y + u.Y * pipeStartOffsetMm,
                teeNode.Z);

            Point pipeEnd = new Point(
                teeNode.X + u.X * pipeEndOffsetMm,
                teeNode.Y + u.Y * pipeEndOffsetMm,
                teeNode.Z);

            Beam pipe = new Beam(pipeStart, pipeEnd);
            pipe.Profile.ProfileString = "PIPE1-1/2X14GA";
            pipe.Material.MaterialString = "A53";
            pipe.Class = "1";
            pipe.Name = "CL11 HORIZ";

            pipe.Position.Plane = Position.PlaneEnum.MIDDLE;
            pipe.Position.Depth = Position.DepthEnum.MIDDLE;
            pipe.Position.Rotation = Position.RotationEnum.TOP;

            if (!pipe.Insert())
                return;

            pipe.Modify();

            // Add weld between the horizontal pipe and the 8" vertical splice
            if (verticalSplice != null)
                CreateSimpleShopFilletWeld(verticalSplice, pipe, InchesToMm(0.125));

            // Cope the horizontal pipe around the vertical splice
            CreateRoundPipeCopeCut(pipe, teeNode, verticalSpliceOdMm, copeCutLenMm);

            // Then add the normal 4" splice at the far end
            Beam endSplice = CreateRailSplice(pipeEnd, u, endSpliceLenMm);

            if (endSplice != null && straightRailPart != null)
            {
                TryCreateRoundSpliceBolts(
                    pipe,
                    endSplice,
                    straightRailPart,
                    pipeEnd);
            }
        }

        private bool CreateSimpleShopFilletWeld(Part mainPart, Part secondaryPart, double sizeMm)
        {
            if (mainPart == null || secondaryPart == null || sizeMm <= 0.0)
                return false;

            Weld weld = new Weld();
            weld.MainObject = mainPart;
            weld.SecondaryObject = secondaryPart;
            weld.ShopWeld = true;
            weld.ConnectAssemblies = false;
            weld.AroundWeld = false;

            weld.SizeAbove = sizeMm;
            weld.SizeBelow = sizeMm;

            weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_FILLET;
            weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_FILLET;

            return weld.Insert();
        }

        private bool CreateRoundPipeCopeCut(
    Beam father,
    Point center,
    double cutDiameterMm,
    double cutLengthMm)
        {
            if (father == null || center == null || cutDiameterMm <= 0.0 || cutLengthMm <= 0.0)
                return false;

            double halfLenMm = cutLengthMm * 0.5;

            Point start = new Point(center.X, center.Y, center.Z - halfLenMm);
            Point end = new Point(center.X, center.Y, center.Z + halfLenMm);

            Beam operative = new Beam(start, end);
            operative.Profile.ProfileString = MakeMetricRoundProfileString(cutDiameterMm);
            operative.Material.MaterialString = "A36";
            operative.Class = BooleanPart.BooleanOperativeClassName;
            operative.Name = "ROUND PIPE COPE";

            operative.Position.Plane = Position.PlaneEnum.MIDDLE;
            operative.Position.Depth = Position.DepthEnum.MIDDLE;
            operative.Position.Rotation = Position.RotationEnum.TOP;

            if (!operative.Insert())
                return false;

            BooleanPart cut = new BooleanPart();
            cut.Father = father;
            cut.Type = BooleanPart.BooleanTypeEnum.BOOLEAN_CUT;

            if (!cut.SetOperativePart(operative))
            {
                operative.Delete();
                return false;
            }

            if (!cut.Insert())
            {
                operative.Delete();
                return false;
            }

            father.Modify();
            operative.Delete();
            return true;
        }

        private bool CreateRoundSpliceCopeCut(Beam father, Point center, double cutDiameterMm)
        {
            if (father == null || center == null || cutDiameterMm <= 0.0)
                return false;

            double halfLenMm = Math.Max(cutDiameterMm, InchesToMm(4.0));

            Point start = new Point(center.X, center.Y, center.Z - halfLenMm);
            Point end = new Point(center.X, center.Y, center.Z + halfLenMm);

            Beam operative = new Beam(start, end);
            operative.Profile.ProfileString = MakeMetricRoundProfileString(cutDiameterMm);
            operative.Material.MaterialString = "A36";
            operative.Class = BooleanPart.BooleanOperativeClassName;
            operative.Name = "ROUND SPLICE COPE";

            operative.Position.Plane = Position.PlaneEnum.MIDDLE;
            operative.Position.Depth = Position.DepthEnum.MIDDLE;
            operative.Position.Rotation = Position.RotationEnum.TOP;

            if (!operative.Insert())
                return false;

            BooleanPart cut = new BooleanPart();
            cut.Father = father;
            cut.Type = BooleanPart.BooleanTypeEnum.BOOLEAN_CUT;

            if (!cut.SetOperativePart(operative))
            {
                operative.Delete();
                return false;
            }

            if (!cut.Insert())
            {
                operative.Delete();
                return false;
            }

            father.Modify();
            operative.Delete();
            return true;
        }

        private static string MakeMetricRoundProfileString(double diameterMm)
        {
            return "D" + diameterMm.ToString("0.###", CultureInfo.InvariantCulture);
        }

        private Point GetRoundTerminalRowPoint(
            RailSideSpec side,
            Point onLine,
            double along,
            double lateral,
            double railFromTopMm,
            double railSpacingMm,
            int row)
        {
            Vector dirXY = GetDirXYUnit(side.Dir);

            return new Point(
                onLine.X + dirXY.X * along + side.Left.X * lateral,
                onLine.Y + dirXY.Y * along + side.Left.Y * lateral,
                side.FirstPostTopZ - railFromTopMm - (row * railSpacingMm));
        }

        private PolyBeam CreateCl11PolyBeam(
    Point railEndPoint,
    Point bendCornerPoint,
    Point verticalEndPoint)
        {
            if (railEndPoint == null || bendCornerPoint == null || verticalEndPoint == null)
                return null;

            PolyBeam pb = new PolyBeam();

            pb.Profile.ProfileString = "PIPE1-1/2X14GA";
            pb.Material.MaterialString = "A53";
            pb.Class = "1";
            pb.Name = "CL11";

            pb.Position.Plane = Position.PlaneEnum.MIDDLE;
            pb.Position.Depth = Position.DepthEnum.MIDDLE;
            pb.Position.Rotation = Position.RotationEnum.TOP;

            double bendRadiusMm = InchesToMm(3.5);
            Chamfer bendChamfer = new Chamfer(bendRadiusMm, 0.0, Chamfer.ChamferTypeEnum.CHAMFER_ROUNDING);

            pb.AddContourPoint(new ContourPoint(railEndPoint, null));
            pb.AddContourPoint(new ContourPoint(bendCornerPoint, bendChamfer));
            pb.AddContourPoint(new ContourPoint(verticalEndPoint, null));

            if (!pb.Insert())
                return null;

            pb.Modify();
            return pb;
        }

        private static bool TryGetRoundOutsideDimMm(string profile, out double outsideMm)
        {
            outsideMm = 0.0;
            if (string.IsNullOrWhiteSpace(profile))
                return false;

            string s = profile.Trim().ToUpperInvariant().Replace(" ", "");

            // D44.45
            Match mDirect = Regex.Match(s, @"^D([0-9]+(?:\.[0-9]+)?)$");
            if (mDirect.Success)
            {
                if (double.TryParse(mDirect.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out outsideMm))
                    return outsideMm > 0.001;
            }

            // PIPE1-1/2X14GA  -> use first dimension as OD
            if (s.StartsWith("PIPE"))
                s = s.Substring(4);

            string[] dims = s.Split(new[] { 'X' }, StringSplitOptions.RemoveEmptyEntries);
            if (dims.Length >= 1)
            {
                if (TryParseProfileDimInches(dims[0], out double odIn))
                {
                    outsideMm = InchesToMm(odIn);
                    return outsideMm > 0.001;
                }
            }

            return false;
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


        private List<Beam> CreateRailPiecesCollect(
            Point a,
            Point b,
            double maxLenMm,
            string profile,
            string material,
            string cls,
            string name,
            bool useInnerSplice)
        {
            var outPieces = new List<Beam>();

            Vector v = new Vector(b.X - a.X, b.Y - a.Y, b.Z - a.Z);
            double total = Math.Sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z);
            if (total < 1.0)
                return outPieces;

            if (maxLenMm <= 1.0)
                maxLenMm = total;

            int pieces = Math.Max(1, (int)Math.Ceiling(total / maxLenMm));

            Vector dir = new Vector(v.X / total, v.Y / total, v.Z / total);

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

                // Add splice at each internal break point
                if (i < pieces - 1)
                {
                    Point splicePoint = p1;

                    if (useInnerSplice)
                        CreateInnerRailSplice(splicePoint, dir, InchesToMm(4.0));
                    else
                        CreateRailSplice(splicePoint, dir, InchesToMm(4.0));
                }
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

        private static bool TryForceBeamEndToPlane(Beam target, bool atStart, Point facePoint, Vector faceNormal)
        {
            try
            {
                if (target == null || facePoint == null || faceNormal == null) return false;

                Point endPoint = atStart ? target.StartPoint : target.EndPoint;
                Point otherPoint = atStart ? target.EndPoint : target.StartPoint;

                Vector d = UnitVector(new Vector(
                    otherPoint.X - endPoint.X,
                    otherPoint.Y - endPoint.Y,
                    otherPoint.Z - endPoint.Z));

                Vector n = UnitVector(faceNormal);
                double den = d.X * n.X + d.Y * n.Y + d.Z * n.Z;
                if (Math.Abs(den) < 1e-9) return false;

                double num =
                    (facePoint.X - endPoint.X) * n.X +
                    (facePoint.Y - endPoint.Y) * n.Y +
                    (facePoint.Z - endPoint.Z) * n.Z;

                double t = num / den;

                Point moved = new Point(
                    endPoint.X + d.X * t,
                    endPoint.Y + d.Y * t,
                    endPoint.Z + d.Z * t);

                if (atStart) target.StartPoint = moved;
                else target.EndPoint = moved;

                return target.Modify();
            }
            catch
            {
                return false;
            }
        }

        private enum Dir4
        {
            East = 0,
            North = 1,
            West = 2,
            South = 3
        }

        private static Dir4 ToDir4(Vector d)
        {
            // dominant-axis binning from start->end vector
            if (Math.Abs(d.X) >= Math.Abs(d.Y))
                return (d.X >= 0.0) ? Dir4.East : Dir4.West;

            return (d.Y >= 0.0) ? Dir4.North : Dir4.South;
        }

        private static int GetCornerRotationIndex(Vector prevDir, Vector nextDir)
        {
            Dir4 inDir = ToDir4(prevDir);
            Dir4 outDir = ToDir4(nextDir);

            if (inDir == Dir4.East && (outDir == Dir4.North || outDir == Dir4.South)) return 0;
            if (inDir == Dir4.North && (outDir == Dir4.West || outDir == Dir4.East)) return 1;
            if (inDir == Dir4.West && (outDir == Dir4.South || outDir == Dir4.North)) return 2;
            if (inDir == Dir4.South && (outDir == Dir4.East || outDir == Dir4.West)) return 3;

            return (int)inDir;
        }

        private static string GetCornerDebugClassByRunDirection(Vector prevDir, Vector nextDir, bool isInside)
        {
            int rot = GetCornerRotationIndex(prevDir, nextDir); // 0..3

            // Outside: 81..84
            // Inside:  85..88
            int cls = isInside ? (85 + rot) : (81 + rot);
            return cls.ToString(CultureInfo.InvariantCulture);
        }

        private Beam CreateInnerRailSplice(Point center, Vector dir, double spliceLenMm)
        {
            if (center == null || dir == null || spliceLenMm <= 0.0)
                return null;

            Vector dirUnit = GetDirXYUnit(dir);
            if (dirUnit == null)
                return null;

            double halfLen = spliceLenMm * 0.5;

            Point start = new Point(
                center.X - dirUnit.X * halfLen,
                center.Y - dirUnit.Y * halfLen,
                center.Z - dirUnit.Z * halfLen);

            Point end = new Point(
                center.X + dirUnit.X * halfLen,
                center.Y + dirUnit.Y * halfLen,
                center.Z + dirUnit.Z * halfLen);

            Beam splice = new Beam(start, end);

            splice.Profile.ProfileString = "TS1-1/4X1-1/4X11 GA";
            splice.Material.MaterialString = "A50";
            splice.Class = "1";
            splice.Name = "INNER RAIL SPLICE";

            splice.PartNumber.Prefix = "SP1";
            splice.PartNumber.StartNumber = 0;
            splice.AssemblyNumber.Prefix = "SP1";
            splice.AssemblyNumber.StartNumber = 0;

            splice.Position.Plane = Position.PlaneEnum.MIDDLE;
            splice.Position.Depth = Position.DepthEnum.MIDDLE;
            splice.Position.Rotation = Position.RotationEnum.TOP;

            if (!splice.Insert())
                return null;

            splice.Modify();
            return splice;
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
                double hx, hy;

                if (TryGetClosestPointOnPartCenterlineXY(host, refPointOnLine, out Point cpt) && cpt != null)
                {
                    hx = cpt.X;
                    hy = cpt.Y;
                }
                else
                {
                    Solid s = host.GetSolid();
                    if (s == null) return 0;
                    hx = 0.5 * (s.MinimumPoint.X + s.MaximumPoint.X);
                    hy = 0.5 * (s.MinimumPoint.Y + s.MaximumPoint.Y);
                }

                double vx = hx - refPointOnLine.X;
                double vy = hy - refPointOnLine.Y;

                double dot = vx * leftUnit.X + vy * leftUnit.Y;
                if (Math.Abs(dot) < 1e-9) return 0;
                return (dot > 0.0) ? +1 : -1;
            }
            catch { return 0; }
        }

        private static int GetLineRefSign(string lineRef)
        {
            if (string.Equals(lineRef, "LEFT", StringComparison.OrdinalIgnoreCase)) return +1;
            if (string.Equals(lineRef, "RIGHT", StringComparison.OrdinalIgnoreCase)) return -1;
            return 0; // MIDDLE/blank
        }

        private static int ResolvePostSideSign(
            string lineRef,
            Vector leftUnit,
            Point refPointOnLine,
            Part hostForSide,
            bool hasDeckEdge)
        {
            int lineSign = GetLineRefSign(lineRef);
            int hostSide = DetermineConnectionSideSign(leftUnit, refPointOnLine, hostForSide);

            if (hostSide != 0)
            {
                if (lineSign > 0) return -hostSide;
                if (lineSign < 0) return +hostSide;

                // MIDDLE
                return hasDeckEdge ? -hostSide : 0;
            }

            if (lineSign != 0) return lineSign;
            return hasDeckEdge ? +1 : 0;
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
            int sideSign = 0;
            if (lineRef == "LEFT") sideSign = +1;
            else if (lineRef == "RIGHT") sideSign = -1;

            if (Math.Abs(deckEdgeMm) > 0.0001)
            {
                if (sideSign == 0) sideSign = +1;
                return sideSign * (deckEdgeMm + halfPostWidthMm);
            }

            return sideSign * halfPostWidthMm;
        }

        private Brep CreateRoundPostRpc(
    Point postTopCenter,
    Vector left,
    Point stationOnLine,
    Part nearestHost,
    double segPostLateralMm,
    double halfPostWidthMm)
        {
            if (postTopCenter == null)
                return null;

            double riseMm = InchesToMm(1.25);

            Point start = new Point(
                postTopCenter.X,
                postTopCenter.Y,
                postTopCenter.Z - riseMm);

            int faceSign = ResolveRpcFacingSign(left, stationOnLine, nearestHost, segPostLateralMm);

            Point end = new Point(
                start.X + left.X * (faceSign * halfPostWidthMm),
                start.Y + left.Y * (faceSign * halfPostWidthMm),
                start.Z);

            var rpc = new Brep
            {
                StartPoint = start,
                EndPoint = end
            };

            rpc.Profile.ProfileString = "RPC";
            rpc.Name = "RPC";
            rpc.Class = "6";
            rpc.Material.MaterialString = "A36";
            rpc.Position.Plane = Position.PlaneEnum.MIDDLE;
            rpc.Position.Depth = Position.DepthEnum.MIDDLE;
            rpc.Position.Rotation = Position.RotationEnum.TOP;

            if (!rpc.Insert())
            {
                Operation.DisplayPrompt("WDRailing: failed to insert RPC item.");
                return null;
            }

            rpc.Modify();
            return rpc;
        }

        private static int ResolveRpcFacingSign(
            Vector left,
            Point stationOnLine,
            Part nearestHost,
            double segPostLateralMm)
        {
            // Prefer the same side logic used for the seat-angle/host side.
            int sign = DetermineConnectionSideSign(left, stationOnLine, nearestHost);
            if (sign != 0)
                return sign;

            // Fallback: face back toward the reference line.
            // If post is on +LEFT side, reference line is toward -LEFT, and vice versa.
            if (Math.Abs(segPostLateralMm) > 0.001)
                return (segPostLateralMm > 0.0) ? -1 : +1;

            return +1;
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
            if (Math.Abs(dirUnit.X) >= Math.Abs(dirUnit.Y))
            {
                return (dirUnit.X >= 0.0)
                    ? Position.RotationEnum.TOP
                    : Position.RotationEnum.BELOW;
            }
            else
            {
                return (dirUnit.Y >= 0.0)
                    ? Position.RotationEnum.BACK
                    : Position.RotationEnum.FRONT;
            }
        }

        // ---------------- Profile parsing ----------------

    }
}