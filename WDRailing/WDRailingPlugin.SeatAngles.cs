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
        // ---------------- Seat angle (rail support) ----------------

        // User-specified properties
        private const string SEAT_ANGLE_NAME = "RAIL POST ANGLE";
        private const string SEAT_ANGLE_PROFILE = "L1-1/2X1-1/2X1/8";
        private const string SEAT_ANGLE_MATERIAL = "A1011-GR.50";
        private const string SEAT_ANGLE_CLASS = "6";

        // MUST match rail settings used in CreateRails (so the seat aligns to the rail centerline)
        private const string RAIL_PROFILE_FOR_SEAT = "TS1-1/2X1-1/2X.188";

        // If you ever need to nudge the seat up/down relative to the rail centerline,
        // tweak this (+ up, - down). Keep at 0 for now.
        private const double SEAT_Z_ADJUST_MM = 0.0;

        private static double GetSeatAngleLengthMm()
        {
            // Fixed: 1-1/2" long piece (not post width)
            return InchesToMm(1.5);
        }


        private static Vector GetDirXYUnit(Vector dirUnit)
        {
            // Keep seat angles horizontal even if the run has Z slope
            var v = new Vector(dirUnit.X, dirUnit.Y, 0.0);
            double len = Math.Sqrt(v.X * v.X + v.Y * v.Y);
            if (len < 1e-9) return new Vector(1.0, 0.0, 0.0);
            return new Vector(v.X / len, v.Y / len, 0.0);
        }


        private static double GetHalfRailWidthMm()
        {
            // half rail width (OD 1.5" fallback)
            double halfRailWidthMm = InchesToMm(1.5) * 0.5;
            if (TryGetOutsideDimMm(RAIL_PROFILE_FOR_SEAT, out var railOutsideMm))
                halfRailWidthMm = railOutsideMm * 0.5;
            return halfRailWidthMm;
        }


        private static Beam CreateRailPostSeatAngle(
    Point stationOnLine,
    Vector dirXYUnit,
    Vector leftUnit,
    double postLineLateralMm,
    double halfPostWidthMm,
    int railSideSign,
    double railCenterZmm,
    double halfRailDepthMm,
    Part postPart,
    double holeLineFromBendIn,
    double slotC2CIn,
    double slotSizeIn,
    string slotStandard,
    double slotCutLengthIn,
    bool slotSpecialFirstLayer,
    double pilotC2CIn,
    double pilotDiaIn,
    string pilotStandard,
    double pilotCutLengthIn)
        {
            try
            {
                // Seat dropped down by 1/2 rail depth (per your latest requirement)
                double zSeat = railCenterZmm - halfRailDepthMm;

                // Keep the angle located on the post face on the rail side
                double seatLateralMm = postLineLateralMm + railSideSign * halfPostWidthMm;

                // Seat reference point (on rail line in XY, lowered in Z)
                Point c = new Point(
                    stationOnLine.X + leftUnit.X * seatLateralMm,
                    stationOnLine.Y + leftUnit.Y * seatLateralMm,
                    zSeat);

                // Fixed seat length = 1-1/2"
                double L = InchesToMm(1.5);

                Point a = new Point(c.X - dirXYUnit.X * (L * 0.5), c.Y - dirXYUnit.Y * (L * 0.5), zSeat);
                Point b = new Point(c.X + dirXYUnit.X * (L * 0.5), c.Y + dirXYUnit.Y * (L * 0.5), zSeat);

                var seat = new Beam(a, b);
                seat.Name = SEAT_ANGLE_NAME;
                seat.Profile.ProfileString = SEAT_ANGLE_PROFILE;
                seat.Material.MaterialString = SEAT_ANGLE_MATERIAL;
                seat.Class = SEAT_ANGLE_CLASS;

                // Requested orientation
                seat.Position.Plane = Position.PlaneEnum.LEFT;
                seat.Position.Rotation = Position.RotationEnum.BELOW;
                seat.Position.Depth = Position.DepthEnum.BEHIND;

                if (!seat.Insert())
                    return null;

                // Add post-side slot + pilot-hole array
                // (best-effort, API-version-safe via reflection for slot-specific properties)
                TryAddSeatPostBolting(
                    seat,
                    postPart,
                    dirXYUnit,
                    holeLineFromBendIn,
                    slotC2CIn,
                    slotSizeIn,
                    slotStandard,
                    slotCutLengthIn,
                    slotSpecialFirstLayer,
                    pilotC2CIn,
                    pilotDiaIn,
                    pilotStandard,
                    pilotCutLengthIn
                );

                return seat;
            }
            catch
            {
                return null;
            }
        }


        private static void TryAddSeatPostBolting(
            Beam seat,
            Part postPart,
            Vector dirXYUnit,
            double holeLineFromBendIn,
    double slotC2CIn,
    double slotSizeIn,
    string slotStandard,
    double slotCutLengthIn,
    bool slotSpecialFirstLayer,
    double pilotC2CIn,
    double pilotDiaIn,
    string pilotStandard,
    double pilotCutLengthIn)
        {
            if (seat == null || postPart == null) return;

            try
            {
                // ---------------- geometry ----------------
                double slotWidthIn = slotSizeIn; // direct width/size from properties

                // Unit vector along seat length
                double dLen = Math.Sqrt(dirXYUnit.X * dirXYUnit.X + dirXYUnit.Y * dirXYUnit.Y + dirXYUnit.Z * dirXYUnit.Z);
                Vector ux = (dLen > 1e-9)
                    ? new Vector(dirXYUnit.X / dLen, dirXYUnit.Y / dLen, dirXYUnit.Z / dLen)
                    : new Vector(1.0, 0.0, 0.0);

                Vector up = new Vector(0.0, 0.0, 1.0);
                Vector side = new Vector(
                    ux.Y * up.Z - ux.Z * up.Y,
                    ux.Z * up.X - ux.X * up.Z,
                    ux.X * up.Y - ux.Y * up.X);
                double sideLen = Math.Sqrt(side.X * side.X + side.Y * side.Y + side.Z * side.Z);
                if (sideLen < 1e-9)
                    side = new Vector(1.0, 0.0, 0.0);
                else
                    side = new Vector(side.X / sideLen, side.Y / sideLen, side.Z / sideLen);

                // Seat center (global)
                Point mid = new Point(
                    (seat.StartPoint.X + seat.EndPoint.X) * 0.5,
                    (seat.StartPoint.Y + seat.EndPoint.Y) * 0.5,
                    (seat.StartPoint.Z + seat.EndPoint.Z) * 0.5);

                // Bias side-vector toward post center so offset goes into the post-side leg, not into the void.
                Point postMid = null;
                if (postPart is Beam pb)
                {
                    postMid = new Point(
                        (pb.StartPoint.X + pb.EndPoint.X) * 0.5,
                        (pb.StartPoint.Y + pb.EndPoint.Y) * 0.5,
                        (pb.StartPoint.Z + pb.EndPoint.Z) * 0.5);
                }
                else
                {
                    try
                    {
                        var cs = postPart.GetCoordinateSystem();
                        postMid = cs?.Origin;
                    }
                    catch { }
                }

                if (postMid != null)
                {
                    Vector toPost = new Vector(postMid.X - mid.X, postMid.Y - mid.Y, 0.0);
                    double dot = side.X * toPost.X + side.Y * toPost.Y + side.Z * toPost.Z;
                    if (dot < 0.0)
                        side = new Vector(-side.X, -side.Y, -side.Z);
                }

                double downMm = InchesToMm(holeLineFromBendIn);
                double legThicknessHalfMm = InchesToMm(0.125 * 0.5); // L1-1/2x1-1/2x1/8 -> 1/16"

                // Main hole reference point:
                //  - down from bend line
                //  - slight side bias into the leg thickness to avoid falling in the angle's inside void
                Point slotCenter = new Point(
                    mid.X - up.X * downMm + side.X * legThicknessHalfMm,
                    mid.Y - up.Y * downMm + side.Y * legThicknessHalfMm,
                    mid.Z - up.Z * downMm + side.Z * legThicknessHalfMm);

                // A small vector length to define bolt-group X direction robustly.
                double orientLenMm = Math.Max(5.0, InchesToMm(0.25));

                // ---------------- 1) SLOT IN SEAT ANGLE (POST LEG) ----------------
                // Hole-only (no physical fastener).
                // Uses bolt size + SlottedHoleX only (SlottedHoleY intentionally 0).
                TryInsertSeatSlotHole(
                    seat,
                    slotCenter,
                    ux,
                    orientLenMm,
                    slotStandard,
                    slotCutLengthIn,
                    slotWidthIn,
                    slotC2CIn,
                    slotSpecialFirstLayer,
                    Position.RotationEnum.BELOW);

                // ---------------- 1b) SLOT IN SEAT ANGLE (RAIL-SUPPORT LEG) ----------------
                // Mirror to the other leg: same slot properties, but located "out" from bend
                // (instead of "down") by the same input distance.
                Vector outFromPost = new Vector(-side.X, -side.Y, -side.Z);
                Point railLegSlotCenter = new Point(
                    mid.X + outFromPost.X * downMm - up.X * legThicknessHalfMm,
                    mid.Y + outFromPost.Y * downMm - up.Y * legThicknessHalfMm,
                    mid.Z + outFromPost.Z * downMm - up.Z * legThicknessHalfMm);

                // No pilot holes on this rail-support side (per request).
                // IMPORTANT: use the rail-leg center and Top rotation.
                bool railLegSlotOk = TryInsertSeatSlotHole(
                    seat,
                    railLegSlotCenter,
                    ux,
                    orientLenMm,
                    slotStandard,
                    slotCutLengthIn,
                    slotWidthIn,
                    slotC2CIn,
                    slotSpecialFirstLayer,
                    Position.RotationEnum.BACK);

                // Fallback: try opposite thickness bias if the first attempt fails.
                if (!railLegSlotOk)
                {
                    Point railLegSlotCenterAlt = new Point(
                        mid.X + outFromPost.X * downMm + up.X * legThicknessHalfMm,
                        mid.Y + outFromPost.Y * downMm + up.Y * legThicknessHalfMm,
                        mid.Z + outFromPost.Z * downMm + up.Z * legThicknessHalfMm);

                    TryInsertSeatSlotHole(
                        seat,
                        railLegSlotCenterAlt,
                        ux,
                        orientLenMm,
                        slotStandard,
                        slotCutLengthIn,
                        slotWidthIn,
                        slotC2CIn,
                        slotSpecialFirstLayer,
                        Position.RotationEnum.TOP);
                }

                // ---------------- 2) TWO PILOT HOLES IN POST ----------------
                {
                    double halfPilotCC = InchesToMm(pilotC2CIn * 0.5);
                    Point pilotFirst = new Point(
                        slotCenter.X - ux.X * halfPilotCC,
                        slotCenter.Y - ux.Y * halfPilotCC,
                        slotCenter.Z - ux.Z * halfPilotCC);

                    Point pilotOrient = new Point(
                        pilotFirst.X + ux.X * orientLenMm,
                        pilotFirst.Y + ux.Y * orientLenMm,
                        pilotFirst.Z + ux.Z * orientLenMm);

                    var pilot = new BoltArray
                    {
                        // REQUIRED: both set to post
                        PartToBoltTo = postPart,
                        PartToBeBolted = postPart,

                        FirstPosition = pilotFirst,
                        SecondPosition = pilotOrient,
                        BoltStandard = pilotStandard,
                        BoltSize = InchesToMm(pilotDiaIn),
                        Tolerance = 0.0,
                        Bolt = false
                    };

                    pilot.Position.Plane = Position.PlaneEnum.MIDDLE;
                    pilot.Position.Rotation = Position.RotationEnum.BELOW;
                    pilot.Position.Depth = Position.DepthEnum.MIDDLE;

                    ApplyNoAssemblyHardware(pilot);
                    TrySetDoubleProperty(pilot, "CutLength", InchesToMm(pilotCutLengthIn));

                    // 2-hole X array at user c/c
                    pilot.AddBoltDistX(InchesToMm(pilotC2CIn));
                    pilot.AddBoltDistY(0.0);

                    // Post-only
                    TrySetBoolProperty(pilot, "Hole1", true);
                    TrySetBoolProperty(pilot, "Hole2", false);
                    TrySetBoolProperty(pilot, "ConnectAssemblies", false);

                    TrySetDoubleProperty(pilot, "SlottedHoleX", 0.0);
                    TrySetDoubleProperty(pilot, "SlottedHoleY", 0.0);
                    TrySetEnumProperty(pilot, "HoleType", new[] { "HOLE_TYPE_STANDARD", "BOLT_HOLE_TYPE_STANDARD" });
                    TrySetEnumProperty(pilot, "PlainHoleType", new[] { "HOLE_TYPE_STANDARD", "BOLT_HOLE_TYPE_STANDARD" });

                    SetSpecialHoleLayerFlags(pilot, firstLayer: false, secondLayer: false);

                    pilot.Insert();
                }
            }
            catch
            {
                // Keep plugin resilient: seat angle insertion should not fail if bolting fails.
            }
        }


        private static bool TryInsertSeatSlotHole(
    Part seat,
    Point slotCenter,
    Vector ux,
    double orientLenMm,
    string slotStandard,
    double slotCutLengthIn,
    double slotWidthIn,
    double slotC2CIn,
    bool slotSpecialFirstLayer,
    Position.RotationEnum rotation = Position.RotationEnum.BELOW)
        {
            Point orientP2 = new Point(
                slotCenter.X + ux.X * orientLenMm,
                slotCenter.Y + ux.Y * orientLenMm,
                slotCenter.Z + ux.Z * orientLenMm);

            var slotHole = new BoltArray
            {
                // REQUIRED: both set to seat angle
                PartToBoltTo = seat,
                PartToBeBolted = seat,

                FirstPosition = slotCenter,
                SecondPosition = orientP2,
                BoltStandard = slotStandard,
                BoltSize = InchesToMm(slotWidthIn), // bolt size drives slot width
                Tolerance = 0.0,
                Bolt = false
            };

            slotHole.Position.Plane = Position.PlaneEnum.MIDDLE;
            slotHole.Position.Rotation = rotation; // BELOW for down slot, TOP for out slot
            slotHole.Position.Depth = Position.DepthEnum.MIDDLE;

            ApplyNoAssemblyHardware(slotHole);
            TrySetDoubleProperty(slotHole, "CutLength", InchesToMm(slotCutLengthIn));

            // one hole
            slotHole.AddBoltDistX(0.0);
            slotHole.AddBoltDistY(0.0);

            TrySetBoolProperty(slotHole, "Hole1", true);
            TrySetBoolProperty(slotHole, "Hole2", false);
            TrySetBoolProperty(slotHole, "ConnectAssemblies", false);

            // Only use bolt size + SlottedHoleX. SlottedHoleY intentionally unused.
            TrySetDoubleProperty(slotHole, "SlottedHoleX", InchesToMm(slotC2CIn));
            TrySetDoubleProperty(slotHole, "SlottedHoleY", 0.0);
            TrySetEnumProperty(slotHole, "HoleType", new[] { "HOLE_TYPE_SLOTTED", "BOLT_HOLE_TYPE_SLOTTED" });
            TrySetEnumProperty(slotHole, "PlainHoleType", new[] { "HOLE_TYPE_SLOTTED", "BOLT_HOLE_TYPE_SLOTTED" });

            SetSpecialHoleLayerFlags(slotHole, firstLayer: slotSpecialFirstLayer, secondLayer: false);

            return slotHole.Insert();
        }




        // ---------------- Corner seat angle (slots only, no pilot holes) ----------------

        private static Beam CreateCornerSeatAngleSlotsOnly(
            Point cornerRailPoint,
            Vector prevDirUnit,
            Vector nextDirUnit,
            bool isInsideCorner,
            double railCenterZmm,
            double halfRailDepthMm,
            double holeLineFromBendIn,
            double slotC2CIn,
            double slotSizeIn,
            string slotStandard,
            double slotCutLengthIn,
            bool slotSpecialFirstLayer,
            string cornerDebugClass = null)
        {
            try
            {
                if (cornerRailPoint == null) return null;

                Vector prevXY = GetDirXYUnit(prevDirUnit);
                Vector nextXY = GetDirXYUnit(nextDirUnit);

                // Vectors pointing away from the corner along each leg.
                Vector legA = new Vector(-prevXY.X, -prevXY.Y, 0.0); // back along previous side
                Vector legB = new Vector(nextXY.X, nextXY.Y, 0.0);   // forward along next side

                // "Inside" bisector of the two legs.
                Vector insideBis = new Vector(legA.X + legB.X, legA.Y + legB.Y, 0.0);
                double insideBisLen = Math.Sqrt(insideBis.X * insideBis.X + insideBis.Y * insideBis.Y);
                if (insideBisLen < 1e-9)
                {
                    // Near 180° fallback.
                    insideBis = new Vector(-prevXY.Y, prevXY.X, 0.0);
                    insideBisLen = Math.Sqrt(insideBis.X * insideBis.X + insideBis.Y * insideBis.Y);
                }
                insideBis = new Vector(insideBis.X / insideBisLen, insideBis.Y / insideBisLen, 0.0);
                Vector outsideBis = new Vector(-insideBis.X, -insideBis.Y, 0.0);

                // Angle between corner legs.
                double dot = legA.X * legB.X + legA.Y * legB.Y;
                if (dot > 1.0) dot = 1.0;
                else if (dot < -1.0) dot = -1.0;

                double phi = Math.Acos(dot);
                double s = Math.Sin(0.5 * phi);
                if (s < 1e-3) s = 1e-3;

                // Centerline-corner -> rail-corner radius.
                double cornerRadiusMm = halfRailDepthMm / s;

                Point insideCorner = new Point(
                    cornerRailPoint.X + insideBis.X * cornerRadiusMm,
                    cornerRailPoint.Y + insideBis.Y * cornerRadiusMm,
                    railCenterZmm);

                Point outsideCorner = new Point(
                    cornerRailPoint.X + outsideBis.X * cornerRadiusMm,
                    cornerRailPoint.Y + outsideBis.Y * cornerRadiusMm,
                    railCenterZmm);

                // Requested behavior:
                //  - outside corners: 1 tube-width in X & Y away from outside corner.
                //  - inside corners: 3/8" inset from inside corner.
                //
                // In practice the inside-corner handles were ending up ~1/4" too deep.
                // Apply a small outward correction so the handles sit on the intended corner.
                // For arbitrary corner angle, 1 tube-width in each leg maps to travel
                // of tubeWidth/sin(phi/2) along the bisector.
                double tubeWidthMm = 2.0 * halfRailDepthMm;
                double outsideTravelMm = tubeWidthMm / s;
                double insideInsetMm = InchesToMm(0.375);
                double insideOutCorrectionMm = InchesToMm(0.25);
                double effectiveInsideInsetMm = Math.Max(0.0, insideInsetMm - insideOutCorrectionMm);

                Vector slotAxisDir;
                double zSeat = railCenterZmm - halfRailDepthMm;
                Point c;

                if (isInsideCorner)
                {
                    c = new Point(
                        insideCorner.X - insideBis.X * effectiveInsideInsetMm,
                        insideCorner.Y - insideBis.Y * effectiveInsideInsetMm,
                        zSeat);

                    slotAxisDir = new Vector(-insideBis.X, -insideBis.Y, 0.0);
                }
                else
                {
                    c = new Point(
                        outsideCorner.X + insideBis.X * outsideTravelMm,
                        outsideCorner.Y + insideBis.Y * outsideTravelMm,
                        zSeat);

                    slotAxisDir = new Vector(insideBis.X, insideBis.Y, 0.0);
                }

                // Keep handles up/down so the angle stays on its side.
                double L = GetSeatAngleLengthMm();
                Point a = new Point(c.X, c.Y, c.Z - (L * 0.5));
                Point b = new Point(c.X, c.Y, c.Z + (L * 0.5));

                var seat = new Beam(a, b);
                seat.Name = SEAT_ANGLE_NAME;
                seat.Profile.ProfileString = SEAT_ANGLE_PROFILE;
                seat.Material.MaterialString = SEAT_ANGLE_MATERIAL;
                seat.Class = string.IsNullOrWhiteSpace(cornerDebugClass) ? SEAT_ANGLE_CLASS : cornerDebugClass;

                // Corner orientation is scenario-dependent.
                // We resolve it from leg directions so the result is not fixed to one orientation.
                ResolveCornerSeatOrientation(
                    legA,
                    legB,
                    prevXY,
                    nextXY,
                    isInsideCorner,
                    out Position.PlaneEnum plane,
                    out Position.RotationEnum vertical,
                    out Position.DepthEnum facing);

                seat.Position.Plane = plane;       // "horizontal" (LEFT/RIGHT)
                seat.Position.Rotation = vertical; // "vertical"   (TOP/BELOW)
                seat.Position.Depth = facing;      // final facing   (FRONT/BEHIND)

                if (!seat.Insert())
                    return null;

                TryAddCornerSlotsOnly(
                    seat,
                    slotAxisDir,
                    holeLineFromBendIn,
                    slotC2CIn,
                    slotSizeIn,
                    slotStandard,
                    slotCutLengthIn,
                    slotSpecialFirstLayer);

                return seat;
            }
            catch
            {
                return null;
            }
        }

        private static void ResolveCornerSeatOrientation(
            Vector legA,
            Vector legB,
            Vector prevXY,
            Vector nextXY,
            bool isInsideCorner,
            out Position.PlaneEnum plane,
            out Position.RotationEnum vertical,
            out Position.DepthEnum facing)
        {
            // Determine coarse compass occupancy from both legs that leave the corner.
            // east/west decides horizontal (Plane), north/south decides vertical (Rotation).
            double east = Math.Max(legA.X, legB.X);
            double west = Math.Max(-legA.X, -legB.X);
            double north = Math.Max(legA.Y, legB.Y);
            double south = Math.Max(-legA.Y, -legB.Y);

            const double axisTol = 0.15;
            bool hasEast = east > axisTol;
            bool hasWest = west > axisTol;
            bool hasNorth = north > axisTol;
            bool hasSouth = south > axisTol;

            // Horizontal handle:
            // - rail to the right (east) => LEFT
            // - rail to the left (west)  => RIGHT
            if (hasEast && !hasWest) plane = Position.PlaneEnum.LEFT;
            else if (hasWest && !hasEast) plane = Position.PlaneEnum.RIGHT;
            else
            {
                // Fallback when near-axis ambiguity exists.
                double sx = (legA.X + legB.X);
                plane = (sx >= 0.0) ? Position.PlaneEnum.LEFT : Position.PlaneEnum.RIGHT;
            }

            // Vertical handle:
            // - rail to the north => BELOW ("down")
            // - rail to the south => TOP
            if (hasNorth && !hasSouth) vertical = Position.RotationEnum.BELOW;
            else if (hasSouth && !hasNorth) vertical = Position.RotationEnum.TOP;
            else
            {
                double sy = (legA.Y + legB.Y);
                vertical = (sy >= 0.0) ? Position.RotationEnum.BELOW : Position.RotationEnum.TOP;
            }

            // Facing (Depth FRONT/BEHIND):
            // Use quadrant when it is clear, then flip for inside corners.
            bool clearX = hasEast ^ hasWest;
            bool clearY = hasNorth ^ hasSouth;
            if (clearX && clearY)
            {
                bool eastSide = hasEast;
                bool northSide = hasNorth;
                bool baseFront = (eastSide == northSide); // NE/SW vs NW/SE
                if (isInsideCorner) baseFront = !baseFront;
                facing = baseFront ? Position.DepthEnum.FRONT : Position.DepthEnum.BEHIND;
            }
            else
            {
                double turnZ = prevXY.X * nextXY.Y - prevXY.Y * nextXY.X;
                bool leftTurn = (turnZ >= 0.0);
                facing = (leftTurn ^ isInsideCorner)
                    ? Position.DepthEnum.BEHIND
                    : Position.DepthEnum.FRONT;
            }
        }


        private static void TryAddCornerSlotsOnly(

            Beam seat,
            Vector slotAxisDir,
            double holeLineFromBendIn,
            double slotC2CIn,
            double slotSizeIn,
            string slotStandard,
            double slotCutLengthIn,
            bool slotSpecialFirstLayer)
        {
            if (seat == null) return;

            try
            {
                Vector ux = GetDirXYUnit(slotAxisDir);
                Vector up = new Vector(0.0, 0.0, 1.0);

                // side = ux x up
                Vector side = new Vector(
                    ux.Y * up.Z - ux.Z * up.Y,
                    ux.Z * up.X - ux.X * up.Z,
                    ux.X * up.Y - ux.Y * up.X);

                double sideLen = Math.Sqrt(side.X * side.X + side.Y * side.Y + side.Z * side.Z);
                if (sideLen < 1e-9)
                    side = new Vector(1.0, 0.0, 0.0);
                else
                    side = new Vector(side.X / sideLen, side.Y / sideLen, side.Z / sideLen);

                Point mid = new Point(
                    (seat.StartPoint.X + seat.EndPoint.X) * 0.5,
                    (seat.StartPoint.Y + seat.EndPoint.Y) * 0.5,
                    (seat.StartPoint.Z + seat.EndPoint.Z) * 0.5);

                double bendOffsetMm = InchesToMm(holeLineFromBendIn);
                double legThicknessHalfMm = InchesToMm(0.125 * 0.5);
                double orientLenMm = Math.Max(5.0, InchesToMm(0.25));

                // leg 1 slot (down from bend)
                Point slot1 = new Point(
                    mid.X - up.X * bendOffsetMm + side.X * legThicknessHalfMm,
                    mid.Y - up.Y * bendOffsetMm + side.Y * legThicknessHalfMm,
                    mid.Z - up.Z * bendOffsetMm + side.Z * legThicknessHalfMm);

                TryInsertSeatSlotHole(
                    seat,
                    slot1,
                    ux,
                    orientLenMm,
                    slotStandard,
                    slotCutLengthIn,
                    slotSizeIn,
                    slotC2CIn,
                    slotSpecialFirstLayer,
                    Position.RotationEnum.BELOW);

                // leg 2 slot (out from bend, opposite side)
                Vector outFromLeg1 = new Vector(-side.X, -side.Y, -side.Z);
                Point slot2 = new Point(
                    mid.X + outFromLeg1.X * bendOffsetMm - up.X * legThicknessHalfMm,
                    mid.Y + outFromLeg1.Y * bendOffsetMm - up.Y * legThicknessHalfMm,
                    mid.Z + outFromLeg1.Z * bendOffsetMm - up.Z * legThicknessHalfMm);

                bool ok2 = TryInsertSeatSlotHole(
                    seat,
                    slot2,
                    ux,
                    orientLenMm,
                    slotStandard,
                    slotCutLengthIn,
                    slotSizeIn,
                    slotC2CIn,
                    slotSpecialFirstLayer,
                    Position.RotationEnum.BACK);

                // fallback orientation for API/version differences
                if (!ok2)
                {
                    Point slot2Alt = new Point(
                        mid.X + outFromLeg1.X * bendOffsetMm + up.X * legThicknessHalfMm,
                        mid.Y + outFromLeg1.Y * bendOffsetMm + up.Y * legThicknessHalfMm,
                        mid.Z + outFromLeg1.Z * bendOffsetMm + up.Z * legThicknessHalfMm);

                    TryInsertSeatSlotHole(
                        seat,
                        slot2Alt,
                        ux,
                        orientLenMm,
                        slotStandard,
                        slotCutLengthIn,
                        slotSizeIn,
                        slotC2CIn,
                        slotSpecialFirstLayer,
                        Position.RotationEnum.TOP);
                }
            }
            catch
            {
                // best-effort only
            }
        }


        private static void ApplyNoAssemblyHardware(BoltArray boltArray)
        {
            if (boltArray == null) return;

            // Explicitly force "nothing selected" in bolt assembly style options (best effort)
            TrySetBoolProperty(boltArray, "Bolt", false);
            TrySetBoolProperty(boltArray, "Nut1", false);
            TrySetBoolProperty(boltArray, "Nut2", false);
            TrySetBoolProperty(boltArray, "Washer1", false);
            TrySetBoolProperty(boltArray, "Washer2", false);
            TrySetBoolProperty(boltArray, "Washer3", false);
            TrySetBoolProperty(boltArray, "HoleOnly", true);
        }


        private static void SetSpecialHoleLayerFlags(object obj, bool firstLayer, bool secondLayer)
        {
            if (obj == null) return;

            // Different Tekla versions expose different names, so try a broad set safely.
            var firstNames = new[]
            {
                "SpecialHoleInFirstLayer",
                "SpecialHoleInPart1",
                "SpecialHoleFirst",
                "SpecialHole1"
            };

            var secondNames = new[]
            {
                "SpecialHoleInSecondLayer",
                "SpecialHoleInPart2",
                "SpecialHoleSecond",
                "SpecialHole2"
            };

            foreach (var n in firstNames) TrySetBoolProperty(obj, n, firstLayer);
            foreach (var n in secondNames) TrySetBoolProperty(obj, n, secondLayer);
        }


        private static void TrySetDoubleProperty(object obj, string propName, double value)
        {
            if (obj == null) return;
            var p = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
            if (p == null || !p.CanWrite) return;

            try
            {
                if (p.PropertyType == typeof(double)) p.SetValue(obj, value, null);
                else if (p.PropertyType == typeof(float)) p.SetValue(obj, (float)value, null);
                else if (p.PropertyType == typeof(int)) p.SetValue(obj, (int)Math.Round(value), null);
            }
            catch { }
        }


        private static void TrySetBoolProperty(object obj, string propName, bool value)
        {
            if (obj == null) return;
            var p = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
            if (p == null || !p.CanWrite || p.PropertyType != typeof(bool)) return;

            try { p.SetValue(obj, value, null); } catch { }
        }


        private static void TrySetEnumProperty(object obj, string propName, string[] enumNames)
        {
            if (obj == null) return;
            var p = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
            if (p == null || !p.CanWrite || !p.PropertyType.IsEnum) return;

            foreach (var n in enumNames)
            {
                try
                {
                    var v = Enum.Parse(p.PropertyType, n, ignoreCase: true);
                    p.SetValue(obj, v, null);
                    return;
                }
                catch
                {
                    // try next enum name
                }
            }
        }


    }
}
