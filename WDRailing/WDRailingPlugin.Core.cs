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
    [Plugin("WDRailingPlugin")]
    [PluginUserInterface("WDRailing.WDRailingDialog")]
    public partial class WDRailingPlugin : PluginBase
    {
        private readonly Model _model;
        private readonly WDRailingData _data;

        public WDRailingPlugin(WDRailingData data)
        {
            _data = data ?? new WDRailingData();
            _model = new Model();

            // IMPORTANT: Persist defaults at INSERT time so properties open populated even if equal to defaults.
            // NO hardcoded defaults. If config missing/invalid -> throw (user requested strict).
            var cfg = WDRailingDefaults.LoadOrThrow();

            if (string.IsNullOrWhiteSpace(_data.SpacingIn)) _data.SpacingIn = cfg.SpacingIn;
            if (string.IsNullOrWhiteSpace(_data.PostHeightIn)) _data.PostHeightIn = cfg.PostHeightIn;
            if (string.IsNullOrWhiteSpace(_data.StartOffsetIn)) _data.StartOffsetIn = cfg.StartOffsetIn;
            if (string.IsNullOrWhiteSpace(_data.EndOffsetIn)) _data.EndOffsetIn = cfg.EndOffsetIn;
            if (string.IsNullOrWhiteSpace(_data.BaseOffsetIn)) _data.BaseOffsetIn = cfg.BaseOffsetIn;
            if (string.IsNullOrWhiteSpace(_data.LineRef)) _data.LineRef = cfg.LineRef;
            if (string.IsNullOrWhiteSpace(_data.DeckEdgeIn)) _data.DeckEdgeIn = cfg.DeckEdgeIn;

            if (string.IsNullOrWhiteSpace(_data.PostProfile)) _data.PostProfile = cfg.PostProfile;
            if (string.IsNullOrWhiteSpace(_data.PostMaterial)) _data.PostMaterial = cfg.PostMaterial;
            if (string.IsNullOrWhiteSpace(_data.PostClass)) _data.PostClass = cfg.PostClass;
            if (string.IsNullOrWhiteSpace(_data.PostName)) _data.PostName = cfg.PostName;

            if (string.IsNullOrWhiteSpace(_data.ConnEnabled)) _data.ConnEnabled = cfg.CreateConnection;
            if (string.IsNullOrWhiteSpace(_data.ConnName)) _data.ConnName = cfg.ConnectionName;
            if (_data.ConnAttr == null) _data.ConnAttr = cfg.ConnectionAttr ?? "";

            if (string.IsNullOrWhiteSpace(_data.RailEnabled)) _data.RailEnabled = cfg.RailEnabled;
            if (string.IsNullOrWhiteSpace(_data.RailStartOffsetIn)) _data.RailStartOffsetIn = cfg.RailStartOffsetIn;
            if (string.IsNullOrWhiteSpace(_data.RailEndOffsetIn)) _data.RailEndOffsetIn = cfg.RailEndOffsetIn;
            if (string.IsNullOrWhiteSpace(_data.RailFromTopIn)) _data.RailFromTopIn = cfg.RailFromTopIn;
            if (string.IsNullOrWhiteSpace(_data.RailCount)) _data.RailCount = cfg.RailCount;
            if (string.IsNullOrWhiteSpace(_data.RailSpacingIn)) _data.RailSpacingIn = cfg.RailSpacingIn;
            if (string.IsNullOrWhiteSpace(_data.SeatHoleLineIn)) _data.SeatHoleLineIn = cfg.SeatHoleLineFromBendIn;

            if (string.IsNullOrWhiteSpace(_data.SeatSlotC2CIn)) _data.SeatSlotC2CIn = cfg.SeatSlotC2CIn;
            if (string.IsNullOrWhiteSpace(_data.SeatSlotSizeIn)) _data.SeatSlotSizeIn = cfg.SeatSlotSizeIn;
            if (string.IsNullOrWhiteSpace(_data.SeatSlotStandard)) _data.SeatSlotStandard = cfg.SeatSlotStandard;
            if (string.IsNullOrWhiteSpace(_data.SeatSlotCutLengthIn)) _data.SeatSlotCutLengthIn = cfg.SeatSlotCutLengthIn;
            if (string.IsNullOrWhiteSpace(_data.SeatSlotSpecial1)) _data.SeatSlotSpecial1 = cfg.SeatSlotSpecial1;

            if (string.IsNullOrWhiteSpace(_data.SeatPilotC2CIn)) _data.SeatPilotC2CIn = cfg.SeatPilotC2CIn;
            if (string.IsNullOrWhiteSpace(_data.SeatPilotDiaIn)) _data.SeatPilotDiaIn = cfg.SeatPilotDiaIn;
            if (string.IsNullOrWhiteSpace(_data.SeatPilotStandard)) _data.SeatPilotStandard = cfg.SeatPilotStandard;
            if (string.IsNullOrWhiteSpace(_data.SeatPilotCutLengthIn)) _data.SeatPilotCutLengthIn = cfg.SeatPilotCutLengthIn;

        }


        public override List<InputDefinition> DefineInput()
        {
            var picker = new Picker();
            var inputs = new List<InputDefinition>();

            // Pick multiple points in ONE Tekla picking command (user can stop after 2 points or keep going)
            ArrayList ptsPayload;
            try
            {
                ptsPayload = picker.PickPoints(
                    Picker.PickPointEnum.PICK_POLYGON,
                    "Pick run points (pick points, then MIDDLE-click to finish). You may stop after 2 points."
                );
            }
            catch (Exception ex)
            {
                // If user cancels the picking command, cancel insertion gracefully
                if (IsPickerCancel(ex)) return null;
                throw;
            }

            if (ptsPayload == null || ptsPayload.Count < 2)
                throw new InvalidOperationException("You must pick at least two points.");

            // Convert payload -> List<Point> for your persistence + Run logic
            var pts = new List<Point>();
            foreach (object o in ptsPayload)
                if (o is Point p) pts.Add(p);

            if (pts.Count < 2)
                throw new InvalidOperationException("You must pick at least two points.");

            // Persist polyline for Modify (RUNPTS1..4) + legacy P1/P2
            SaveRunPointsToData(pts);

            // Input #1: ALL picked points (handles)
            inputs.Add(new InputDefinition(ptsPayload));

            // Input #2: pick hosts (this already finishes with middle-click)
            ModelObjectEnumerator picked = picker.PickObjects(
                Picker.PickObjectsEnum.PICK_N_PARTS,
                "Pick framing members (middle-click to finish)"
            );

            var ids = new ArrayList();
            var hostIdInts = new List<int>();

            while (picked.MoveNext())
            {
                if (picked.Current is Part part)
                {
                    ids.Add(part.Identifier);
                    hostIdInts.Add(part.Identifier.ID);
                }
            }

            inputs.Add(new InputDefinition(ids));
            SetHostCsvToData(JoinIds(hostIdInts));

            return inputs;
        }


        public override bool Run(List<InputDefinition> input)
        {
            try
            {
                if (!_model.GetConnectionStatus())
                    return KeepWithPrompt("No model connection.");

                // Strict config only
                WDRailingDefaults cfg = WDRailingDefaults.LoadOrThrow();

                // ===================== Resolve polyline points + host payload =====================
                List<Point> runPts = null;
                object hostPayload = null;

                // NEW schema: input[0] = ArrayList of Points (>=2), input[1] = host payload (ArrayList of Identifiers)
                if (input != null && input.Count >= 1)
                {
                    var ptsPayload = input[0].GetInput() as ArrayList;
                    if (ptsPayload != null && ptsPayload.Count >= 2 && ptsPayload[0] is Point)
                    {
                        runPts = new List<Point>();
                        foreach (var o in ptsPayload)
                            if (o is Point pt) runPts.Add(pt);

                        if (input.Count >= 2) hostPayload = input[1].GetInput();

                        // Persist for Modify (requires you to add RunPts fields + hidden bindings)
                        SaveRunPointsToData(runPts);
                    }
                }

                // Modify fallback
                if (runPts == null || runPts.Count < 2)
                {
                    runPts = LoadRunPointsFromData();
                    if (runPts == null || runPts.Count < 2)
                    {
                        // Legacy P1/P2 fallback
                        Point lp1, lp2;
                        if (!TryReadStoredPoints(out lp1, out lp2))
                            throw new InvalidOperationException("Could not resolve points on Modify.");
                        runPts = new List<Point> { lp1, lp2 };
                    }
                }

                // Remove last point if it exactly repeats the first (common “closed loop” picking habit)
                if (runPts.Count >= 3 && AreSamePoint(runPts[0], runPts[runPts.Count - 1], 0.5))
                    runPts.RemoveAt(runPts.Count - 1);

                int sideCount = Math.Max(0, runPts.Count - 1);
                if (sideCount < 1)
                    throw new InvalidOperationException("Need at least 2 points.");

                // ===================== Resolve host parts =====================
                List<Part> hostParts = ResolveHostPartsFromPayload(hostPayload);
                if (hostParts.Count == 0)
                {
                    string storedCsv = GetHostCsvFromData();
                    if (!string.IsNullOrWhiteSpace(storedCsv))
                        hostParts = ResolveHostPartsFromCsv(storedCsv);
                }

                if (hostParts.Count > 0)
                {
                    var idInts = new List<int>();
                    foreach (Part hp in hostParts) idInts.Add(hp.Identifier.ID);
                    SetHostCsvToData(JoinIds(idInts));
                }

                // ===================== Effective values (instance or config) =====================
                double spacingIn = ParseImperialInchesOrThrow(string.IsNullOrWhiteSpace(_data.SpacingIn) ? cfg.SpacingIn : _data.SpacingIn, allowNegative: false);
                double heightIn = ParseImperialInchesOrThrow(string.IsNullOrWhiteSpace(_data.PostHeightIn) ? cfg.PostHeightIn : _data.PostHeightIn, allowNegative: false);
                double startIn = ParseImperialInchesOrThrow(string.IsNullOrWhiteSpace(_data.StartOffsetIn) ? cfg.StartOffsetIn : _data.StartOffsetIn, allowNegative: true);
                double endIn = ParseImperialInchesOrThrow(string.IsNullOrWhiteSpace(_data.EndOffsetIn) ? cfg.EndOffsetIn : _data.EndOffsetIn, allowNegative: true);
                double baseIn = ParseImperialInchesOrThrow(string.IsNullOrWhiteSpace(_data.BaseOffsetIn) ? cfg.BaseOffsetIn : _data.BaseOffsetIn, allowNegative: true);
                double deckEdgeIn = ParseImperialInchesOrThrow(string.IsNullOrWhiteSpace(_data.DeckEdgeIn) ? cfg.DeckEdgeIn : _data.DeckEdgeIn, allowNegative: true);

                string lineRef = (string.IsNullOrWhiteSpace(_data.LineRef) ? cfg.LineRef : _data.LineRef).Trim().ToUpperInvariant();
                if (lineRef != "LEFT" && lineRef != "RIGHT" && lineRef != "MIDDLE")
                    throw new InvalidDataException("LINE_REF must be LEFT, RIGHT, or MIDDLE. Got: " + lineRef);

                string profile = (string.IsNullOrWhiteSpace(_data.PostProfile) ? cfg.PostProfile : _data.PostProfile).Trim();
                string material = (string.IsNullOrWhiteSpace(_data.PostMaterial) ? cfg.PostMaterial : _data.PostMaterial).Trim();
                string postClass = (string.IsNullOrWhiteSpace(_data.PostClass) ? cfg.PostClass : _data.PostClass).Trim();
                string postName = (string.IsNullOrWhiteSpace(_data.PostName) ? cfg.PostName : _data.PostName).Trim();

                // Connection
                string connEnabledRaw = (string.IsNullOrWhiteSpace(_data.ConnEnabled) ? cfg.CreateConnection : _data.ConnEnabled).Trim();
                if (connEnabledRaw != "0" && connEnabledRaw != "1")
                    throw new InvalidDataException("CONN_ENABLED must be 0 or 1. Got: " + connEnabledRaw);
                bool connEnabled = (connEnabledRaw == "1");

                string connName = (string.IsNullOrWhiteSpace(_data.ConnName) ? cfg.ConnectionName : _data.ConnName).Trim();
                if (connEnabled && string.IsNullOrWhiteSpace(connName))
                    throw new InvalidDataException("ConnectionName cannot be blank when CreateConnection=1.");

                string connAttr = (string.IsNullOrWhiteSpace(_data.ConnAttr) ? (cfg.ConnectionAttr ?? "") : _data.ConnAttr).Trim();

                // Rails
                string railEnabledRaw = (string.IsNullOrWhiteSpace(_data.RailEnabled) ? cfg.RailEnabled : _data.RailEnabled).Trim();
                if (railEnabledRaw != "0" && railEnabledRaw != "1")
                    throw new InvalidDataException("RAIL_ENABLED must be 0 or 1. Got: " + railEnabledRaw);
                bool railEnabled = (railEnabledRaw == "1");

                double railStartIn = ParseImperialInchesOrThrow(string.IsNullOrWhiteSpace(_data.RailStartOffsetIn) ? cfg.RailStartOffsetIn : _data.RailStartOffsetIn, allowNegative: true);
                double railEndIn = ParseImperialInchesOrThrow(string.IsNullOrWhiteSpace(_data.RailEndOffsetIn) ? cfg.RailEndOffsetIn : _data.RailEndOffsetIn, allowNegative: true);
                double railFromTopIn = ParseImperialInchesOrThrow(string.IsNullOrWhiteSpace(_data.RailFromTopIn) ? cfg.RailFromTopIn : _data.RailFromTopIn, allowNegative: false);

                string railCountRaw = (string.IsNullOrWhiteSpace(_data.RailCount) ? cfg.RailCount : _data.RailCount).Trim();
                if (!int.TryParse(railCountRaw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int railCount))
                    throw new InvalidDataException("RAIL_COUNT must be an integer. Got: " + railCountRaw);
                railCount = Math.Max(0, railCount);

                double railSpacingIn = ParseImperialInchesOrThrow(
                    string.IsNullOrWhiteSpace(_data.RailSpacingIn) ? cfg.RailSpacingIn : _data.RailSpacingIn,
                    allowNegative: false);

                double seatHoleLineIn = ParseImperialInchesOrThrow(
                    string.IsNullOrWhiteSpace(_data.SeatHoleLineIn) ? cfg.SeatHoleLineFromBendIn : _data.SeatHoleLineIn,
                    allowNegative: false);

                double seatSlotC2CIn = ParseImperialInchesOrThrow(
                    string.IsNullOrWhiteSpace(_data.SeatSlotC2CIn) ? cfg.SeatSlotC2CIn : _data.SeatSlotC2CIn,
                    allowNegative: false);

                double seatSlotSizeIn = ParseImperialInchesOrThrow(
                    string.IsNullOrWhiteSpace(_data.SeatSlotSizeIn) ? cfg.SeatSlotSizeIn : _data.SeatSlotSizeIn,
                    allowNegative: false);

                string seatSlotStandard = (string.IsNullOrWhiteSpace(_data.SeatSlotStandard)
                    ? cfg.SeatSlotStandard
                    : _data.SeatSlotStandard).Trim();
                if (string.IsNullOrWhiteSpace(seatSlotStandard))
                    throw new InvalidDataException("SEAT_SLOT_STANDARD cannot be blank.");

                double seatSlotCutLenIn = ParseImperialInchesOrThrow(
                    string.IsNullOrWhiteSpace(_data.SeatSlotCutLengthIn) ? cfg.SeatSlotCutLengthIn : _data.SeatSlotCutLengthIn,
                    allowNegative: false);

                string seatSlotSpecial1Raw = (string.IsNullOrWhiteSpace(_data.SeatSlotSpecial1) ? cfg.SeatSlotSpecial1 : _data.SeatSlotSpecial1).Trim();
                if (seatSlotSpecial1Raw != "0" && seatSlotSpecial1Raw != "1")
                    throw new InvalidDataException("SEAT_SLOT_SPECIAL1 must be 0 or 1. Got: " + seatSlotSpecial1Raw);
                bool seatSlotSpecial1 = (seatSlotSpecial1Raw == "1");

                double seatPilotC2CIn = ParseImperialInchesOrThrow(
                    string.IsNullOrWhiteSpace(_data.SeatPilotC2CIn) ? cfg.SeatPilotC2CIn : _data.SeatPilotC2CIn,
                    allowNegative: false);

                double seatPilotDiaIn = ParseImperialInchesOrThrow(
                    string.IsNullOrWhiteSpace(_data.SeatPilotDiaIn) ? cfg.SeatPilotDiaIn : _data.SeatPilotDiaIn,
                    allowNegative: false);

                string seatPilotStandard = (string.IsNullOrWhiteSpace(_data.SeatPilotStandard)
                    ? cfg.SeatPilotStandard
                    : _data.SeatPilotStandard).Trim();
                if (string.IsNullOrWhiteSpace(seatPilotStandard))
                    throw new InvalidDataException("SEAT_PILOT_STANDARD cannot be blank.");

                double seatPilotCutLenIn = ParseImperialInchesOrThrow(
                    string.IsNullOrWhiteSpace(_data.SeatPilotCutLengthIn) ? cfg.SeatPilotCutLengthIn : _data.SeatPilotCutLengthIn,
                    allowNegative: false);

                // Persist effective values
                _data.SpacingIn = DistanceFormat.ToTeklaFeetInches(spacingIn, 16);
                _data.PostHeightIn = DistanceFormat.ToTeklaFeetInches(heightIn, 16);
                _data.StartOffsetIn = DistanceFormat.ToTeklaFeetInches(startIn, 16);
                _data.EndOffsetIn = DistanceFormat.ToTeklaFeetInches(endIn, 16);
                _data.BaseOffsetIn = DistanceFormat.ToTeklaFeetInches(baseIn, 16);
                _data.DeckEdgeIn = DistanceFormat.ToTeklaFeetInches(deckEdgeIn, 16);
                _data.LineRef = lineRef;

                _data.PostProfile = profile;
                _data.PostMaterial = material;
                _data.PostClass = postClass;
                _data.PostName = postName;

                _data.ConnEnabled = connEnabledRaw;
                _data.ConnName = connName;
                _data.ConnAttr = connAttr;

                _data.RailEnabled = railEnabledRaw;
                _data.RailCount = railCount.ToString(CultureInfo.InvariantCulture);
                _data.RailSpacingIn = DistanceFormat.ToTeklaFeetInches(railSpacingIn, 16);

                _data.RailStartOffsetIn = (railStartIn < 0 ? "-" : "") + DistanceFormat.ToTeklaFeetInches(Math.Abs(railStartIn), 16);
                _data.RailEndOffsetIn = (railEndIn < 0 ? "-" : "") + DistanceFormat.ToTeklaFeetInches(Math.Abs(railEndIn), 16);
                _data.RailFromTopIn = DistanceFormat.ToTeklaFeetInches(railFromTopIn, 16);

                _data.SeatHoleLineIn = DistanceFormat.ToTeklaFeetInches(seatHoleLineIn, 16);

                _data.SeatSlotC2CIn = DistanceFormat.ToTeklaFeetInches(seatSlotC2CIn, 16);
                _data.SeatSlotSizeIn = DistanceFormat.ToTeklaFeetInches(seatSlotSizeIn, 16);
                _data.SeatSlotStandard = seatSlotStandard;
                _data.SeatSlotCutLengthIn = DistanceFormat.ToTeklaFeetInches(seatSlotCutLenIn, 16);
                _data.SeatSlotSpecial1 = seatSlotSpecial1Raw;

                _data.SeatPilotC2CIn = DistanceFormat.ToTeklaFeetInches(seatPilotC2CIn, 16);
                _data.SeatPilotDiaIn = DistanceFormat.ToTeklaFeetInches(seatPilotDiaIn, 16);
                _data.SeatPilotStandard = seatPilotStandard;
                _data.SeatPilotCutLengthIn = DistanceFormat.ToTeklaFeetInches(seatPilotCutLenIn, 16);

                // ===================== Convert to mm =====================
                double spacingMm = InchesToMm(spacingIn);
                double postHeightMm = InchesToMm(heightIn);
                double startOffsetMmAll = InchesToMm(startIn);
                double endOffsetMmAll = InchesToMm(endIn);
                double baseOffsetMm = InchesToMm(baseIn);
                double deckEdgeMm = InchesToMm(deckEdgeIn);

                double railStartOffsetMm = InchesToMm(railStartIn);  // keep sign
                double railEndOffsetMm = InchesToMm(railEndIn);    // keep sign
                double railFromTopMm = InchesToMm(railFromTopIn);
                double railSpacingMm = InchesToMm(railSpacingIn);

                // Half-width for face-based offsets (best effort)
                double halfPostWidthMm = 0.0;
                if (TryGetOutsideDimMm(profile, out var outsideMm)) halfPostWidthMm = outsideMm * 0.5;

                // magnitude only; direction comes from per-segment left vector
                double lateralOffsetMagMm = ComputeLateralOffsetMm(lineRef, deckEdgeMm, halfPostWidthMm);

                // ===================== Process each segment (each side is independent) =====================
                int inserted = 0, failed = 0, connOk = 0, connFail = 0;
                int railSides = 0;

                // Collect each side rail definition first; build corner-aware rails after all sides are processed.
                var railSideSpecs = new List<RailSideSpec>();

                // Optional: prevents accidental duplicates if user picks extra colinear points
                var usedPostKeys = new HashSet<string>();

                for (int seg = 0; seg < sideCount; seg++)
                {
                    Point p1 = runPts[seg];
                    Point p2 = runPts[seg + 1];

                    Vector run = new Vector(p2.X - p1.X, p2.Y - p1.Y, p2.Z - p1.Z);
                    double runLen = Math.Sqrt(run.X * run.X + run.Y * run.Y + run.Z * run.Z);
                    if (runLen < 1.0) continue;

                    Vector dir = new Vector(run.X / runLen, run.Y / runLen, run.Z / runLen);
                    Vector left = GetLeftVectorXY(dir);
                    Position.RotationEnum postRotation = GetPostRotationFromRun(dir);

                    // IMPORTANT: apply start/end offsets to EVERY side
                    double segStartOffsetMm = startOffsetMmAll;
                    double segEndOffsetMm = endOffsetMmAll;

                    double usableLen = runLen - segStartOffsetMm - segEndOffsetMm;
                    if (usableLen <= 1.0) continue;

                    int stations = Math.Max(1, (int)Math.Ceiling(usableLen / spacingMm));
                    double actualSpacingMm = usableLen / stations;

                    Point firstStationOnLine = null;
                    Point lastStationOnLine = null;
                    double firstPostTopZ = 0.0;
                    double lastPostTopZ = 0.0;
                    Part firstHost = null;
                    Part lastHost = null;

                    // IMPORTANT: do NOT skip i=0 on later segments (no “sharing”)
                    for (int i = 0; i <= stations; i++)
                    {
                        double d = segStartOffsetMm + i * actualSpacingMm;

                        Point stationOnLine = new Point(
                            p1.X + dir.X * d,
                            p1.Y + dir.Y * d,
                            p1.Z + dir.Z * d);

                        Point station = new Point(
                            stationOnLine.X + left.X * lateralOffsetMagMm,
                            stationOnLine.Y + left.Y * lateralOffsetMagMm,
                            stationOnLine.Z);

                        Part nearestHost = null;
                        if (hostParts.Count > 0)
                            nearestHost = FindBestHostPartByXY(stationOnLine, hostParts);

                        double baseZ = stationOnLine.Z + baseOffsetMm;
                        if (nearestHost != null && TryGetPartTopZ(nearestHost, out double topZ))
                            baseZ = topZ + baseOffsetMm;

                        Point postStart = new Point(station.X, station.Y, baseZ);
                        Point postEnd = new Point(station.X, station.Y, baseZ + postHeightMm);

                        // Optional duplicate guard (helps if user adds extra points on the same straight line)
                        string key = PostKey(postStart, 1.0); // 1mm snap
                        if (usedPostKeys.Contains(key))
                            continue;
                        usedPostKeys.Add(key);

                        Beam postObj = CreatePost(postStart, postEnd, profile, material, postClass, postName, postRotation);
                        if (postObj == null)
                        {
                            failed++;
                            continue;
                        }

                        inserted++;
                        postObj.Modify();
                        CreatePostCap(postObj);

                        if (railEnabled && railCount > 0)
                        {
                            // Corner stations are handled by dedicated corner-seat logic in corner-aware rail build.
                            bool isCornerStation = (sideCount > 1) &&
                                ((seg > 0 && i == 0) || (seg < sideCount - 1 && i == stations));

                            if (!isCornerStation)
                            {
                                // Match the same rail-side logic used for rails
                                int seatSideSign = DetermineConnectionSideSign(left, stationOnLine, nearestHost);
                                if (seatSideSign == 0) seatSideSign = +1;

                                Vector dirXY = GetDirXYUnit(dir);

                                for (int r = 0; r < railCount; r++)
                                {
                                    // Rail centerline Z for this row at THIS post
                                    double railZ = postEnd.Z - railFromTopMm - (r * railSpacingMm);

                                    double halfRailDepthMm = InchesToMm(1.5) * 0.5;
                                    if (TryGetOutsideDimMm("TS1-1/2X1-1/2X.188", out var railOutsideMm))
                                        halfRailDepthMm = railOutsideMm * 0.5;

                                    CreateRailPostSeatAngle(
                                        stationOnLine,
                                        dirXY,
                                        left,
                                        lateralOffsetMagMm,
                                        halfPostWidthMm,
                                        seatSideSign,
                                        railZ,
                                        halfRailDepthMm,
                                        postObj,
                                        seatHoleLineIn,
                                        seatSlotC2CIn,
                                        seatSlotSizeIn,
                                        seatSlotStandard,
                                        seatSlotCutLenIn,
                                        seatSlotSpecial1,
                                        seatPilotC2CIn,
                                        seatPilotDiaIn,
                                        seatPilotStandard,
                                        seatPilotCutLenIn
                                    );
                                }
                            }
                        }

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
                            if (TryCreatePostConnection(nearestHost, postObj, connName, connAttr))
                                connOk++;
                            else
                                connFail++;
                        }
                    }

                    // Store this side for a corner-aware rail pass later.
                    if (railEnabled && railCount > 0 && firstStationOnLine != null && lastStationOnLine != null)
                    {
                        double halfRailWidthMm = InchesToMm(1.5) * 0.5;
                        if (TryGetOutsideDimMm("TS1-1/2X1-1/2X.188", out var railOutsideMmForSide))
                            halfRailWidthMm = railOutsideMmForSide * 0.5;

                        int railSideSign = DetermineConnectionSideSign(left, firstStationOnLine, (firstHost ?? lastHost));
                        if (railSideSign == 0) railSideSign = +1;

                        double railLateralMm = lateralOffsetMagMm + railSideSign * (halfPostWidthMm + halfRailWidthMm);

                        railSideSpecs.Add(new RailSideSpec
                        {
                            StartOnLine = firstStationOnLine,
                            EndOnLine = lastStationOnLine,
                            Dir = dir,
                            Left = left,
                            PostLineLateralMm = lateralOffsetMagMm,
                            HalfPostWidthMm = halfPostWidthMm,
                            RailLateralMm = railLateralMm,
                            FirstPostTopZ = firstPostTopZ,
                            LastPostTopZ = lastPostTopZ,
                            AnyHost = (firstHost ?? lastHost)
                        });

                        railSides++;
                    }
                }



                // Build rails in one pass so corners can trim/extend correctly and get non-butt end caps.
                if (railEnabled && railCount > 0 && railSideSpecs.Count > 0)
                {
                    bool isClosed = false; // closed loops are currently represented by repeating first point; we strip that above.
                    CreateCornerAwareRailsForPolyline(
                        railSideSpecs,
                        isClosed,
                        railStartOffsetMm,
                        railEndOffsetMm,
                        railFromTopMm,
                        railCount,
                        railSpacingMm,
                        seatHoleLineIn,
                        seatSlotC2CIn,
                        seatSlotSizeIn,
                        seatSlotStandard,
                        seatSlotCutLenIn,
                        seatSlotSpecial1
                    );
                }

                _model.CommitChanges();

                Operation.DisplayPrompt(
                    "WDRailing: sides=" + sideCount +
                    ", inserted=" + inserted +
                    ", failed=" + failed +
                    ", rail_sides=" + railSides +
                    ", conn_ok=" + connOk +
                    ", conn_fail=" + connFail);

                return true;
            }
            catch (Exception ex)
            {
                Operation.DisplayPrompt("WDRailing ERROR: " + ex.Message);
                return true;
            }
        }


        private static bool AreSamePoint(Point a, Point b, double tolMm)
        {
            if (a == null || b == null) return false;
            double dx = a.X - b.X;
            double dy = a.Y - b.Y;
            double dz = a.Z - b.Z;
            return (dx * dx + dy * dy + dz * dz) <= (tolMm * tolMm);
        }

        // ---------------- Connection ----------------


        private static string PostKey(Point p, double snapMm)
        {
            long xi = (long)Math.Round(p.X / snapMm);
            long yi = (long)Math.Round(p.Y / snapMm);
            long zi = (long)Math.Round(p.Z / snapMm);
            return xi + "|" + yi + "|" + zi;
        }



        private static double ParseDoubleOrThrow(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                throw new InvalidDataException("Stored coordinate value is blank.");

            double v;
            if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out v)) return v;
            if (double.TryParse(raw, NumberStyles.Float, CultureInfo.CurrentCulture, out v)) return v;

            throw new InvalidDataException("Invalid stored coordinate: " + raw);
        }

        private static double InchesToMm(double inches) { return inches * 25.4; }
        private static double MmToInches(double mm) { return mm / 25.4; }

        private static string ToInv(double d) => d.ToString("0.###", CultureInfo.InvariantCulture);


        private bool KeepWithPrompt(string msg)
        {
            Operation.DisplayPrompt("WDRailing: " + msg);
            return true;
        }
    }
}
