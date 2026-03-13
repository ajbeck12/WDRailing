using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
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

            if (string.IsNullOrWhiteSpace(_data.StartLoopEnabled)) _data.StartLoopEnabled = cfg.StartLoopEnabled;
            if (string.IsNullOrWhiteSpace(_data.EndLoopEnabled)) _data.EndLoopEnabled = cfg.EndLoopEnabled;
            if (_data.ConnFlipPosts == null) _data.ConnFlipPosts = cfg.ConnFlipPosts ?? "";

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
                string spacingModeRaw = (string.IsNullOrWhiteSpace(_data.SpacingMode) ? cfg.SpacingMode : _data.SpacingMode).Trim().ToUpperInvariant();
                if (spacingModeRaw != "AUTOMATIC" && spacingModeRaw != "EXACT START" && spacingModeRaw != "EXACT END" && spacingModeRaw != "MAX")
                    throw new InvalidDataException("SPACING_MODE must be AUTOMATIC, EXACT START, EXACT END, or MAX. Got: " + spacingModeRaw);

                string spacingRaw = (string.IsNullOrWhiteSpace(_data.SpacingIn) ? cfg.SpacingIn : _data.SpacingIn).Trim();

                double spacingIn = 0.0;
                List<List<double>> spacingPatternBySegmentIn = null;

                if (spacingModeRaw == "AUTOMATIC" || spacingModeRaw == "MAX")
                {
                    spacingIn = ParseImperialInchesOrThrow(spacingRaw, allowNegative: false);
                }
                else
                {
                    spacingPatternBySegmentIn = ParseSegmentedSpacingPatternOrThrow(spacingRaw, allowNegative: false);
                }

                _data.SpacingMode = spacingModeRaw;
                _data.SpacingIn = spacingRaw;

                double heightIn = ParseImperialInchesOrThrow(string.IsNullOrWhiteSpace(_data.PostHeightIn) ? cfg.PostHeightIn : _data.PostHeightIn, allowNegative: false);
                List<double> startOffsetsIn = ParseImperialDistanceListOrThrow(
                    string.IsNullOrWhiteSpace(_data.StartOffsetIn) ? cfg.StartOffsetIn : _data.StartOffsetIn,
                    allowNegative: true);

                List<double> endOffsetsIn = ParseImperialDistanceListOrThrow(
                    string.IsNullOrWhiteSpace(_data.EndOffsetIn) ? cfg.EndOffsetIn : _data.EndOffsetIn,
                    allowNegative: true);
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

                string connFlipPostsRaw = (_data.ConnFlipPosts ?? cfg.ConnFlipPosts ?? "").Trim();
                HashSet<int> connFlipPosts = ParsePositiveIndexSet(connFlipPostsRaw);
                _data.ConnFlipPosts = NormalizePositiveIndexSet(connFlipPosts);

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

                string startLoopEnabledRaw = (string.IsNullOrWhiteSpace(_data.StartLoopEnabled) ? cfg.StartLoopEnabled : _data.StartLoopEnabled).Trim();
                if (startLoopEnabledRaw != "0" && startLoopEnabledRaw != "1")
                    throw new InvalidDataException("START_LOOP_ENABLED must be 0 or 1. Got: " + startLoopEnabledRaw);
                bool startLoopEnabled = (startLoopEnabledRaw == "1");

                string endLoopEnabledRaw = (string.IsNullOrWhiteSpace(_data.EndLoopEnabled) ? cfg.EndLoopEnabled : _data.EndLoopEnabled).Trim();
                if (endLoopEnabledRaw != "0" && endLoopEnabledRaw != "1")
                    throw new InvalidDataException("END_LOOP_ENABLED must be 0 or 1. Got: " + endLoopEnabledRaw);
                bool endLoopEnabled = (endLoopEnabledRaw == "1");

                _data.StartLoopEnabled = startLoopEnabledRaw;
                _data.EndLoopEnabled = endLoopEnabledRaw;

                bool[] startLoopMask = BuildAllRowsMask(railCount, startLoopEnabled);
                bool[] endLoopMask = BuildAllRowsMask(railCount, endLoopEnabled);

                string startPostEnabledRaw = (string.IsNullOrWhiteSpace(_data.StartPostEnabled) ? cfg.StartPostEnabled : _data.StartPostEnabled).Trim();
                if (startPostEnabledRaw != "0" && startPostEnabledRaw != "1")
                    throw new InvalidDataException("START_POST_ENABLED must be 0 or 1. Got: " + startPostEnabledRaw);
                bool startPostEnabled = (startPostEnabledRaw == "1");

                string endPostEnabledRaw = (string.IsNullOrWhiteSpace(_data.EndPostEnabled) ? cfg.EndPostEnabled : _data.EndPostEnabled).Trim();
                if (endPostEnabledRaw != "0" && endPostEnabledRaw != "1")
                    throw new InvalidDataException("END_POST_ENABLED must be 0 or 1. Got: " + endPostEnabledRaw);
                bool endPostEnabled = (endPostEnabledRaw == "1");

                _data.StartPostEnabled = startPostEnabledRaw;
                _data.EndPostEnabled = endPostEnabledRaw;

                // Persist effective values
                if (spacingModeRaw == "AUTOMATIC" || spacingModeRaw == "MAX")
                    _data.SpacingIn = DistanceFormat.ToTeklaFeetInches(spacingIn, 16);
                else
                    _data.SpacingIn = spacingRaw;
                _data.PostHeightIn = DistanceFormat.ToTeklaFeetInches(heightIn, 16);
                _data.StartOffsetIn = NormalizeImperialDistanceList(startOffsetsIn);
                _data.EndOffsetIn = NormalizeImperialDistanceList(endOffsetsIn);
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
                double lateralOffsetBaseMm = Math.Abs(ComputeLateralOffsetMm(lineRef, deckEdgeMm, halfPostWidthMm));
                bool hasDeckEdge = Math.Abs(deckEdgeMm) > 0.0001;

                // ===================== Process each segment (each side is independent) =====================
                int inserted = 0, failed = 0, connOk = 0, connFail = 0;
                int railSides = 0;

                // Collect each side rail definition first; build corner-aware rails after all sides are processed.
                var railSideSpecs = new List<RailSideSpec>();

                // Optional: prevents accidental duplicates if user picks extra colinear points
                var usedPostKeys = new HashSet<string>();

                int globalPostNumber = 0;

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
                    double segStartOffsetMm = InchesToMm(GetDistanceForSegment(startOffsetsIn, seg));
                    double segEndOffsetMm = InchesToMm(GetDistanceForSegment(endOffsetsIn, seg));

                    // Clamp the effective start/end distances to the actual segment length
                    double startD = Math.Max(0.0, Math.Min(runLen, segStartOffsetMm));
                    double endD = Math.Max(0.0, Math.Min(runLen, runLen - segEndOffsetMm));
                    double usableLen = endD - startD;

                    Point railStartRefOnLine = new Point(
                        p1.X + dir.X * startD,
                        p1.Y + dir.Y * startD,
                        p1.Z + dir.Z * startD);

                    Point railEndRefOnLine = new Point(
                        p1.X + dir.X * endD,
                        p1.Y + dir.Y * endD,
                        p1.Z + dir.Z * endD);

                    // Resolve post side per segment using a midpoint on the effective segment span.
                    double dMid = (startD + endD) * 0.5;
                    Point segMidOnLine = new Point(
                        p1.X + dir.X * dMid,
                        p1.Y + dir.Y * dMid,
                        p1.Z + dir.Z * dMid);

                    Part segHostForSide = null;
                    if (hostParts.Count > 0)
                        segHostForSide = FindBestHostPartByXY(segMidOnLine, hostParts);

                    int segPostSideSign = ResolvePostSideSign(lineRef, left, segMidOnLine, segHostForSide, hasDeckEdge);
                    double segPostLateralMm = (segPostSideSign == 0) ? 0.0 : (segPostSideSign * lateralOffsetBaseMm);

                    List<double> stationOffsetsMm;

                    double railStartTopZ = ResolvePostTopZAtStation(railStartRefOnLine, hostParts, baseOffsetMm, postHeightMm);
                    double railEndTopZ = ResolvePostTopZAtStation(railEndRefOnLine, hostParts, baseOffsetMm, postHeightMm);

                    Part railStartHost = null;
                    if (hostParts.Count > 0)
                        railStartHost = FindBestHostPartByXY(railStartRefOnLine, hostParts);

                    Part railEndHost = null;
                    if (hostParts.Count > 0)
                        railEndHost = FindBestHostPartByXY(railEndRefOnLine, hostParts);

                    // If the segment is too short for both offsets, still place one post.
                    if (usableLen <= 1.0)
                    {
                        stationOffsetsMm = new List<double>();

                        if (startPostEnabled)
                            stationOffsetsMm.Add(startD);
                        else if (endPostEnabled)
                            stationOffsetsMm.Add(endD);
                    }
                    else
                    {
                        List<double> segSpacingPatternIn = null;
                        if (spacingModeRaw == "EXACT START" || spacingModeRaw == "EXACT END")
                            segSpacingPatternIn = GetSpacingPatternForSegment(spacingPatternBySegmentIn, seg);

                        stationOffsetsMm = BuildStationOffsetsMm(
                            usableLen,
                            spacingModeRaw,
                            spacingMm,
                            segSpacingPatternIn);

                        // Shift from 0..usableLen to actual distance along the segment
                        for (int i = 0; i < stationOffsetsMm.Count; i++)
                            stationOffsetsMm[i] = startD + stationOffsetsMm[i];

                        if (!startPostEnabled && stationOffsetsMm.Count > 0)
                            stationOffsetsMm.RemoveAt(0);

                        if (!endPostEnabled && stationOffsetsMm.Count > 0)
                            stationOffsetsMm.RemoveAt(stationOffsetsMm.Count - 1);
                    }

                    if (stationOffsetsMm.Count == 0)
                        continue;

                    Point firstStationOnLine = null;
                    Point lastStationOnLine = null;
                    double firstPostTopZ = 0.0;
                    double lastPostTopZ = 0.0;
                    Part firstHost = null;
                    Part lastHost = null;

                    // IMPORTANT: do NOT skip i=0 on later segments (no “sharing”)
                    for (int i = 0; i < stationOffsetsMm.Count; i++)
                    {
                        double d = stationOffsetsMm[i];

                        Point stationOnLine = new Point(
                            p1.X + dir.X * d,
                            p1.Y + dir.Y * d,
                            p1.Z + dir.Z * d);

                        Point station = new Point(
                            stationOnLine.X + left.X * segPostLateralMm,
                            stationOnLine.Y + left.Y * segPostLateralMm,
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
                        globalPostNumber++;
                        postObj.Modify();
                        CreatePostCap(postObj);

                        bool flipThisPostConnection = connFlipPosts.Contains(globalPostNumber);
                        string csValue = flipThisPostConnection ? "L" : "R";

                        if (railEnabled && railCount > 0)
                        {
                            // Regular post seat angles should be independent from corner-seat logic.
                            // Keep creating these at every post (including posts next to corners).
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
                                    segPostLateralMm,
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
                            if (TryCreatePostConnection(nearestHost, postObj, connName, connAttr, csValue))
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

                        double railLateralMm = segPostLateralMm + railSideSign * (halfPostWidthMm + halfRailWidthMm);

                        railSideSpecs.Add(new RailSideSpec
                        {
                            StartOnLine = railStartRefOnLine,
                            EndOnLine = railEndRefOnLine,
                            Dir = dir,
                            Left = left,
                            PostLineLateralMm = segPostLateralMm,
                            HalfPostWidthMm = halfPostWidthMm,
                            RailLateralMm = railLateralMm,
                            FirstPostTopZ = railStartTopZ,
                            LastPostTopZ = railEndTopZ,
                            AnyHost = (railStartHost ?? railEndHost ?? firstHost ?? lastHost)
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
                        seatSlotSpecial1,
                        startLoopEnabled,
                        startLoopMask,
                        endLoopEnabled,
                        endLoopMask
                    );
                }

                if (railEnabled && railCount > 0 && railSideSpecs.Count > 0)
                {
                    CreateEndLoopsForPolyline(
                        railSideSpecs,
                        railCount,
                        railStartOffsetMm,
                        railEndOffsetMm,
                        railFromTopMm,
                        railSpacingMm,
                        startLoopEnabled,
                        startLoopMask,
                        endLoopEnabled,
                        endLoopMask
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

        private static bool[] BuildAllRowsMask(int railCount, bool enabled)
        {
            var mask = new bool[Math.Max(0, railCount)];

            if (!enabled)
                return mask;

            for (int i = 0; i < mask.Length; i++)
                mask[i] = true;

            return mask;
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

        private static List<double> BuildStationOffsetsMm(
    double usableLenMm,
    string spacingMode,
    double targetSpacingMm,
    List<double> patternIn)
        {
            var offsets = new List<double>();

            if (usableLenMm <= 1.0)
            {
                offsets.Add(0.0);
                return offsets;
            }

            spacingMode = (spacingMode ?? "AUTOMATIC").Trim().ToUpperInvariant();

            if (spacingMode == "AUTOMATIC")
            {
                if (targetSpacingMm <= 0.0)
                    throw new InvalidDataException("Automatic spacing must be > 0.");

                int stations = Math.Max(1, (int)Math.Ceiling(usableLenMm / targetSpacingMm));
                double actualSpacingMm = usableLenMm / stations;

                for (int i = 0; i <= stations; i++)
                    offsets.Add(i * actualSpacingMm);

                return offsets;
            }

            if (spacingMode == "MAX")
            {
                if (targetSpacingMm <= 0.0)
                    throw new InvalidDataException("Max spacing must be > 0.");

                offsets.Add(0.0);

                double d = 0.0;
                while (d + targetSpacingMm < usableLenMm - 0.5)
                {
                    d += targetSpacingMm;
                    offsets.Add(d);
                }

                if (Math.Abs(offsets[offsets.Count - 1] - usableLenMm) > 0.5)
                    offsets.Add(usableLenMm);

                return offsets;
            }

            if (patternIn == null || patternIn.Count == 0)
                throw new InvalidDataException("Exact spacing mode requires a spacing pattern.");

            var patternMm = new List<double>();
            foreach (double v in patternIn)
                patternMm.Add(InchesToMm(v));

            if (spacingMode == "EXACT START")
            {
                offsets.Add(0.0);

                double d = 0.0;
                for (int i = 0; i < patternMm.Count; i++)
                {
                    double next = d + patternMm[i];
                    if (next >= usableLenMm - 0.5)
                        break;

                    offsets.Add(next);
                    d = next;
                }

                if (Math.Abs(offsets[offsets.Count - 1] - usableLenMm) > 0.5)
                    offsets.Add(usableLenMm);

                return offsets;
            }

            if (spacingMode == "EXACT END")
            {
                var rev = new List<double>();
                rev.Add(usableLenMm);

                double d = usableLenMm;
                for (int i = 0; i < patternMm.Count; i++)
                {
                    double next = d - patternMm[i];
                    if (next <= 0.5)
                        break;

                    rev.Add(next);
                    d = next;
                }

                rev.Add(0.0);
                rev.Sort();

                double last = double.MinValue;
                foreach (double v in rev)
                {
                    if (offsets.Count == 0 || Math.Abs(v - last) > 0.5)
                    {
                        offsets.Add(v);
                        last = v;
                    }
                }

                return offsets;
            }

            throw new InvalidDataException("Unsupported spacing mode: " + spacingMode);
        }

        private static double ResolvePostTopZAtStation(Point stationOnLine, List<Part> hostParts, double baseOffsetMm, double postHeightMm)
        {
            Part nearestHost = null;
            if (hostParts != null && hostParts.Count > 0)
                nearestHost = FindBestHostPartByXY(stationOnLine, hostParts);

            double baseZ = stationOnLine.Z + baseOffsetMm;
            if (nearestHost != null && TryGetPartTopZ(nearestHost, out double topZ))
                baseZ = topZ + baseOffsetMm;

            return baseZ + postHeightMm;
        }

        private void CreateEndLoopsForPolyline(
    List<RailSideSpec> railSideSpecs,
    int railCount,
    double railStartOffsetMm,
    double railEndOffsetMm,
    double railFromTopMm,
    double railSpacingMm,
    bool startLoopEnabled,
    bool[] startLoopMask,
    bool endLoopEnabled,
    bool[] endLoopMask)
        {
            if (railSideSpecs == null || railSideSpecs.Count == 0 || railCount <= 0)
                return;

            if (startLoopEnabled)
            {
                CreateSingleTerminalEndLoop(
                    railSideSpecs[0],
                    atStart: true,
                    railCount: railCount,
                    railTerminalOffsetMm: railStartOffsetMm,
                    railFromTopMm: railFromTopMm,
                    railSpacingMm: railSpacingMm,
                    rowMask: startLoopMask);
            }

            if (endLoopEnabled)
            {
                CreateSingleTerminalEndLoop(
                    railSideSpecs[railSideSpecs.Count - 1],
                    atStart: false,
                    railCount: railCount,
                    railTerminalOffsetMm: railEndOffsetMm,
                    railFromTopMm: railFromTopMm,
                    railSpacingMm: railSpacingMm,
                    rowMask: endLoopMask);
            }
        }

        private void CreateSingleTerminalEndLoop(
    RailSideSpec side,
    bool atStart,
    int railCount,
    double railTerminalOffsetMm,
    double railFromTopMm,
    double railSpacingMm,
    bool[] rowMask)
        {
            if (side == null || rowMask == null || railCount <= 0)
                return;

            if (!TryGetSelectedRowRange(rowMask, railCount, out int topRow, out int bottomRow))
                return;

            if (topRow == bottomRow)
                return;

            double railOutsideMm = InchesToMm(1.5);
            if (TryGetOutsideDimMm("TS1-1/2X1-1/2X.188", out var outsideMm))
                railOutsideMm = outsideMm;

            double halfRailMm = railOutsideMm * 0.5;

            Vector dirXY = GetDirXYUnit(side.Dir);
            Point onLine = atStart ? side.StartOnLine : side.EndOnLine;
            double topOfPostZ = atStart ? side.FirstPostTopZ : side.LastPostTopZ;

            // Rail terminal plane:
            //   start = StartOnLine - Dir * railStartOffset
            //   end   = EndOnLine   + Dir * railEndOffset
            //
            // We want the OUTSIDE FACE of the vertical loop on that plane,
            // so move the loop centerline inward by half the rail OD.
            double along = atStart
                ? (-railTerminalOffsetMm + halfRailMm)
                : (railTerminalOffsetMm - halfRailMm);

            double x = onLine.X + side.Left.X * side.RailLateralMm + dirXY.X * along;
            double y = onLine.Y + side.Left.Y * side.RailLateralMm + dirXY.Y * along;

            double zTopRow = topOfPostZ - railFromTopMm - (topRow * railSpacingMm);
            double zBottomRow = topOfPostZ - railFromTopMm - (bottomRow * railSpacingMm);

            Point p1 = new Point(x, y, zTopRow - halfRailMm);
            Point p2 = new Point(x, y, zBottomRow + halfRailMm);

            Beam closure = CreateStraightRailMember(p1, p2);
            if (closure == null)
                return;

            closure.SetUserProperty("WD_END_LOOP", 1);
            closure.SetUserProperty("WD_END_LOOP_END", atStart ? "START" : "END");
            closure.SetUserProperty(
                "WD_END_LOOP_ROWS",
                (topRow + 1).ToString(CultureInfo.InvariantCulture) + "-" +
                (bottomRow + 1).ToString(CultureInfo.InvariantCulture));

            closure.Modify();
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

        private static List<double> ParseImperialDistanceListOrThrow(string raw, bool allowNegative)
        {
            if (string.IsNullOrWhiteSpace(raw))
                throw new InvalidDataException("Distance list is blank.");

            string[] parts;

            // If user uses commas/semicolons, prefer those.
            if (raw.IndexOf(',') >= 0 || raw.IndexOf(';') >= 0)
            {
                parts = raw.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
            }
            else
            {
                // Space-separated list, e.g. 1'-6" 1'-0"
                parts = Regex.Split(raw.Trim(), @"\s+");
            }

            var values = new List<double>();
            foreach (string part in parts)
            {
                string s = (part ?? "").Trim();
                if (s.Length == 0) continue;

                values.Add(ParseImperialInchesOrThrow(s, allowNegative));
            }

            if (values.Count == 0)
                throw new InvalidDataException("No valid distances were found in: " + raw);

            return values;
        }

        private static double GetDistanceForSegment(IReadOnlyList<double> values, int segmentIndex)
        {
            if (values == null || values.Count == 0)
                return 0.0;

            if (segmentIndex < values.Count)
                return values[segmentIndex];

            return values[values.Count - 1];
        }

        private static string NormalizeImperialDistanceList(IEnumerable<double> values)
        {
            var parts = new List<string>();

            foreach (double v in values)
            {
                if (v < 0)
                    parts.Add("-" + DistanceFormat.ToTeklaFeetInches(Math.Abs(v), 16));
                else
                    parts.Add(DistanceFormat.ToTeklaFeetInches(v, 16));
            }

            return string.Join(" ", parts);
        }
        private static HashSet<int> ParsePositiveIndexSet(string raw)
        {
            var set = new HashSet<int>();
            if (string.IsNullOrWhiteSpace(raw))
                return set;

            foreach (Match m in Regex.Matches(raw, @"\d+"))
            {
                if (int.TryParse(m.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int n) && n > 0)
                    set.Add(n);
            }

            return set;
        }

        private static List<List<double>> ParseSegmentedSpacingPatternOrThrow(string raw, bool allowNegative)
        {
            if (string.IsNullOrWhiteSpace(raw))
                throw new InvalidDataException("Spacing pattern is blank.");

            string[] segmentGroups = raw.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            var result = new List<List<double>>();

            foreach (string groupRaw in segmentGroups)
            {
                string group = (groupRaw ?? "").Trim();
                if (group.Length == 0) continue;

                result.Add(ParseSpacingPatternGroupOrThrow(group, allowNegative));
            }

            if (result.Count == 0)
                throw new InvalidDataException("No valid spacing segment groups were found in: " + raw);

            return result;
        }

        private static List<double> ParseSpacingPatternGroupOrThrow(string raw, bool allowNegative)
        {
            if (string.IsNullOrWhiteSpace(raw))
                throw new InvalidDataException("Spacing pattern group is blank.");

            string[] tokens;

            // Prefer commas inside one segment pattern.
            if (raw.IndexOf(',') >= 0)
                tokens = raw.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            else
                tokens = Regex.Split(raw.Trim(), @"\s+");

            var values = new List<double>();

            foreach (string tokenRaw in tokens)
            {
                string token = (tokenRaw ?? "").Trim();
                if (token.Length == 0) continue;

                int repeat = 1;
                string distRaw = token;

                Match m = Regex.Match(token, @"^\s*(\d+)\s*\*\s*(.+?)\s*$");
                if (m.Success)
                {
                    if (!int.TryParse(m.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out repeat) || repeat < 1)
                        throw new InvalidDataException("Invalid spacing repeat token: " + token);

                    distRaw = m.Groups[2].Value.Trim();
                }

                double distIn = ParseImperialInchesOrThrow(distRaw, allowNegative);
                if (distIn <= 0.0)
                    throw new InvalidDataException("Spacing values must be > 0. Got: " + distRaw);

                for (int i = 0; i < repeat; i++)
                    values.Add(distIn);
            }

            if (values.Count == 0)
                throw new InvalidDataException("No valid spacing values found in group: " + raw);

            return values;
        }

        private static List<double> GetSpacingPatternForSegment(List<List<double>> patternsBySegment, int segmentIndex)
        {
            if (patternsBySegment == null || patternsBySegment.Count == 0)
                throw new InvalidDataException("No exact spacing pattern groups were provided.");

            if (segmentIndex < patternsBySegment.Count)
                return patternsBySegment[segmentIndex];

            return patternsBySegment[patternsBySegment.Count - 1];
        }

        private static string NormalizePositiveIndexSet(HashSet<int> set)
        {
            if (set == null || set.Count == 0)
                return "";

            var vals = new List<int>(set);
            vals.Sort();
            return string.Join(" ", vals);
        }

        private static double InchesToMm(double inches) { return inches * 25.4; }
        private static double MmToInches(double mm) { return mm / 25.4; }

        private static string ToInv(double d) => d.ToString("0.###", CultureInfo.InvariantCulture);

        private Beam CreateStraightRailMember(Point start, Point end)
        {
            var rail = new Beam(start, end);

            rail.Profile.ProfileString = "TS1-1/2X1-1/2X.188";
            rail.Material.MaterialString = "A50";
            rail.Class = "1";
            rail.Name = "RAIL END LOOP";

            rail.Position.Plane = Position.PlaneEnum.MIDDLE;
            rail.Position.Depth = Position.DepthEnum.MIDDLE;
            rail.Position.Rotation = Position.RotationEnum.TOP;

            if (!rail.Insert())
                return null;

            return rail;
        }

        private bool KeepWithPrompt(string msg)
        {
            Operation.DisplayPrompt("WDRailing: " + msg);
            return true;
        }
    }
}
