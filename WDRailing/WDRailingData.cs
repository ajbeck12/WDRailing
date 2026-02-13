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
    public class WDRailingData
    {
        [StructuresField("SPACING_IN")] public string SpacingIn;
        [StructuresField("POST_HEIGHT_IN")] public string PostHeightIn;
        [StructuresField("START_OFFSET_IN")] public string StartOffsetIn;
        [StructuresField("END_OFFSET_IN")] public string EndOffsetIn;
        [StructuresField("BASE_OFFSET_IN")] public string BaseOffsetIn;

        [StructuresField("LINE_REF")] public string LineRef;
        [StructuresField("DECK_EDGE_IN")] public string DeckEdgeIn;

        [StructuresField("POST_PROFILE")] public string PostProfile;
        [StructuresField("POST_MATERIAL")] public string PostMaterial;
        [StructuresField("POST_CLASS")] public string PostClass;
        [StructuresField("POST_NAME")] public string PostName;

        // Connection (post -> closest host)
        [StructuresField("CONN_ENABLED")] public string ConnEnabled; // "0" or "1"
        [StructuresField("CONN_NAME")] public string ConnName;       // component name or number
        [StructuresField("CONN_ATTR")] public string ConnAttr;       // optional attributes file

        // Persist points for Modify
        [StructuresField("P1X")] public string P1X;
        [StructuresField("P1Y")] public string P1Y;
        [StructuresField("P1Z")] public string P1Z;
        [StructuresField("P2X")] public string P2X;
        [StructuresField("P2Y")] public string P2Y;
        [StructuresField("P2Z")] public string P2Z;

        // Persist host IDs across Modify
        [StructuresField("HOST_IDS1")] public string HostIds1;
        [StructuresField("HOST_IDS2")] public string HostIds2;
        [StructuresField("HOST_IDS3")] public string HostIds3;
        [StructuresField("HOST_IDS4")] public string HostIds4;

        [StructuresField("RAIL_ENABLED")] public string RailEnabled;      // "0" or "1"
        [StructuresField("RAIL_START_OFF_IN")] public string RailStartOffsetIn;
        [StructuresField("RAIL_END_OFFSET_IN")] public string RailEndOffsetIn;
        [StructuresField("RAIL_FROM_TOP_IN")] public string RailFromTopIn; // distance DOWN from top of post to rail center
        [StructuresField("RAIL_COUNT")] public string RailCount;          // integer
        [StructuresField("RAIL_SPACING_IN")] public string RailSpacingIn; // c/c spacing between rail rows

        // Seat angle / post hole pattern (editable)
        [StructuresField("SEAT_HOLELINE_IN")] public string SeatHoleLineIn;         // from outside bend, down

        [StructuresField("SEAT_SLOT_C2C_IN")] public string SeatSlotC2CIn;          // slot center-to-center (slotted hole X)
        [StructuresField("SEAT_SLOT_SIZE_IN")] public string SeatSlotSizeIn;         // slot width / bolt size
        [StructuresField("SEAT_SLOT_STANDARD")] public string SeatSlotStandard;      // e.g. CARRIAGE
        [StructuresField("SEAT_SLOT_CUTLEN_IN")] public string SeatSlotCutLengthIn;  // e.g. 1
        [StructuresField("SEAT_SLOT_SPECIAL1")] public string SeatSlotSpecial1;      // "0" or "1" (special hole first layer)

        [StructuresField("SEAT_PILOT_C2C_IN")] public string SeatPilotC2CIn;         // pilot hole center-to-center (bolt dist X)
        [StructuresField("SEAT_PILOT_DIA_IN")] public string SeatPilotDiaIn;         // pilot hole diameter / bolt size
        [StructuresField("SEAT_PILOT_STANDARD")] public string SeatPilotStandard;    // e.g. TEK-3
        [StructuresField("SEAT_PILOT_CUTLEN_IN")] public string SeatPilotCutLengthIn;// e.g. 1

        // Persist polyline points across Modify (x|y|z;...)
        [StructuresField("RUNPTS1")] public string RunPts1;
        [StructuresField("RUNPTS2")] public string RunPts2;
        [StructuresField("RUNPTS3")] public string RunPts3;
        [StructuresField("RUNPTS4")] public string RunPts4;

    }
}
