using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace WDRailing
{
    /// <summary>
    /// Strict config loader.
    /// - No hardcoded numeric/profile defaults.
    /// - Required keys must exist and be non-blank.
    /// - Supports legacy key aliases (e.g., ConnEnabled -> CreateConnection).
    /// </summary>
    internal sealed class WDRailingDefaults
    {
        public const string FileName = "WDRailing.defaults.ini";

        public string SpacingIn { get; private set; }
        public string PostHeightIn { get; private set; }
        public string StartOffsetIn { get; private set; }
        public string EndOffsetIn { get; private set; }
        public string BaseOffsetIn { get; private set; }
        public string LineRef { get; private set; }
        public string DeckEdgeIn { get; private set; }

        public string PostProfile { get; private set; }
        public string PostMaterial { get; private set; }
        public string PostClass { get; private set; }
        public string PostName { get; private set; }

        // Connection defaults
        public string CreateConnection { get; private set; } // "0" or "1"
        public string ConnectionName { get; private set; }   // name or number
        public string ConnectionAttr { get; private set; }   // may be blank, key required

        // Horizontal Rail
        public string RailEnabled { get; private set; }
        public string RailStartOffsetIn { get; private set; }
        public string RailEndOffsetIn { get; private set; }
        public string RailFromTopIn { get; private set; }
        public string RailCount { get; private set; }
        public string RailSpacingIn { get; private set; }

        // Seat angle / post-hole defaults
        public string SeatHoleLineFromBendIn { get; private set; }

        public string SeatSlotC2CIn { get; private set; }
        public string SeatSlotSizeIn { get; private set; }
        public string SeatSlotStandard { get; private set; }
        public string SeatSlotCutLengthIn { get; private set; }
        public string SeatSlotSpecial1 { get; private set; } // "0" or "1"

        public string SeatPilotC2CIn { get; private set; }
        public string SeatPilotDiaIn { get; private set; }
        public string SeatPilotStandard { get; private set; }
        public string SeatPilotCutLengthIn { get; private set; }

        public static string GetConfigPath()
        {
            string dllPath = Assembly.GetExecutingAssembly().Location;
            string dir = Path.GetDirectoryName(dllPath) ?? "";
            return Path.Combine(dir, FileName);
        }

        public static WDRailingDefaults LoadOrThrow()
        {
            string path = GetConfigPath();
            if (!File.Exists(path))
                throw new FileNotFoundException("Missing required config file: " + path);

            Dictionary<string, string> kv = ReadIni(path);

            var d = new WDRailingDefaults
            {
                SpacingIn = Require(kv, "SpacingIn"),
                PostHeightIn = Require(kv, "PostHeightIn"),
                StartOffsetIn = Require(kv, "StartOffsetIn"),
                EndOffsetIn = Require(kv, "EndOffsetIn"),
                BaseOffsetIn = Require(kv, "BaseOffsetIn"),
                LineRef = Require(kv, "LineRef"),
                DeckEdgeIn = Require(kv, "DeckEdgeIn"),

                PostProfile = Require(kv, "PostProfile"),
                PostMaterial = Require(kv, "PostMaterial"),
                PostClass = Require(kv, "PostClass"),
                PostName = Require(kv, "PostName"),

                // Accept both new and legacy key names
                CreateConnection = RequireAny(kv, "CreateConnection", "ConnEnabled"),
                ConnectionName = RequireAny(kv, "ConnectionName", "ConnName"),
                ConnectionAttr = GetAllowBlankAny(kv, "ConnectionAttr", "ConnAttr"),

                RailEnabled = Require(kv, "RailEnabled"),
                RailStartOffsetIn = Require(kv, "RailStartOffsetIn"),
                RailEndOffsetIn = Require(kv, "RailEndOffsetIn"),
                RailFromTopIn = Require(kv, "RailFromTopIn"),
                RailCount = Require(kv, "RailCount"),
                RailSpacingIn = Require(kv, "RailSpacingIn"),

                SeatHoleLineFromBendIn = Require(kv, "SeatHoleLineFromBendIn"),

                SeatSlotC2CIn = Require(kv, "SeatSlotC2CIn"),
                SeatSlotSizeIn = Require(kv, "SeatSlotSizeIn"),
                SeatSlotStandard = Require(kv, "SeatSlotStandard"),
                SeatSlotCutLengthIn = Require(kv, "SeatSlotCutLengthIn"),
                SeatSlotSpecial1 = Require(kv, "SeatSlotSpecial1"),

                SeatPilotC2CIn = Require(kv, "SeatPilotC2CIn"),
                SeatPilotDiaIn = Require(kv, "SeatPilotDiaIn"),
                SeatPilotStandard = Require(kv, "SeatPilotStandard"),
                SeatPilotCutLengthIn = Require(kv, "SeatPilotCutLengthIn"),
            };

            string lr = d.LineRef.Trim().ToUpperInvariant();
            if (lr != "LEFT" && lr != "RIGHT" && lr != "MIDDLE")
                throw new InvalidDataException("Config LineRef must be LEFT, RIGHT, or MIDDLE. Got: " + d.LineRef);
            d.LineRef = lr;

            d.CreateConnection = Normalize01(d.CreateConnection, "CreateConnection");
            d.RailEnabled = Normalize01(d.RailEnabled, "RailEnabled");
            d.SeatSlotSpecial1 = Normalize01(d.SeatSlotSpecial1, "SeatSlotSpecial1");

            return d;
        }

        private static string Normalize01(string raw, string key)
        {
            string v = (raw ?? "").Trim();
            if (v == "0" || v == "1") return v;

            // Optional tolerance for common bool strings.
            string u = v.ToUpperInvariant();
            if (u == "TRUE" || u == "YES" || u == "ON") return "1";
            if (u == "FALSE" || u == "NO" || u == "OFF") return "0";

            throw new InvalidDataException("Config " + key + " must be 0 or 1. Got: " + raw);
        }

        private static Dictionary<string, string> ReadIni(string path)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string raw in File.ReadAllLines(path))
            {
                string line = (raw ?? "").Trim();
                if (line.Length == 0) continue;
                if (line.StartsWith("#") || line.StartsWith(";")) continue;

                int hash = line.IndexOf('#');
                if (hash >= 0) line = line.Substring(0, hash).Trim();

                int eq = line.IndexOf('=');
                if (eq <= 0) continue;

                string key = line.Substring(0, eq).Trim();
                string val = line.Substring(eq + 1).Trim();

                if (key.Length == 0) continue;
                dict[key] = val;
            }
            return dict;
        }

        private static string Require(Dictionary<string, string> kv, string key)
        {
            if (!kv.TryGetValue(key, out string v) || v == null)
                throw new InvalidDataException("Missing required config key: " + key);
            if (string.IsNullOrWhiteSpace(v))
                throw new InvalidDataException("Config key cannot be blank: " + key);
            return v.Trim();
        }

        private static string RequireAny(Dictionary<string, string> kv, params string[] keys)
        {
            foreach (string key in keys)
            {
                if (kv.TryGetValue(key, out string v) && v != null)
                {
                    if (string.IsNullOrWhiteSpace(v))
                        throw new InvalidDataException("Config key cannot be blank: " + key);
                    return v.Trim();
                }
            }
            throw new InvalidDataException("Missing required config key. Expected one of: " + string.Join(", ", keys));
        }

        private static string GetAllowBlankAny(Dictionary<string, string> kv, params string[] keys)
        {
            foreach (string key in keys)
            {
                if (kv.TryGetValue(key, out string v) && v != null)
                    return v.Trim(); // may be empty
            }
            throw new InvalidDataException("Missing required config key. Expected one of: " + string.Join(", ", keys));
        }

        private WDRailingDefaults() { }
    }
}
