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
        private static bool TryGetOutsideDimMm(string profile, out double outsideMm)
        {
            outsideMm = 0.0;
            if (string.IsNullOrWhiteSpace(profile)) return false;

            string s = profile.Trim().ToUpperInvariant().Replace(" ", "");

            if (s.StartsWith("HSS")) s = s.Substring(3);
            else if (s.StartsWith("TS")) s = s.Substring(2);
            else if (s.StartsWith("RHS")) s = s.Substring(3);
            else if (s.StartsWith("SHS")) s = s.Substring(3);
            else return false;

            string[] dims = s.Split(new[] { 'X' }, StringSplitOptions.RemoveEmptyEntries);
            if (dims.Length < 2) return false;

            double d1In, d2In;
            if (!TryParseProfileDimInches(dims[0], out d1In)) return false;
            if (!TryParseProfileDimInches(dims[1], out d2In)) return false;

            outsideMm = InchesToMm(Math.Max(d1In, d2In));
            return outsideMm > 0.001;
        }


        private static bool TryParseProfileDimInches(string token, out double inches)
        {
            inches = 0.0;
            if (string.IsNullOrWhiteSpace(token)) return false;

            Match m = Regex.Match(token, @"[0-9][0-9\.\-\/]*");
            if (!m.Success) return false;

            string t = m.Value.Replace("-", " "); // mixed number
            return TryParseMixedNumber(t, out inches);
        }

        // ---------------- Input / persistence ----------------

        private static bool[] ParseRailRowMask(string raw, int railCount)
        {
            var mask = new bool[Math.Max(0, railCount)];
            if (railCount <= 0) return mask;

            if (string.IsNullOrWhiteSpace(raw) || raw.Trim().Equals("ALL", StringComparison.OrdinalIgnoreCase))
            {
                for (int i = 0; i < railCount; i++) mask[i] = true;
                return mask;
            }

            if (raw.Trim().Equals("NONE", StringComparison.OrdinalIgnoreCase))
                return mask;

            string[] parts = raw.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts)
            {
                if (!int.TryParse(part.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int row))
                    throw new InvalidDataException("Invalid rail row in loop mask: " + part);

                if (row < 1 || row > railCount)
                    throw new InvalidDataException("Rail row out of range in loop mask: " + row);

                mask[row - 1] = true; // 1-based input, 0-based internal
            }

            return mask;
        }

        private static string NormalizeRailRowMask(bool[] mask)
        {
            if (mask == null || mask.Length == 0) return "NONE";

            var rows = new List<string>();
            for (int i = 0; i < mask.Length; i++)
                if (mask[i]) rows.Add((i + 1).ToString(CultureInfo.InvariantCulture));

            if (rows.Count == 0) return "NONE";
            if (rows.Count == mask.Length) return "ALL";

            return string.Join(",", rows);
        }

        private static double ParseImperialInchesOrThrow(string raw, bool allowNegative)
        {
            if (string.IsNullOrWhiteSpace(raw))
                throw new InvalidDataException("Distance value is blank.");

            string s = raw.Trim();

            s = s.Replace("’", "'")
                 .Replace("′", "'")
                 .Replace("“", "\"")
                 .Replace("”", "\"")
                 .Replace("″", "\"")
                 .Replace("–", "-")
                 .Replace("—", "-")
                 .Trim();

            bool negative = false;
            if (s.StartsWith("-", StringComparison.Ordinal))
            {
                negative = true;
                s = s.Substring(1).Trim();
            }

            if (negative && !allowNegative)
                throw new InvalidDataException("Negative value not allowed: " + raw);

            // Fix glued inch+fraction input:
            // 2"3/4    -> 2 3/4
            // 1'-6"3/4 -> 1'-6 3/4
            s = Regex.Replace(s, @"(?<=\d)""\s*(?=\d+\s*/\s*\d+\s*$)", " ");

            double feet = 0.0;
            double inches = 0.0;

            int ftIdx = s.IndexOf('\'');
            if (ftIdx >= 0)
            {
                string ftPart = s.Substring(0, ftIdx).Trim();
                string inchPart = s.Substring(ftIdx + 1).Trim();

                if (ftPart.Length > 0)
                {
                    if (!double.TryParse(ftPart, NumberStyles.Float, CultureInfo.InvariantCulture, out feet) &&
                        !double.TryParse(ftPart, NumberStyles.Float, CultureInfo.CurrentCulture, out feet))
                    {
                        throw new InvalidDataException("Invalid feet value: " + raw);
                    }
                }

                // IMPORTANT:
                // In feet-inch notation, the dash is usually a separator, not a negative sign.
                // So:
                //   1'-6"     => 1' + 6"
                //   1'-6 3/4" => 1' + 6 3/4"
                if (inchPart.StartsWith("-", StringComparison.Ordinal))
                    inchPart = inchPart.Substring(1).Trim();

                inches = ParseImperialInchPartOrThrow(inchPart, raw);
            }
            else
            {
                inches = ParseImperialInchPartOrThrow(s, raw);
            }

            double total = feet * 12.0 + inches;
            return negative ? -total : total;
        }

        private static double ParseImperialInchPartOrThrow(string raw, string originalRaw)
        {
            string s = (raw ?? "").Trim();
            if (s.Length == 0)
                return 0.0;

            s = s.Replace("\"", "").Trim();

            // 6-3/4 -> 6 3/4
            s = Regex.Replace(s, @"(?<=\d)-(?=\d+\s*/\s*\d+$)", " ");

            string[] parts = Regex.Split(s, @"\s+");
            var tokens = new List<string>();

            foreach (string p in parts)
            {
                string t = (p ?? "").Trim();
                if (t.Length > 0)
                    tokens.Add(t);
            }

            if (tokens.Count == 0)
                return 0.0;

            // Fraction only: 3/4
            if (tokens.Count == 1 && tokens[0].Contains("/"))
                return ParseFractionOrThrow(tokens[0], originalRaw);

            // Whole inches only: 6 or 6.25
            if (tokens.Count == 1)
            {
                if (double.TryParse(tokens[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double wholeOnly))
                    return wholeOnly;

                if (double.TryParse(tokens[0], NumberStyles.Float, CultureInfo.CurrentCulture, out wholeOnly))
                    return wholeOnly;

                throw new InvalidDataException("Invalid inch value: " + originalRaw);
            }

            // Whole + fraction: 6 3/4
            if (tokens.Count == 2)
            {
                if (!double.TryParse(tokens[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double whole) &&
                    !double.TryParse(tokens[0], NumberStyles.Float, CultureInfo.CurrentCulture, out whole))
                {
                    throw new InvalidDataException("Invalid whole-inch value: " + originalRaw);
                }

                double frac = ParseFractionOrThrow(tokens[1], originalRaw);
                return whole + frac;
            }

            throw new InvalidDataException("Invalid imperial distance: " + originalRaw);
        }

        private static double ParseFractionOrThrow(string raw, string originalRaw)
        {
            Match m = Regex.Match((raw ?? "").Trim(), @"^(\d+)\s*/\s*(\d+)$");
            if (!m.Success)
                throw new InvalidDataException("Invalid fraction: " + originalRaw);

            if (!double.TryParse(m.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double num) &&
                !double.TryParse(m.Groups[1].Value, NumberStyles.Float, CultureInfo.CurrentCulture, out num))
            {
                throw new InvalidDataException("Invalid fraction numerator: " + originalRaw);
            }

            if (!double.TryParse(m.Groups[2].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double den) &&
                !double.TryParse(m.Groups[2].Value, NumberStyles.Float, CultureInfo.CurrentCulture, out den))
            {
                throw new InvalidDataException("Invalid fraction denominator: " + originalRaw);
            }

            if (Math.Abs(den) < 0.0000001)
                throw new InvalidDataException("Fraction denominator cannot be zero: " + originalRaw);

            return num / den;
        }
        private static bool TryParseMixedNumber(string s, out double value)
        {
            value = 0.0;
            if (string.IsNullOrWhiteSpace(s)) return false;

            string t = s.Trim().Replace("-", " "); // allows 11-1/2
            string[] parts = t.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return false;

            if (parts.Length == 1)
            {
                if (parts[0].Contains("/"))
                    return TryParseFraction(parts[0], out value);

                return TryParseDoubleLoose(parts[0], out value);
            }

            double whole;
            if (!TryParseDoubleLoose(parts[0], out whole)) return false;

            double tail;
            if (parts[1].Contains("/"))
            {
                if (!TryParseFraction(parts[1], out tail)) return false;
            }
            else
            {
                if (!TryParseDoubleLoose(parts[1], out tail)) return false;
            }

            value = whole + tail;
            return true;
        }


        private static bool TryParseFraction(string token, out double value)
        {
            value = 0.0;
            if (string.IsNullOrWhiteSpace(token)) return false;

            string[] nd = token.Split('/');
            if (nd.Length != 2) return false;

            double n, d;
            if (!TryParseDoubleLoose(nd[0], out n)) return false;
            if (!TryParseDoubleLoose(nd[1], out d)) return false;
            if (Math.Abs(d) < 1e-9) return false;

            value = n / d;
            return true;
        }


        private static bool TryParseDoubleLoose(string s, out double v)
        {
            string t = (s ?? "").Trim();

            if (double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out v))
                return true;

            if (double.TryParse(t, NumberStyles.Float, CultureInfo.CurrentCulture, out v))
                return true;

            t = t.Replace(",", "");
            if (double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out v))
                return true;

            return false;
        }

    }
}
