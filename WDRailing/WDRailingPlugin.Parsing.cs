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


        private static double ParseImperialInchesOrThrow(string raw, bool allowNegative)
        {
            if (string.IsNullOrWhiteSpace(raw))
                throw new InvalidDataException("Distance string is blank.");

            string s = raw.Trim();
            s = s.Replace("”", "\"").Replace("“", "\"")
                 .Replace("′", "'").Replace("’", "'");
            s = s.ToLowerInvariant();
            s = s.Replace("inches", "").Replace("inch", "").Replace(" in", "").Trim();

            bool neg = false;
            if (s.StartsWith("+")) s = s.Substring(1).Trim();
            else if (s.StartsWith("-")) { neg = true; s = s.Substring(1).Trim(); }

            double inches;
            int apos = s.IndexOf('\'');
            if (apos >= 0)
            {
                string feetPart = s.Substring(0, apos).Trim();
                string inchPart = s.Substring(apos + 1).Trim().Replace("\"", "").Trim();
                if (inchPart.StartsWith("-")) inchPart = inchPart.Substring(1).Trim();

                double feet = 0.0;
                if (feetPart.Length > 0 && !TryParseDoubleLoose(feetPart, out feet))
                    throw new InvalidDataException("Invalid feet value: " + raw);

                double inchOnly = 0.0;
                if (inchPart.Length > 0 && !TryParseMixedNumber(inchPart, out inchOnly))
                    throw new InvalidDataException("Invalid inch value: " + raw);

                inches = feet * 12.0 + inchOnly;
            }
            else
            {
                s = s.Replace("\"", "").Trim();
                if (!TryParseMixedNumber(s, out inches))
                    throw new InvalidDataException("Invalid distance: " + raw);
            }

            if (neg) inches = -Math.Abs(inches);

            if (!allowNegative && inches <= 0.0)
                throw new InvalidDataException("Distance must be > 0: " + raw);

            return inches;
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
