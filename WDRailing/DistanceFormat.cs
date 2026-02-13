using System;
using System.Globalization;

namespace WDRailing
{
    internal static class DistanceFormat
    {
        /// <summary>
        /// Formats an inches value into Tekla-like feet/inches with fractions.
        /// Rules:
        ///  - |in| < 12: X" or X"n/d (no feet)
        ///  - |in| >= 12: F'-I" or F'-I"n/d
        /// Fractions are rounded to nearest 1/16 by default and reduced.
        /// </summary>
        public static string ToTeklaFeetInches(double inches, int denom = 16)
        {
            if (denom <= 0) denom = 16;

            bool neg = inches < 0;
            double a = Math.Abs(inches);

            // Round to nearest 1/denom inch
            double rounded = Math.Round(a * denom, MidpointRounding.AwayFromZero) / denom;

            int feet = 0;
            double inchRem = rounded;

            if (rounded >= 12.0 - 1e-9)
            {
                feet = (int)Math.Floor(rounded / 12.0);
                inchRem = rounded - feet * 12.0;
            }

            // Handle rounding that pushed inches to 12
            if (inchRem >= 12.0 - 1e-9)
            {
                feet += 1;
                inchRem = 0.0;
            }

            int inchWhole = (int)Math.Floor(inchRem + 1e-9);
            double frac = inchRem - inchWhole;

            int num = (int)Math.Round(frac * denom, MidpointRounding.AwayFromZero);
            int den = denom;

            // Carry if fraction rounds to 1
            if (num >= den)
            {
                inchWhole += 1;
                num = 0;
            }

            if (feet > 0 && inchWhole >= 12)
            {
                feet += 1;
                inchWhole = 0;
                num = 0;
            }

            // Reduce
            if (num != 0)
            {
                int g = Gcd(num, den);
                num /= g;
                den /= g;
            }

            string fracStr = (num == 0) ? "" : $"{num}/{den}";

            string core;
            if (feet == 0)
            {
                core = (num == 0) ? $"{inchWhole}\"" : $"{inchWhole}\"{fracStr}";
            }
            else
            {
                core = (num == 0) ? $"{feet}'-{inchWhole}\"" : $"{feet}'-{inchWhole}\"{fracStr}";
            }

            return neg ? "-" + core : core;
        }

        private static int Gcd(int a, int b)
        {
            a = Math.Abs(a);
            b = Math.Abs(b);
            while (b != 0)
            {
                int t = a % b;
                a = b;
                b = t;
            }
            return a == 0 ? 1 : a;
        }
    }
}
