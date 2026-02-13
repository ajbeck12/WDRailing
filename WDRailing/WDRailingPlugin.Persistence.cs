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
        private static bool TryReadInputPoints(List<InputDefinition> input, out Point p1, out Point p2, out object hostPayload)
        {
            p1 = null; p2 = null; hostPayload = null;
            if (input == null || input.Count < 2) return false;

            p1 = input[0].GetInput() as Point;
            p2 = input[1].GetInput() as Point;
            if (input.Count >= 3) hostPayload = input[2].GetInput();

            return p1 != null && p2 != null;
        }


        private void SavePointsToData(Point p1, Point p2)
        {
            _data.P1X = ToInv(p1.X); _data.P1Y = ToInv(p1.Y); _data.P1Z = ToInv(p1.Z);
            _data.P2X = ToInv(p2.X); _data.P2Y = ToInv(p2.Y); _data.P2Z = ToInv(p2.Z);
        }


        private string GetRunPtsCsvFromData()
        {
            return (_data.RunPts1 ?? "") + (_data.RunPts2 ?? "") + (_data.RunPts3 ?? "") + (_data.RunPts4 ?? "");
        }


        private void SetRunPtsCsvToData(string csv)
        {
            string[] chunks = SplitFixed(csv ?? "", 78, 4);
            _data.RunPts1 = chunks[0];
            _data.RunPts2 = chunks[1];
            _data.RunPts3 = chunks[2];
            _data.RunPts4 = chunks[3];
        }


        private static string SerializePoints(List<Point> pts)
        {
            if (pts == null || pts.Count == 0) return "";
            // x|y|z;...
            var parts = new List<string>(pts.Count);
            foreach (var p in pts)
                parts.Add(ToInv(p.X) + "|" + ToInv(p.Y) + "|" + ToInv(p.Z));
            return string.Join(";", parts);
        }


        private static List<Point> DeserializePoints(string csv)
        {
            var pts = new List<Point>();
            if (string.IsNullOrWhiteSpace(csv)) return pts;

            string[] items = csv.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var it in items)
            {
                var a = it.Split('|');
                if (a.Length != 3) continue;

                if (!double.TryParse(a[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double x)) continue;
                if (!double.TryParse(a[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double y)) continue;
                if (!double.TryParse(a[2], NumberStyles.Float, CultureInfo.InvariantCulture, out double z)) continue;

                pts.Add(new Point(x, y, z));
            }
            return pts;
        }


        private void SaveRunPointsToData(List<Point> pts)
        {
            SetRunPtsCsvToData(SerializePoints(pts));
            // Keep legacy P1/P2 as first/last so old logic still has something sane
            if (pts != null && pts.Count >= 2)
                SavePointsToData(pts[0], pts[pts.Count - 1]);
        }


        private List<Point> LoadRunPointsFromData()
        {
            return DeserializePoints(GetRunPtsCsvFromData());
        }


        private static bool IsPickerCancel(Exception ex)
        {
            string n = (ex.GetType().Name ?? "").ToLowerInvariant();
            if (n.Contains("cancel") || n.Contains("interrupt")) return true;
            string m = (ex.Message ?? "").ToLowerInvariant();
            return m.Contains("cancel") || m.Contains("interrupted");
        }


        private static double DistXY(Point a, Point b)
        {
            double dx = a.X - b.X;
            double dy = a.Y - b.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }



        private bool TryReadStoredPoints(out Point p1, out Point p2)
        {
            p1 = null; p2 = null;

            double p1x = ParseDoubleOrThrow(_data.P1X);
            double p1y = ParseDoubleOrThrow(_data.P1Y);
            double p1z = ParseDoubleOrThrow(_data.P1Z);
            double p2x = ParseDoubleOrThrow(_data.P2X);
            double p2y = ParseDoubleOrThrow(_data.P2Y);
            double p2z = ParseDoubleOrThrow(_data.P2Z);

            bool p1Zero = Math.Abs(p1x) < 0.0001 && Math.Abs(p1y) < 0.0001 && Math.Abs(p1z) < 0.0001;
            bool p2Zero = Math.Abs(p2x) < 0.0001 && Math.Abs(p2y) < 0.0001 && Math.Abs(p2z) < 0.0001;
            if (p1Zero && p2Zero) return false;

            p1 = new Point(p1x, p1y, p1z);
            p2 = new Point(p2x, p2y, p2z);

            double dx = p2.X - p1.X, dy = p2.Y - p1.Y, dz = p2.Z - p1.Z;
            double len = Math.Sqrt(dx * dx + dy * dy + dz * dz);

            return len > 0.5;
        }


        private static string JoinIds(List<int> ids)
        {
            if (ids == null || ids.Count == 0) return string.Empty;
            return string.Join(",", ids.ToArray());
        }


        private string GetHostCsvFromData()
        {
            return (_data.HostIds1 ?? "") + (_data.HostIds2 ?? "") + (_data.HostIds3 ?? "") + (_data.HostIds4 ?? "");
        }


        private void SetHostCsvToData(string csv)
        {
            string[] chunks = SplitFixed(csv ?? "", 78, 4);
            _data.HostIds1 = chunks[0];
            _data.HostIds2 = chunks[1];
            _data.HostIds3 = chunks[2];
            _data.HostIds4 = chunks[3];
        }


        private static string[] SplitFixed(string s, int chunkSize, int chunkCount)
        {
            var result = new string[chunkCount];
            for (int i = 0; i < chunkCount; i++) result[i] = "";

            if (string.IsNullOrEmpty(s)) return result;

            int idx = 0;
            for (int i = 0; i < chunkCount && idx < s.Length; i++)
            {
                int take = Math.Min(chunkSize, s.Length - idx);
                result[i] = s.Substring(idx, take);
                idx += take;
            }
            return result;
        }

        // ---------------- Host selection ----------------

    }
}
