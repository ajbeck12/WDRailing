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
        private List<Part> ResolveHostPartsFromPayload(object payload)
        {
            var parts = new List<Part>();
            if (payload == null) return parts;

            ArrayList list = payload as ArrayList;
            if (list == null) return parts;

            foreach (object o in list)
            {
                Identifier id = o as Identifier;
                if (id == null) continue;

                Part p = _model.SelectModelObject(id) as Part;
                if (p != null) parts.Add(p);
            }
            return parts;
        }


        private List<Part> ResolveHostPartsFromCsv(string csv)
        {
            var parts = new List<Part>();
            if (string.IsNullOrWhiteSpace(csv)) return parts;

            string[] tokens = csv.Split(',');
            foreach (string token in tokens)
            {
                string t = token.Trim();
                if (t.Length == 0) continue;

                int idInt;
                if (!int.TryParse(t, out idInt)) continue;

                Part p = _model.SelectModelObject(new Identifier(idInt)) as Part;
                if (p != null) parts.Add(p);
            }
            return parts;
        }


        private static bool TryGetPartTopZ(Part part, out double topZ)
        {
            topZ = 0.0;
            try
            {
                Solid s = part.GetSolid();
                topZ = s.MaximumPoint.Z;
                return true;
            }
            catch
            {
                return false;
            }
        }


        private static Part FindBestHostPartByXY(Point pt, List<Part> parts)
        {
            if (pt == null || parts == null || parts.Count == 0) return null;

            Part best = null;
            double bestScore = double.MaxValue;

            foreach (Part p in parts)
            {
                if (p == null) continue;

                try
                {
                    Solid s = p.GetSolid();
                    if (s == null) continue;

                    Point min = s.MinimumPoint;
                    Point max = s.MaximumPoint;

                    // Keep a mild inside-XY preference so overlapping members still win,
                    // but primarily choose by distance to MEMBER CENTERLINE in XY.
                    bool insideXY =
                        pt.X >= min.X && pt.X <= max.X &&
                        pt.Y >= min.Y && pt.Y <= max.Y;

                    double dist;
                    if (TryGetClosestPointOnPartCenterlineXY(p, pt, out Point cpt))
                    {
                        double dx = pt.X - cpt.X;
                        double dy = pt.Y - cpt.Y;
                        dist = Math.Sqrt(dx * dx + dy * dy);
                    }
                    else
                    {
                        double cx = (min.X + max.X) * 0.5;
                        double cy = (min.Y + max.Y) * 0.5;
                        double dx = pt.X - cx;
                        double dy = pt.Y - cy;
                        dist = Math.Sqrt(dx * dx + dy * dy);
                    }

                    double score = insideXY ? dist - 1000.0 : dist;
                    if (score < bestScore)
                    {
                        bestScore = score;
                        best = p;
                    }
                }
                catch { }
            }
            return best;
        }

        private static bool TryGetClosestPointOnPartCenterlineXY(Part part, Point refPt, out Point closest)
        {
            closest = null;
            if (part == null || refPt == null) return false;

            try
            {
                if (part is Beam b)
                {
                    Point a = b.StartPoint;
                    Point z = b.EndPoint;
                    if (a == null || z == null) return false;

                    double vx = z.X - a.X;
                    double vy = z.Y - a.Y;
                    double vv = vx * vx + vy * vy;
                    if (vv < 1e-9) return false;

                    double t = ((refPt.X - a.X) * vx + (refPt.Y - a.Y) * vy) / vv;
                    if (t < 0.0) t = 0.0;
                    else if (t > 1.0) t = 1.0;

                    closest = new Point(a.X + vx * t, a.Y + vy * t, refPt.Z);
                    return true;
                }
            }
            catch { }

            return false;
        }


        // ---------------- Imperial parsing (strict) ----------------

    }
}
