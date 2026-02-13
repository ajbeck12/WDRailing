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
        private bool TryCreatePostConnection(Part host, Part post, string connNameOrNumber, string attrFile)
        {
            try
            {
                if (host == null || post == null) return false;
                if (string.IsNullOrWhiteSpace(connNameOrNumber))
                    throw new InvalidOperationException("ConnectionName is blank.");

                // IMPORTANT: insert as a true Connection (main + secondary), not a generic Component.
                // Try host as main first, then swap if that fails.
                return TryInsertConnection(host, post, connNameOrNumber, attrFile)
                    || TryInsertConnection(post, host, connNameOrNumber, attrFile);
            }
            catch (Exception ex)
            {
                Operation.DisplayPrompt("WDRailing: Connection failed: " + ex.Message);
                return false;
            }
        }


        private bool TryInsertConnection(Part mainPart, Part secondaryPart, string connNameOrNumber, string attrFile)
        {
            var conn = new Tekla.Structures.Model.Connection();

            if (int.TryParse(connNameOrNumber.Trim(), out int num))
                conn.Number = num;
            else
                conn.Name = connNameOrNumber.Trim();

            // This matches Tekla UI behavior: "Main part" then "Secondary part"
            conn.SetPrimaryObject(mainPart);
            conn.SetSecondaryObject(secondaryPart);

            if (!string.IsNullOrWhiteSpace(attrFile))
            {
                InvokeLoadAttributes(conn, attrFile.Trim()); // use reflection; Connection is not a Component in this TS version
            }

            return conn.Insert();
        }

        // ---------------- Post cap ----------------


        private static ContourPlate CreatePostCap(Part post)
        {
            try
            {
                if (post == null) return null;

                // Use the post's solid extents so the cap always matches the post size and respects rotation.
                Solid s = post.GetSolid();
                if (s == null) return null;

                double minX = s.MinimumPoint.X;
                double minY = s.MinimumPoint.Y;
                double maxX = s.MaximumPoint.X;
                double maxY = s.MaximumPoint.Y;
                double topZ = s.MaximumPoint.Z;

                // Plate thickness 1/8" = 3.175 mm. Contour points define the plate plane.
                // Place the plane at topZ + t/2 so the bottom face sits on the top of the post.
                double tMm = InchesToMm(0.125);
                double z = topZ + (tMm * 0.5);

                // Tekla profile naming varies by environment; try a couple common plate-profile strings.
                string[] plateProfiles = new string[] { "PL3.175" };
                foreach (string plateProfile in plateProfiles)
                {
                    var cap = new ContourPlate();
                    cap.Profile.ProfileString = plateProfile;
                    cap.Material.MaterialString = "A36";
                    cap.Class = "4";
                    cap.Name = "POST CAP";

                    cap.Position.Plane = Position.PlaneEnum.MIDDLE;
                    cap.Position.Rotation = Position.RotationEnum.TOP;
                    cap.Position.Depth = Position.DepthEnum.MIDDLE;

                    cap.Contour.AddContourPoint(new ContourPoint(new Point(minX, minY, z), null));
                    cap.Contour.AddContourPoint(new ContourPoint(new Point(maxX, minY, z), null));
                    cap.Contour.AddContourPoint(new ContourPoint(new Point(maxX, maxY, z), null));
                    cap.Contour.AddContourPoint(new ContourPoint(new Point(minX, maxY, z), null));

                    if (cap.Insert())
                        return cap;
                }

                return null;
            }
            catch
            {
                // Keep caps best-effort; do not fail the whole plugin run.
                return null;
            }
        }



        private static void InvokeLoadAttributes(object obj, string file)
        {
            if (obj == null) throw new ArgumentNullException(nameof(obj));

            var t = obj.GetType();

            // Most TS objects that support attribute files expose this method
            var m = t.GetMethod("LoadAttributesFromFile", new[] { typeof(string) });
            if (m == null)
                throw new MissingMethodException(t.FullName + ".LoadAttributesFromFile(string) not found.");

            object r = m.Invoke(obj, new object[] { file });

            // If method returns bool, enforce it.
            if (m.ReturnType == typeof(bool) && r is bool ok && !ok)
                throw new InvalidOperationException("LoadAttributesFromFile failed for: " + file);
        }

        // ---------------- Rail creation (multi-rail) ----------------

    }
}
