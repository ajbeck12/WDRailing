using System;
using System.Drawing;
using System.Windows.Forms;
using Tekla.Structures.Dialog;

namespace WDRailing
{
    public class WDRailingDialog : PluginFormBase
    {
        private TextBox _tbSpacingIn, _tbPostHeightIn, _tbStartOffsetIn, _tbEndOffsetIn, _tbBaseOffsetIn;
        private ComboBox _cbLineRef;
        private TextBox _tbDeckEdgeIn;

        private TextBox _tbProfile, _tbMaterial, _tbClass, _tbName;

        // Connection
        private ComboBox _cbConnEnabled;
        private TextBox _tbConnName, _tbConnAttr;

        // hidden
        private TextBox _bfP1X, _bfP1Y, _bfP1Z, _bfP2X, _bfP2Y, _bfP2Z;
        private TextBox _bfHostIds1, _bfHostIds2, _bfHostIds3, _bfHostIds4;

        private WDRailingDefaults _cfg;

        // Horizontal Rails
        private ComboBox _cbRailEnabled;
        private TextBox _tbRailStartOffsetIn, _tbRailEndOffsetIn, _tbRailFromTopIn;
        private TextBox _tbRailCount, _tbRailSpacingIn;

        // Seat angle / post hole pattern (editable)
        private TextBox _tbSeatHoleLineIn;

        private TextBox _tbSeatSlotC2CIn, _tbSeatSlotSizeIn, _tbSeatSlotStandard, _tbSeatSlotCutLenIn;
        private ComboBox _cbSeatSlotSpecial1;

        private TextBox _tbSeatPilotC2CIn, _tbSeatPilotDiaIn, _tbSeatPilotStandard, _tbSeatPilotCutLenIn;

        private TextBox _bfRunPts1, _bfRunPts2, _bfRunPts3, _bfRunPts4;

        public WDRailingDialog()
        {
            BuildUi();
            Shown += OnShown;
        }

        private void OnShown(object sender, EventArgs e)
        {
            _cfg = WDRailingDefaults.LoadOrThrow();

            try { base.Get(); } catch { }

            ApplyConfigToEmptyFields();
            FormatAllDistances();
        }

        private void BuildUi()
        {
            Text = "WDRailing";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            TopMost = true;
            ClientSize = new Size(780, 540);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 2
            };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48f));

            var tabs = new TabControl { Dock = DockStyle.Fill };

            // ---------------- Create controls + bind ----------------

            _tbSpacingIn = NewText(); BindString(_tbSpacingIn, "SPACING_IN");
            _tbPostHeightIn = NewText(); BindString(_tbPostHeightIn, "POST_HEIGHT_IN");
            _tbStartOffsetIn = NewText(); BindString(_tbStartOffsetIn, "START_OFFSET_IN");
            _tbEndOffsetIn = NewText(); BindString(_tbEndOffsetIn, "END_OFFSET_IN");
            _tbBaseOffsetIn = NewText(); BindString(_tbBaseOffsetIn, "BASE_OFFSET_IN");

            _cbLineRef = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _cbLineRef.Items.AddRange(new object[] { "MIDDLE", "LEFT", "RIGHT" });
            BindString(_cbLineRef, "LINE_REF", "SelectedItem");

            _tbDeckEdgeIn = NewText(); BindString(_tbDeckEdgeIn, "DECK_EDGE_IN");

            _tbProfile = NewText(); BindString(_tbProfile, "POST_PROFILE");
            _tbMaterial = NewText(); BindString(_tbMaterial, "POST_MATERIAL");
            _tbClass = NewText(); BindString(_tbClass, "POST_CLASS");
            _tbName = NewText(); BindString(_tbName, "POST_NAME");

            _cbConnEnabled = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _cbConnEnabled.Items.AddRange(new object[] { "0", "1" });
            BindString(_cbConnEnabled, "CONN_ENABLED", "SelectedItem");

            _tbConnName = NewText(); BindString(_tbConnName, "CONN_NAME");
            _tbConnAttr = NewText(); BindString(_tbConnAttr, "CONN_ATTR");

            _cbRailEnabled = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _cbRailEnabled.Items.AddRange(new object[] { "0", "1" });
            BindString(_cbRailEnabled, "RAIL_ENABLED", "SelectedItem");

            _tbRailStartOffsetIn = NewText(); BindString(_tbRailStartOffsetIn, "RAIL_START_OFF_IN");
            _tbRailEndOffsetIn = NewText(); BindString(_tbRailEndOffsetIn, "RAIL_END_OFFSET_IN");
            _tbRailFromTopIn = NewText(); BindString(_tbRailFromTopIn, "RAIL_FROM_TOP_IN");

            _tbRailCount = NewText(); BindString(_tbRailCount, "RAIL_COUNT");
            _tbRailSpacingIn = NewText(); BindString(_tbRailSpacingIn, "RAIL_SPACING_IN");

            _tbSeatHoleLineIn = NewText(); BindString(_tbSeatHoleLineIn, "SEAT_HOLELINE_IN");

            _tbSeatSlotC2CIn = NewText(); BindString(_tbSeatSlotC2CIn, "SEAT_SLOT_C2C_IN");
            _tbSeatSlotSizeIn = NewText(); BindString(_tbSeatSlotSizeIn, "SEAT_SLOT_SIZE_IN");
            _tbSeatSlotStandard = NewText(); BindString(_tbSeatSlotStandard, "SEAT_SLOT_STANDARD");
            _tbSeatSlotCutLenIn = NewText(); BindString(_tbSeatSlotCutLenIn, "SEAT_SLOT_CUTLEN_IN");
            _cbSeatSlotSpecial1 = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _cbSeatSlotSpecial1.Items.AddRange(new object[] { "0", "1" });
            BindString(_cbSeatSlotSpecial1, "SEAT_SLOT_SPECIAL1", "Text");

            _tbSeatPilotC2CIn = NewText(); BindString(_tbSeatPilotC2CIn, "SEAT_PILOT_C2C_IN");
            _tbSeatPilotDiaIn = NewText(); BindString(_tbSeatPilotDiaIn, "SEAT_PILOT_DIA_IN");
            _tbSeatPilotStandard = NewText(); BindString(_tbSeatPilotStandard, "SEAT_PILOT_STANDARD");
            _tbSeatPilotCutLenIn = NewText(); BindString(_tbSeatPilotCutLenIn, "SEAT_PILOT_CUTLEN_IN");

            // ---------------- Formatting hooks ----------------

            HookFmt(_tbRailSpacingIn, false);
            HookFmt(_tbRailStartOffsetIn, true);
            HookFmt(_tbRailEndOffsetIn, true);
            HookFmt(_tbRailFromTopIn, false);

            HookFmt(_tbSeatHoleLineIn, false);

            HookFmt(_tbSeatSlotC2CIn, false);
            HookFmt(_tbSeatSlotSizeIn, false);
            HookFmt(_tbSeatSlotCutLenIn, false);

            HookFmt(_tbSeatPilotC2CIn, false);
            HookFmt(_tbSeatPilotDiaIn, false);
            HookFmt(_tbSeatPilotCutLenIn, false);

            HookFmt(_tbSpacingIn, false);
            HookFmt(_tbPostHeightIn, false);
            HookFmt(_tbStartOffsetIn, true);
            HookFmt(_tbEndOffsetIn, true);
            HookFmt(_tbBaseOffsetIn, true);
            HookFmt(_tbDeckEdgeIn, true);

            // ---------------- Tabs ----------------

            // General tab
            {
                var tab = NewTab("General", out var table);
                int r = 0;
                AddRow(table, r++, "Target spacing", _tbSpacingIn);
                AddRow(table, r++, "Post height", _tbPostHeightIn);
                AddRow(table, r++, "Start offset", _tbStartOffsetIn);
                AddRow(table, r++, "End offset", _tbEndOffsetIn);
                AddRow(table, r++, "Base offset (can be negative)", _tbBaseOffsetIn);
                tabs.TabPages.Add(tab);
            }

            // Placement tab
            {
                var tab = NewTab("Placement", out var table);
                int r = 0;
                AddRow(table, r++, "Line reference (Start ➜ End)", _cbLineRef);
                AddRow(table, r++, "Deck edge (face-to-line)", _tbDeckEdgeIn);
                tabs.TabPages.Add(tab);
            }

            // Post tab
            {
                var tab = NewTab("Post", out var table);
                int r = 0;
                AddRow(table, r++, "Post profile", _tbProfile);
                AddRow(table, r++, "Post material", _tbMaterial);
                AddRow(table, r++, "Post class", _tbClass);
                AddRow(table, r++, "Post name", _tbName);
                tabs.TabPages.Add(tab);
            }

            // Connection tab
            {
                var tab = NewTab("Connection", out var table);
                int r = 0;
                AddRow(table, r++, "Create connection (0/1)", _cbConnEnabled);
                AddRow(table, r++, "Connection name or number", _tbConnName);
                AddRow(table, r++, "Connection attributes file (optional)", _tbConnAttr);
                tabs.TabPages.Add(tab);
            }

            // Rail tab
            {
                var tab = NewTab("Rail", out var table);
                int r = 0;
                AddRow(table, r++, "Create rail (0/1)", _cbRailEnabled);
                AddRow(table, r++, "Rail start offset", _tbRailStartOffsetIn);
                AddRow(table, r++, "Rail end offset", _tbRailEndOffsetIn);
                AddRow(table, r++, "Rail center down from top of post", _tbRailFromTopIn);
                AddRow(table, r++, "Rail count", _tbRailCount);
                AddRow(table, r++, "Rail spacing (c/c)", _tbRailSpacingIn);
                tabs.TabPages.Add(tab);
            }

            // Seat/Post holes tab
            {
                var tab = NewTab("Seat / Post Holes", out var table);
                int r = 0;
                AddRow(table, r++, "Hole line from outside bend (down)", _tbSeatHoleLineIn);

                AddRow(table, r++, "Slot size", _tbSeatSlotSizeIn);
                AddRow(table, r++, "Slot standard", _tbSeatSlotStandard);
                AddRow(table, r++, "Slot cut length", _tbSeatSlotCutLenIn);
                AddRow(table, r++, "Slot special hole (first layer 0/1)", _cbSeatSlotSpecial1);
                AddRow(table, r++, "Slotted hole X", _tbSeatSlotC2CIn);

                AddRow(table, r++, "Pilot hole diameter", _tbSeatPilotDiaIn);
                AddRow(table, r++, "Pilot standard", _tbSeatPilotStandard);
                AddRow(table, r++, "Pilot cut length", _tbSeatPilotCutLenIn);
                AddRow(table, r++, "Pilot bolt distance X", _tbSeatPilotC2CIn);

                tabs.TabPages.Add(tab);
            }

            // ---------------- Hidden bound fields (keep these on the form) ----------------


            _bfP1X = Hidden(); _bfP1Y = Hidden(); _bfP1Z = Hidden();
            _bfP2X = Hidden(); _bfP2Y = Hidden(); _bfP2Z = Hidden();
            _bfHostIds1 = Hidden(); _bfHostIds2 = Hidden(); _bfHostIds3 = Hidden(); _bfHostIds4 = Hidden();

            BindString(_bfP1X, "P1X"); BindString(_bfP1Y, "P1Y"); BindString(_bfP1Z, "P1Z");
            BindString(_bfP2X, "P2X"); BindString(_bfP2Y, "P2Y"); BindString(_bfP2Z, "P2Z");
            BindString(_bfHostIds1, "HOST_IDS1"); BindString(_bfHostIds2, "HOST_IDS2");
            BindString(_bfHostIds3, "HOST_IDS3"); BindString(_bfHostIds4, "HOST_IDS4");

            _bfRunPts1 = Hidden(); _bfRunPts2 = Hidden(); _bfRunPts3 = Hidden(); _bfRunPts4 = Hidden();
            BindString(_bfRunPts1, "RUNPTS1");
            BindString(_bfRunPts2, "RUNPTS2");
            BindString(_bfRunPts3, "RUNPTS3");
            BindString(_bfRunPts4, "RUNPTS4");

            var hiddenPanel = new Panel { Visible = false, Width = 1, Height = 1 };
            hiddenPanel.Controls.AddRange(new Control[] { _bfRunPts1, _bfRunPts2, _bfRunPts3, _bfRunPts4 });

            // ---------------- Compose ----------------

            root.Controls.Add(tabs, 0, 0);
            root.Controls.Add(BuildButtonsPanel(), 0, 1);

            Controls.Add(root);
            Controls.Add(hiddenPanel); // keep bound hidden fields alive
        }


        private Control BuildButtonsPanel()
        {
            var btnPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Height = 40 };

            var btnOk = new Button { Text = "OK", Width = 90 };
            var btnApply = new Button { Text = "Apply", Width = 90 };
            var btnModify = new Button { Text = "Modify", Width = 90 };
            var btnGet = new Button { Text = "Get", Width = 90 };
            var btnCancel = new Button { Text = "Cancel", Width = 90 };

            btnGet.Click += (s, e) => { try { base.Get(); } catch { } ApplyConfigToEmptyFields(); FormatAllDistances(); };
            btnApply.Click += (s, e) => { ApplyConfigToEmptyFields(); FormatAllDistances(); Apply(); };
            btnModify.Click += (s, e) => { ApplyConfigToEmptyFields(); FormatAllDistances(); Modify(); };
            btnOk.Click += (s, e) => { ApplyConfigToEmptyFields(); FormatAllDistances(); Apply(); Close(); };
            btnCancel.Click += (s, e) => Close();

            btnPanel.Controls.AddRange(new Control[] { btnOk, btnApply, btnModify, btnGet, btnCancel });
            return btnPanel;
        }

        private void ApplyConfigToEmptyFields()
        {
            if (_cfg == null) _cfg = WDRailingDefaults.LoadOrThrow();

            SetIfEmpty(_tbSpacingIn, _cfg.SpacingIn);
            SetIfEmpty(_tbPostHeightIn, _cfg.PostHeightIn);
            SetIfEmpty(_tbStartOffsetIn, _cfg.StartOffsetIn);
            SetIfEmpty(_tbEndOffsetIn, _cfg.EndOffsetIn);
            SetIfEmpty(_tbBaseOffsetIn, _cfg.BaseOffsetIn);
            SetIfEmpty(_tbDeckEdgeIn, _cfg.DeckEdgeIn);

            if ((_cbLineRef.SelectedItem == null && string.IsNullOrWhiteSpace(_cbLineRef.Text)) && !string.IsNullOrWhiteSpace(_cfg.LineRef))
                _cbLineRef.SelectedItem = _cfg.LineRef.Trim().ToUpperInvariant();

            SetIfEmpty(_tbProfile, _cfg.PostProfile);
            SetIfEmpty(_tbMaterial, _cfg.PostMaterial);
            SetIfEmpty(_tbClass, _cfg.PostClass);
            SetIfEmpty(_tbName, _cfg.PostName);

            if ((_cbConnEnabled.SelectedItem == null && string.IsNullOrWhiteSpace(_cbConnEnabled.Text)) && !string.IsNullOrWhiteSpace(_cfg.CreateConnection))
                _cbConnEnabled.SelectedItem = _cfg.CreateConnection.Trim();

            SetIfEmpty(_tbConnName, _cfg.ConnectionName);
            if (_tbConnAttr.Text == null) _tbConnAttr.Text = "";
            if (string.IsNullOrWhiteSpace(_tbConnAttr.Text) && _cfg.ConnectionAttr != null) _tbConnAttr.Text = _cfg.ConnectionAttr;

            if ((_cbRailEnabled.SelectedItem == null && string.IsNullOrWhiteSpace(_cbRailEnabled.Text)) && !string.IsNullOrWhiteSpace(_cfg.RailEnabled))
                _cbRailEnabled.SelectedItem = _cfg.RailEnabled.Trim();

            SetIfEmpty(_tbRailStartOffsetIn, _cfg.RailStartOffsetIn);
            SetIfEmpty(_tbRailEndOffsetIn, _cfg.RailEndOffsetIn);
            SetIfEmpty(_tbRailFromTopIn, _cfg.RailFromTopIn);
            SetIfEmpty(_tbRailCount, _cfg.RailCount);
            SetIfEmpty(_tbRailSpacingIn, _cfg.RailSpacingIn);

            SetIfEmpty(_tbSeatHoleLineIn, _cfg.SeatHoleLineFromBendIn);

            SetIfEmpty(_tbSeatSlotC2CIn, _cfg.SeatSlotC2CIn);
            SetIfEmpty(_tbSeatSlotSizeIn, _cfg.SeatSlotSizeIn);
            SetIfEmpty(_tbSeatSlotStandard, _cfg.SeatSlotStandard);
            SetIfEmpty(_tbSeatSlotCutLenIn, _cfg.SeatSlotCutLengthIn);
            if ((_cbSeatSlotSpecial1.SelectedItem == null && string.IsNullOrWhiteSpace(_cbSeatSlotSpecial1.Text)) && !string.IsNullOrWhiteSpace(_cfg.SeatSlotSpecial1))
                _cbSeatSlotSpecial1.Text = _cfg.SeatSlotSpecial1.Trim();

            SetIfEmpty(_tbSeatPilotC2CIn, _cfg.SeatPilotC2CIn);
            SetIfEmpty(_tbSeatPilotDiaIn, _cfg.SeatPilotDiaIn);
            SetIfEmpty(_tbSeatPilotStandard, _cfg.SeatPilotStandard);
            SetIfEmpty(_tbSeatPilotCutLenIn, _cfg.SeatPilotCutLengthIn);
        }

        private void FormatAllDistances()
        {
            Fmt(_tbSpacingIn, false);
            Fmt(_tbPostHeightIn, false);
            Fmt(_tbStartOffsetIn, true);
            Fmt(_tbEndOffsetIn, true);
            Fmt(_tbBaseOffsetIn, true);
            Fmt(_tbDeckEdgeIn, true);
            Fmt(_tbRailStartOffsetIn, true);
            Fmt(_tbRailEndOffsetIn, true);
            Fmt(_tbRailFromTopIn, false);
            Fmt(_tbRailSpacingIn, false);

            Fmt(_tbSeatHoleLineIn, false);

            Fmt(_tbSeatSlotC2CIn, false);
            Fmt(_tbSeatSlotSizeIn, false);
            Fmt(_tbSeatSlotCutLenIn, false);

            Fmt(_tbSeatPilotC2CIn, false);
            Fmt(_tbSeatPilotDiaIn, false);
            Fmt(_tbSeatPilotCutLenIn, false);
        }

        private void HookFmt(TextBox tb, bool allowNeg) => tb.Leave += (s, e) => Fmt(tb, allowNeg);

        private void Fmt(TextBox tb, bool allowNeg)
        {
            if (tb == null) return;
            if (string.IsNullOrWhiteSpace(tb.Text)) return;

            try
            {
                // best effort parsing using main parser style not available here; just reuse DistanceFormat formatting on numeric.
                // If user types Tekla-like, keep as-is if parse fails.
                double inches;
                if (!TryParseLooseDistance(tb.Text, out inches, allowNeg)) return;
                tb.Text = (inches < 0)
                    ? "-" + DistanceFormat.ToTeklaFeetInches(Math.Abs(inches), 16)
                    : DistanceFormat.ToTeklaFeetInches(inches, 16);
            }
            catch { }
        }

        private bool TryParseLooseDistance(string raw, out double inches, bool allowNeg)
        {
            inches = 0.0;
            string s = (raw ?? "").Trim();
            if (s.Length == 0) return false;

            // very small parser: try numeric directly first
            double v;
            if (double.TryParse(s.Replace("\"", "").Replace("'", ""), out v))
            {
                inches = v;
                if (!allowNeg && inches <= 0) return false;
                return true;
            }
            return false;
        }

        private static void SetIfEmpty(TextBox tb, string val)
        {
            if (tb == null) return;
            if (string.IsNullOrWhiteSpace(tb.Text)) tb.Text = val ?? "";
        }

        private static TextBox NewText() => new TextBox { Dock = DockStyle.Fill, Text = "" };
        private static TextBox Hidden() => new TextBox { Visible = false, TabStop = false, Text = "" };

        private void BindString(Control c, string attr) => BindString(c, attr, null);

        private void BindString(Control c, string attr, string bindPropertyName)
        {
            structuresExtender.SetAttributeName(c, attr);
            structuresExtender.SetAttributeTypeName(c, "String");
            structuresExtender.SetBindPropertyName(c, bindPropertyName);
        }

        private static void AddRow(TableLayoutPanel table, int row, string label, Control control)
        {
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            var lbl = new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Padding = new Padding(0, 8, 8, 0) };
            table.Controls.Add(lbl, 0, row);
            table.Controls.Add(control, 1, row);
        }

        private TabPage NewTab(string title, out TableLayoutPanel table)
        {
            var tab = new TabPage(title);

            var panel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
            tab.Controls.Add(panel);

            table = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                Dock = DockStyle.Top
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 44f));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 56f));

            panel.Controls.Add(table);
            return tab;
        }
    }
}
