using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace WordColorTool
{
    /// <summary>
    /// 一个颜色。色板配置用。
    /// </summary>
    public class ColorBinding
    {
        public string HexColor;  // 6 位十六进制 RRGGBB，例如 "FF0000"
        public string Name;      // 显示名字（悬停提示）

        public ColorBinding(string hexColor, string name)
        {
            this.HexColor = hexColor;
            this.Name = name;
        }
    }

    /// <summary>
    /// 一条用户保存的格式预设：颜色 + 字号 + 字体名。
    /// Size 为 0 表示不限定字号；FontName 为 null 表示不限定字体。
    /// </summary>
    public class Preset
    {
        public string Name;
        public string HexColor;      // 空 = 自动颜色
        public float Size;
        public string FontName;
    }

    public class MainForm : Form
    {
        // =====================================================================
        //  用户配置区 —— 色板。想加颜色就往数组里整行加：new ColorBinding("RRGGBB", "")
        //  第二个参数是悬停提示文字（可留空）。色板固定按 8 列网格排列，
        //  增删颜色不影响其它色块能否点到。
        // =====================================================================
        private readonly ColorBinding[] ColorConfig = new ColorBinding[]
        {
            new ColorBinding("E60000", "鲜红"),      // 批改红：鲜艳、不浅
            new ColorBinding("FF0000", "正红"),
            new ColorBinding("FF2A2A", "亮红"),
            new ColorBinding("C00000", "深红"),
            new ColorBinding("FF6600", "橙红"),
            new ColorBinding("FF8000", "橙"),
            new ColorBinding("FFA500", "橘黄"),
            new ColorBinding("FFC000", "琥珀"),
            new ColorBinding("FFFF00", "明黄"),
            new ColorBinding("CCFF00", "黄绿"),
            new ColorBinding("92D050", "浅绿"),
            new ColorBinding("00B050", "绿"),
            new ColorBinding("00B0F0", "天蓝"),
            new ColorBinding("0070C0", "蓝"),
            new ColorBinding("002060", "深蓝"),
            new ColorBinding("7030A0", "紫"),
            new ColorBinding("9400D3", "亮紫"),
            new ColorBinding("FF00FF", "品红"),
            new ColorBinding("FF6699", "粉"),
            new ColorBinding("FF1493", "玫红"),
            new ColorBinding("00FFFF", "青"),
            new ColorBinding("008080", "深青"),
            new ColorBinding("808080", "灰"),
            new ColorBinding("000000", "黑"),
        };

        // ===== 配置区结束 =====

        private const long WdColorAutomatic = -16777216;   // Word 的 wdColorAutomatic
        private const int LockCheckIntervalMs = 300;       // 锁定检查间隔

        // 色板网格
        private const int SwX = 10;      // 色板左上角 X
        private const int SwY = 58;      // 色板左上角 Y
        private const int SwSize = 26;   // 色块边长
        private const int SwGap = 5;     // 色块间距
        private const int SwCols = 8;    // 每行几个

        // ---- 控件 ----
        private NotifyIcon _trayIcon;
        private ContextMenuStrip _trayMenu;

        private Panel _titleBar;
        private Label _lblTitle;
        private Panel _previewBox;
        private Label _lblSelected;
        private ComboBox _cmbSize;
        private ComboBox _cmbFont;
        private ListBox _lstPresets;
        private TextBox _txtPresetName;
        private Label _lblStatus;

        private Button _selectedSwatch;
        private bool _dragging;
        private Point _dragStartPoint = new Point(0, 0);

        // 面板里当前选中的颜色（还没点确定，只是待定状态）
        private bool _selAutoColor = true;
        private uint _selRgb;

        private readonly List<Preset> _presets = new List<Preset>();

        // ---- 锁定状态（点「确定」后生效）----
        private bool _lockActive;
        private bool _lockColorAuto;
        private uint _lockColorRgb;
        private float _lockSize;          // 0 = 不限定字号
        private string _lockFontName;      // null = 不限定字体
        private int _lockStartPos = -1;
        private bool _appliedNoticeShown;  // 「已应用」气泡只提示一次

        // 缓存的 Word COM 对象；失效时置空，下一轮重新获取
        private object _wordApp;

        private readonly Timer _lockTimer = new Timer();
        private bool _busy;
        private bool _shuttingDown;
        private bool _textInputActive;   // 正在输入（临时允许窗口被激活）

        public MainForm()
        {
            this.Text = "WordColorTool";
            this.FormBorderStyle = FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.StartPosition = FormStartPosition.CenterScreen;   // 默认屏幕正中，保证看得见
            this.ClientSize = new Size(320, 420);
            this.AutoScaleMode = AutoScaleMode.Dpi;   // 高 DPI 屏幕下控件等比缩放，不会错位
            this.FormClosing += OnFormClosing;

            BuildPanel();
            BuildTrayIcon();
            LoadPresets();
            StartLockTimer();
        }

        /// <summary>窗口显示后校正位置，确保整个面板都在屏幕内。</summary>
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            EnsureOnScreen();
        }

        private void EnsureOnScreen()
        {
            try
            {
                Rectangle vs = SystemInformation.VirtualScreen;
                if (vs.Width <= 0 || vs.Height <= 0)
                {
                    return;
                }
                int x = Math.Max(vs.Left, Math.Min(this.Left, vs.Right - this.Width));
                int y = Math.Max(vs.Top, Math.Min(this.Top, vs.Bottom - this.Height));
                if (x != this.Left || y != this.Top)
                {
                    this.Location = new Point(x, y);
                }
            }
            catch (Exception)
            {
            }
        }

        // ================================================================
        //  悬浮窗：不抢焦点、不占任务栏（窗口内控件仍可正常点击）
        // ================================================================

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x08000000;    // WS_EX_NOACTIVATE
                cp.ExStyle |= 0x00000080;    // WS_EX_TOOLWINDOW
                return cp;
            }
        }

        protected override bool ShowWithoutActivation
        {
            get { return true; }
        }

        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_NOACTIVATE = 0x08000000;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        /// <summary>
        /// 输入框拿到焦点时临时允许窗口被激活（否则 NOACTIVATE 窗口里的 TextBox
        /// 无法接收键盘输入）。焦点离开时恢复，Word 只在打字那会儿让出焦点。
        /// </summary>
        private void OnTextInputGotFocus(object sender, EventArgs e)
        {
            SetNoActivate(false);
        }

        private void OnTextInputLostFocus(object sender, EventArgs e)
        {
            SetNoActivate(true);
        }

        private void SetNoActivate(bool on)
        {
            if (!this.IsHandleCreated || _shuttingDown)
            {
                return;
            }

            try
            {
                int ex = GetWindowLong(this.Handle, GWL_EXSTYLE);
                if (on)
                {
                    SetWindowLong(this.Handle, GWL_EXSTYLE, ex | WS_EX_NOACTIVATE);
                    _textInputActive = false;
                }
                else
                {
                    SetWindowLong(this.Handle, GWL_EXSTYLE, ex & ~WS_EX_NOACTIVATE);
                    _textInputActive = true;
                    SetForegroundWindow(this.Handle);
                    this.Activate();
                }
            }
            catch (Exception)
            {
            }
        }

        // ================================================================
        //  界面
        // ================================================================

        private void BuildPanel()
        {
            Font labelFont = new Font(SystemFonts.DefaultFont, FontStyle.Bold);

            // ---- 标题栏（只有左键按在这里才能拖动）----
            _titleBar = new Panel();
            _titleBar.Location = new Point(0, 0);
            _titleBar.Size = new Size(320, 28);
            _titleBar.BackColor = Color.FromArgb(45, 45, 48);
            _titleBar.MouseDown += OnTitleBarMouseDown;
            _titleBar.MouseMove += OnTitleBarMouseMove;
            _titleBar.MouseUp += OnTitleBarMouseUp;
            this.Controls.Add(_titleBar);

            _lblTitle = MakeLabel("  WordColorTool", 0, 6, 200, 16, null);
            _lblTitle.ForeColor = Color.White;
            _lblTitle.Font = labelFont;
            _lblTitle.MouseDown += OnTitleBarMouseDown;
            _lblTitle.MouseMove += OnTitleBarMouseMove;
            _lblTitle.MouseUp += OnTitleBarMouseUp;
            _titleBar.Controls.Add(_lblTitle);

            Label lblHideTip = MakeLabel("左键按住标题栏可拖动", 210, 7, 104, 14, null);
            lblHideTip.ForeColor = Color.FromArgb(150, 150, 150);
            lblHideTip.Font = new Font(SystemFonts.DefaultFont.FontFamily, 7f, FontStyle.Regular);
            lblHideTip.TextAlign = ContentAlignment.MiddleRight;
            _titleBar.Controls.Add(lblHideTip);

            // ---- 颜色区 ----
            Label lblColor = MakeLabel("① 颜色（选中后点下方「确定」）", 10, 34, 240, 16, labelFont);
            this.Controls.Add(lblColor);

            // 当前色预览
            _previewBox = new Panel();
            _previewBox.Location = new Point(284, 32);
            _previewBox.Size = new Size(26, 20);
            _previewBox.BorderStyle = BorderStyle.FixedSingle;
            this.Controls.Add(_previewBox);

            _lblSelected = MakeLabel("自动", 278, 54, 38, 14, null);
            _lblSelected.Font = new Font(SystemFonts.DefaultFont.FontFamily, 7f, FontStyle.Regular);
            _lblSelected.TextAlign = ContentAlignment.MiddleCenter;
            this.Controls.Add(_lblSelected);

            // 色板：绝对定位网格，保证每个色块都点得到
            for (int i = 0; i < ColorConfig.Length; i++)
            {
                ColorBinding c = ColorConfig[i];
                if (c == null)
                {
                    continue;
                }
                int col = i % SwCols;
                int row = i / SwCols;

                Button b = new Button();
                b.Size = new Size(SwSize, SwSize);
                b.Location = new Point(SwX + col * (SwSize + SwGap), SwY + row * (SwSize + SwGap));
                b.FlatStyle = FlatStyle.Flat;
                b.FlatAppearance.BorderSize = 1;
                b.FlatAppearance.BorderColor = Color.Gray;
                b.Tag = c;
                b.BackColor = RgbFromHex(c.HexColor, Color.LightGray);
                b.Click += OnSwatchClick;
                b.MouseEnter += OnSwatchMouseEnter;
                b.MouseLeave += OnSwatchMouseLeave;

                ToolTip tip = new ToolTip();
                tip.SetToolTip(b, (c.Name == null || c.Name.Length == 0 ? "" : c.Name + " ") + c.HexColor);

                this.Controls.Add(b);
            }

            int swatchRows = (ColorConfig.Length + SwCols - 1) / SwCols;
            int afterSwatchY = SwY + swatchRows * (SwSize + SwGap) + 8;

            // ---- 格式区 ----
            Label lblFormat = MakeLabel("② 字号 / 字体", 10, afterSwatchY, 200, 18, labelFont);
            this.Controls.Add(lblFormat);

            int fmtY = afterSwatchY + 24;

            Label lblSize = MakeLabel("字号", 10, fmtY + 4, 28, 16, null);
            this.Controls.Add(lblSize);
            _cmbSize = new ComboBox();
            _cmbSize.Location = new Point(40, fmtY);
            _cmbSize.Size = new Size(72, 22);
            _cmbSize.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbSize.AccessibleName = "字号下拉";
            _cmbSize.Items.AddRange(new object[] { "默认", "8", "9", "10", "10.5", "11", "12", "14", "15", "16", "18", "22", "24", "26", "36", "42" });
            _cmbSize.SelectedIndex = 0;
            this.Controls.Add(_cmbSize);

            Label lblFont = MakeLabel("字体", 122, fmtY + 4, 28, 16, null);
            this.Controls.Add(lblFont);
            _cmbFont = new ComboBox();
            _cmbFont.Location = new Point(152, fmtY);
            _cmbFont.Size = new Size(158, 22);
            _cmbFont.DropDownStyle = ComboBoxStyle.DropDown;
            _cmbFont.AccessibleName = "字体下拉";
            _cmbFont.Items.AddRange(new object[] { "宋体", "黑体", "微软雅黑", "楷体", "仿宋", "Arial", "Times New Roman" });
            _cmbFont.Text = "";
            _cmbFont.Enter += OnTextInputGotFocus;
            _cmbFont.Leave += OnTextInputLostFocus;
            this.Controls.Add(_cmbFont);

            // ---- 操作按钮 ----
            int btnY = fmtY + 34;

            Button btnOk = MakeButton("确定", 10, btnY, 92, 30, OnOkClick);
            btnOk.Font = labelFont;
            this.Controls.Add(btnOk);

            Button btnClear = MakeButton("清除格式", 106, btnY, 92, 30, OnClearFormatClick);
            this.Controls.Add(btnClear);

            Button btnAuto = MakeButton("自动颜色", 202, btnY, 108, 30, OnAutoClick);
            this.Controls.Add(btnAuto);

            // ---- 预设区 ----
            int preY = btnY + 42;

            Label lblPresets = MakeLabel("③ 预设（点一下载入，再点「确定」）", 10, preY, 260, 16, labelFont);
            this.Controls.Add(lblPresets);

            _txtPresetName = new TextBox();
            _txtPresetName.Location = new Point(10, preY + 26);
            _txtPresetName.Size = new Size(148, 22);
            _txtPresetName.Enter += OnTextInputGotFocus;
            _txtPresetName.Leave += OnTextInputLostFocus;
            this.Controls.Add(_txtPresetName);

            Button btnSave = MakeButton("存为预设", 162, preY + 25, 84, 24, OnSavePresetClick);
            this.Controls.Add(btnSave);

            Button btnDelete = MakeButton("删除", 250, preY + 25, 60, 24, OnDeletePresetClick);
            this.Controls.Add(btnDelete);

            _lstPresets = new ListBox();
            _lstPresets.Location = new Point(10, preY + 56);
            _lstPresets.Size = new Size(300, 72);
            _lstPresets.IntegralHeight = false;
            _lstPresets.Font = new Font(SystemFonts.DefaultFont.FontFamily, 7.5f, FontStyle.Regular);
            _lstPresets.SelectedIndexChanged += OnPresetSelectionChanged;
            this.Controls.Add(_lstPresets);

            // ---- 状态与隐藏 ----
            _lblStatus = MakeLabel("当前：未锁定", 10, preY + 134, 220, 16, null);
            _lblStatus.Font = new Font(SystemFonts.DefaultFont.FontFamily, 7.5f, FontStyle.Regular);
            this.Controls.Add(_lblStatus);

            Button btnHide = MakeButton("隐藏", 250, preY + 130, 60, 26, OnHideClick);
            this.Controls.Add(btnHide);

            this.BackColor = Color.FromArgb(240, 240, 240);
            UpdateSelectionPreview();
        }

        // ================================================================
        //  拖动：只有左键按在标题栏上才拖
        // ================================================================

        private void OnTitleBarMouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
            {
                return;
            }
            _dragging = true;
            _dragStartPoint = e.Location;
            _titleBar.Capture = true;
        }

        private void OnTitleBarMouseMove(object sender, MouseEventArgs e)
        {
            if (!_dragging)
            {
                return;   // 没按左键就不拖
            }
            // e.Location 是相对发送控件的；换算到标题栏坐标
            Point p = e.Location;
            Control s = sender as Control;
            if (s != null && !object.ReferenceEquals(s, _titleBar))
            {
                p = _titleBar.PointToClient(s.PointToScreen(e.Location));
            }
            this.Location = new Point(
                this.Location.X + (p.X - _dragStartPoint.X),
                this.Location.Y + (p.Y - _dragStartPoint.Y));
        }

        private void OnTitleBarMouseUp(object sender, MouseEventArgs e)
        {
            _dragging = false;
            _titleBar.Capture = false;
            // 拖完松手若窗口被拖出屏幕，拉回可见区域（不然就找不着了）
            EnsureOnScreen();
        }

        // ================================================================
        //  托盘
        // ================================================================

        private void BuildTrayIcon()
        {
            _trayMenu = new ContextMenuStrip();

            ToolStripMenuItem showItem = new ToolStripMenuItem("显示/隐藏面板");
            showItem.Click += delegate(object s, EventArgs e) { ToggleVisible(); };

            ToolStripMenuItem unlockItem = new ToolStripMenuItem("解除锁定");
            unlockItem.Click += OnUnlockClick;

            ToolStripMenuItem exitItem = new ToolStripMenuItem("退出");
            exitItem.Click += OnExitClick;

            _trayMenu.Items.Add(showItem);
            _trayMenu.Items.Add(unlockItem);
            _trayMenu.Items.Add(new ToolStripSeparator());
            _trayMenu.Items.Add(exitItem);

            _trayIcon = new NotifyIcon();
            _trayIcon.Icon = SystemIcons.Application;
            _trayIcon.Text = "WordColorTool（左键显示/隐藏）";
            _trayIcon.ContextMenuStrip = _trayMenu;
            _trayIcon.Visible = true;
            _trayIcon.MouseClick += OnTrayMouseClick;
        }

        private void OnTrayMouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ToggleVisible();
            }
        }

        private void ToggleVisible()
        {
            if (this.Visible)
            {
                this.Hide();
            }
            else
            {
                RefreshPanelFromLock();
                EnsureOnScreen();          // 之前在屏幕外的话先拉回来
                SetNoActivate(false);      // 临时允许激活，确保能显示到最前
                this.Show();
                this.Activate();
                SetNoActivate(true);
            }
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            if (_shuttingDown)
            {
                return;
            }
            e.Cancel = true;
            this.Hide();
        }

        private void OnHideClick(object sender, EventArgs e)
        {
            this.Hide();
        }

        // ================================================================
        //  颜色交互（只改待选状态，点「确定」才应用）
        // ================================================================

        private void OnSwatchClick(object sender, EventArgs e)
        {
            Button b = sender as Button;
            if (b == null)
            {
                return;
            }
            ColorBinding c = b.Tag as ColorBinding;
            if (c == null)
            {
                return;
            }

            uint v;
            if (!TryParseHexColor(c.HexColor, out v))
            {
                return;
            }

            _selectedSwatch = b;
            UpdateSwatchBorders();
            SelectColor(v);
        }

        private void OnSwatchMouseEnter(object sender, EventArgs e)
        {
            Button b = sender as Button;
            if (b != null && !object.ReferenceEquals(b, _selectedSwatch))
            {
                b.FlatAppearance.BorderSize = 2;
            }
        }

        private void OnSwatchMouseLeave(object sender, EventArgs e)
        {
            Button b = sender as Button;
            if (b != null && !object.ReferenceEquals(b, _selectedSwatch))
            {
                b.FlatAppearance.BorderSize = 1;
            }
        }

        private void SelectColor(uint rgb)
        {
            _selAutoColor = false;
            _selRgb = rgb;
            UpdateSelectionPreview();
        }

        private void UpdateSwatchBorders()
        {
            for (int i = 0; i < Controls.Count; i++)
            {
                Button b = Controls[i] as Button;
                if (b == null || b.Size != new Size(SwSize, SwSize))
                {
                    continue;
                }
                b.FlatAppearance.BorderSize = object.ReferenceEquals(b, _selectedSwatch) ? 3 : 1;
            }
        }

        private void UpdateSelectionPreview()
        {
            if (_previewBox == null || _lblSelected == null)
            {
                return;
            }

            if (_selAutoColor)
            {
                _previewBox.BackColor = SystemColors.Control;
                _lblSelected.Text = "自动";
                return;
            }

            _previewBox.BackColor = Color.FromArgb(
                (int)(_selRgb & 0xFF),
                (int)((_selRgb >> 8) & 0xFF),
                (int)((_selRgb >> 16) & 0xFF));
            _lblSelected.Text = RgbToHex(_selRgb);
        }

        private void OnAutoClick(object sender, EventArgs e)
        {
            _selAutoColor = true;
            _selectedSwatch = null;
            UpdateSwatchBorders();
            UpdateSelectionPreview();
        }

        private void OnCustomColorClick(object sender, EventArgs e)
        {
            ColorDialog dlg = new ColorDialog();
            dlg.AnyColor = true;
            dlg.FullOpen = true;
            dlg.AllowFullOpen = true;

            // 本窗口是 NOACTIVATE 窗口，直接 ShowDialog(this) 有时对话框不跳到前台。
            // 先临时允许激活、把面板带到前台，关掉对话框后再恢复。
            SetNoActivate(false);
            DialogResult r;
            try
            {
                r = dlg.ShowDialog(this);
            }
            finally
            {
                SetNoActivate(true);
            }

            if (r == DialogResult.OK)
            {
                Color col = dlg.Color;
                _selectedSwatch = null;
                UpdateSwatchBorders();
                SelectColor((uint)(col.R | (col.G << 8) | (col.B << 16)));
            }
        }

        private float ReadSelectedSize()
        {
            object sel = _cmbSize.SelectedItem;
            if (sel == null)
            {
                return 0f;
            }
            string s = Convert.ToString(sel);
            if (string.IsNullOrEmpty(s) || s == "默认")
            {
                return 0f;
            }
            float f;
            if (float.TryParse(s, out f) && f > 0)
            {
                return f;
            }
            return 0f;
        }

        private string ReadSelectedFont()
        {
            string s = _cmbFont.Text == null ? "" : _cmbFont.Text.Trim();
            if (s.Length == 0)
            {
                return null;
            }
            return s;
        }

        // ================================================================
        //  确定 / 清除格式
        // ================================================================

        /// <summary>把面板设置应用到 Word 当前光标处并锁定。</summary>
        private void OnOkClick(object sender, EventArgs e)
        {
            object app = GetWordApp();
            if (app == null)
            {
                MessageBox.Show(this,
                    "还没有检测到正在运行的 Word。\n请先打开 Word，把光标放到要改格式的位置，再点「确定」。",
                    "WordColorTool", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 连到了 Word 但它没开任何文档 —— 典型是"后台残留实例"占着 Word 的进程注册，
            // 导致 GetActiveObject 拿到的是那个空实例，而不是你正在用的窗口。
            if (!WordHasDocuments(app))
            {
                HandleZombieWordInstance();
                return;
            }

            ApplyPanelToWord(app);
        }

        /// <summary>真正把面板设置写进 Word 并锁定（OnOkClick 和自动重试都走这里）。</summary>
        private void ApplyPanelToWord(object app)
        {
            if (app == null)
            {
                app = GetWordApp();
                if (app == null)
                {
                    MessageBox.Show(this, "还是没有连上 Word，请确认文档已打开后重试。",
                        "WordColorTool", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            _lockColorAuto = _selAutoColor;
            _lockColorRgb = _selRgb;
            _lockSize = ReadSelectedSize();
            _lockFontName = ReadSelectedFont();

            try
            {
                dynamic wordApp = app;
                dynamic sel = wordApp.Selection;
                _lockStartPos = GetInt(sel.Start, -1);

                ApplyLockToSelection(sel);
                _lockActive = true;

                if (!_appliedNoticeShown)
                {
                    _appliedNoticeShown = true;
                    ShowBalloon("已应用并锁定", DescribeLock(), ToolTipIcon.Info);
                }
            }
            catch (Exception)
            {
                InvalidateWordApp();
                _lockActive = false;
                MessageBox.Show(this,
                    "没能把格式写进 Word。常见原因：\n" +
                    "  1. 文档处于「受保护的视图」或只读（网上下载的文件常见）——先点 Word 顶部黄条里的「启用编辑」；\n" +
                    "  2. 选区在批注、页眉页脚或文本框里，这些位置可能不允许改字体颜色；\n" +
                    "  3. 连到的 Word 实例异常——重试一次，或重启 Word。",
                    "WordColorTool", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            UpdateStatus();
        }

        /// <summary>
        /// 处理"后台残留 Word 实例"：找出没有窗口的 WINWORD 进程，问用户后结束它们，
        /// 然后重新连接并自动重试。有窗口的实例（用户正在用的文档）绝不碰。
        /// </summary>
        private void HandleZombieWordInstance()
        {
            InvalidateWordApp();   // 先丢掉连错的实例

            List<int> zombiePids = new List<int>();
            try
            {
                System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName("WINWORD");
                for (int i = 0; i < procs.Length; i++)
                {
                    try
                    {
                        string title = procs[i].MainWindowTitle;
                        if (title == null || title.Trim().Length == 0)
                        {
                            zombiePids.Add(procs[i].Id);
                        }
                    }
                    catch (Exception)
                    {
                    }
                    finally
                    {
                        try { procs[i].Dispose(); } catch (Exception) { }
                    }
                }
            }
            catch (Exception)
            {
            }

            if (zombiePids.Count == 0)
            {
                MessageBox.Show(this,
                    "连到的 Word 实例没有打开任何文档。\n\n" +
                    "你正在用的 Word 窗口：\n" + GetVisibleWordWindows() +
                    "\n请在该窗口里点一下，然后再点「确定」。",
                    "WordColorTool", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult r = MessageBox.Show(this,
                "连上的 Word 是个后台残留实例（没有打开文档），\n" +
                "所以改不到你正在用的文档。检测到 " + zombiePids.Count + " 个无窗口的 Word 残留进程。\n\n" +
                "你正在用的 Word 窗口：\n" + GetVisibleWordWindows() +
                "\n是否结束这些残留进程并重试？（不会关闭你正在用的文档）",
                "WordColorTool", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (r != DialogResult.Yes)
            {
                return;
            }

            for (int i = 0; i < zombiePids.Count; i++)
            {
                try
                {
                    System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(zombiePids[i]);
                    p.Kill();
                    p.WaitForExit(3000);
                    p.Dispose();
                }
                catch (Exception)
                {
                    // 进程可能已经自己退了，忽略
                }
            }

            // 等 Word 的 ROT 重新注册到有文档的实例（最多约 2 秒）
            object app = null;
            for (int retry = 0; retry < 5; retry++)
            {
                System.Threading.Thread.Sleep(400);
                app = GetWordApp();
                if (app != null && WordHasDocuments(app))
                {
                    break;
                }
                app = null;
            }

            if (app == null)
            {
                MessageBox.Show(this,
                    "残留进程已结束，但还是没连上你的 Word。\n" +
                    "请在 Word 窗口里点一下，然后再点一次「确定」。",
                    "WordColorTool", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ApplyPanelToWord(app);
        }

        /// <summary>清除 Word 选区的颜色/字号/字体（恢复默认）并解除锁定。</summary>
        private void OnClearFormatClick(object sender, EventArgs e)
        {
            ClearLock();

            object app = GetWordApp();

            if (app == null)
            {
                // Word 没开：只解除锁定，面板复位，不动 Word
                ResetPanelSelection();
                ShowBalloon("已解除锁定", "没有检测到正在运行的 Word，Word 选区未改动。", ToolTipIcon.Info);
                return;
            }

            if (!WordHasDocuments(app))
            {
                // 连到的是后台残留实例：走同一套清理 + 重试流程
                HandleZombieWordInstance();
                return;
            }

            bool ok = DoClearWordFormat(app);
            ResetPanelSelection();

            if (ok)
            {
                ShowBalloon("已清除格式",
                    "Word 选区已恢复默认颜色/字号/字体，并解除锁定。",
                    ToolTipIcon.Info);
            }
            else
            {
                ShowBalloon("清除未完全成功",
                    "可能文档处于只读/受保护视图，或选区不支持修改。已解除锁定。",
                    ToolTipIcon.Warning);
            }
        }

        /// <summary>把 Word 选区格式恢复默认。返回 false 说明有项目设置失败。</summary>
        private bool DoClearWordFormat(object app)
        {
            bool ok = true;
            try
            {
                dynamic wordApp = app;
                dynamic sel = wordApp.Selection;

                if (!TrySetOk(() => sel.Font.Color = WdColorAutomatic)) ok = false;
                if (!TrySetOk(() => sel.Font.Name = "")) ok = false;
                if (!TrySetOk(() => sel.Font.NameAscii = "")) ok = false;
                if (!TrySetOk(() => sel.Font.NameFarEast = "")) ok = false;

                float def = 0f;
                try
                {
                    dynamic normal = wordApp.ActiveDocument.Styles("Normal");
                    def = Convert.ToSingle(normal.Font.Size);
                }
                catch (Exception)
                {
                }
                if (!TrySetOk(() => sel.Font.Size = def > 0 ? def : 10.5f)) ok = false;
            }
            catch (Exception)
            {
                InvalidateWordApp();
                ok = false;
            }
            return ok;
        }

        /// <summary>面板待选状态复位（不碰 Word）。</summary>
        private void ResetPanelSelection()
        {
            _selAutoColor = true;
            _selRgb = 0;
            _selectedSwatch = null;
            UpdateSwatchBorders();
            UpdateSelectionPreview();
            _cmbSize.SelectedIndex = 0;
            _cmbFont.Text = "";
            UpdateStatus();
        }

        /// <summary>
        /// 把锁定状态写到 Word 当前选区。每项单独 try，单项失败不影响其它项。
        /// </summary>
        private void ApplyLockToSelection(dynamic sel)
        {
            if (_lockColorAuto)
            {
                TrySet(() => sel.Font.Color = WdColorAutomatic);
            }
            else
            {
                TrySet(() => SetFontColor(sel, _lockColorRgb));
            }

            if (_lockSize > 0) TrySet(() => sel.Font.Size = _lockSize);
            if (_lockFontName != null)
            {
                TrySet(() => sel.Font.Name = _lockFontName);
                TrySet(() => sel.Font.NameAscii = _lockFontName);
                TrySet(() => sel.Font.NameFarEast = _lockFontName);
            }
        }

        // ================================================================
        //  锁定维持
        // ================================================================

        private void StartLockTimer()
        {
            _lockTimer.Interval = LockCheckIntervalMs;
            _lockTimer.Tick += OnLockTimerTick;
            _lockTimer.Start();
        }

        private void OnLockTimerTick(object sender, EventArgs e)
        {
            if (_busy || !_lockActive)
            {
                return;
            }

            _busy = true;
            try
            {
                EnforceLock();
            }
            finally
            {
                _busy = false;
            }
        }

        /// <summary>
        /// 每 300ms 对比 Word 当前选区与锁定值，不一致就整套刷回去，
        /// 这样接着打出来的字始终是锁定状态，而不是继承前一个字符的格式。
        /// </summary>
        private void EnforceLock()
        {
            object app = GetWordApp();
            if (app == null)
            {
                ClearLock();
                return;
            }

            // 连到的是无文档的后台残留实例：解除锁定并丢掉它，别每 300ms 白刷一次
            if (!WordHasDocuments(app))
            {
                InvalidateWordApp();
                ClearLock();
                return;
            }

            try
            {
                dynamic wordApp = app;
                dynamic sel = wordApp.Selection;
                int pos = GetInt(sel.Start, -1);

                if (_lockStartPos >= 0 && pos >= 0 && pos < _lockStartPos)
                {
                    ClearLock();
                    return;
                }

                bool needApply = false;

                long curColor = GetFontColorRaw(sel);
                if (_lockColorAuto)
                {
                    if (curColor != WdColorAutomatic) needApply = true;
                }
                else
                {
                    if (curColor != (long)_lockColorRgb) needApply = true;
                }

                if (_lockSize > 0 && Math.Abs(GetFloat(sel.Font.Size, -1f) - _lockSize) > 0.01f) needApply = true;
                if (_lockFontName != null)
                {
                    string curFont = GetString(sel.Font.Name);
                    if (curFont == null || curFont != _lockFontName) needApply = true;
                }

                if (needApply)
                {
                    ApplyLockToSelection(sel);
                }
            }
            catch (Exception)
            {
                InvalidateWordApp();
                ClearLock();
            }
        }

        private void ClearLock()
        {
            _lockActive = false;
            _lockStartPos = -1;
            if (_lblStatus != null && _lblStatus.IsHandleCreated)
            {
                UpdateStatus();
            }
        }

        private void OnUnlockClick(object sender, EventArgs e)
        {
            ClearLock();
            ShowBalloon("已解除锁定", "不再干预你在 Word 里的输入。", ToolTipIcon.Info);
        }

        // ================================================================
        //  预设（点列表项 = 载入到面板，再点「确定」应用）
        // ================================================================

        private void OnPresetSelectionChanged(object sender, EventArgs e)
        {
            LoadSelectedPreset();
        }

        private void LoadSelectedPreset()
        {
            int idx = _lstPresets.SelectedIndex;
            if (idx < 0 || idx >= _presets.Count || _presets[idx] == null)
            {
                return;
            }
            Preset p = _presets[idx];

            uint v;
            if (string.IsNullOrEmpty(p.HexColor) || !TryParseHexColor(p.HexColor, out v))
            {
                _selAutoColor = true;
                _selectedSwatch = null;
                UpdateSwatchBorders();
                UpdateSelectionPreview();
            }
            else
            {
                _selAutoColor = false;
                _selRgb = v;
                _selectedSwatch = null;
                UpdateSwatchBorders();
                UpdateSelectionPreview();
            }

            if (p.Size > 0)
            {
                string sz = p.Size.ToString("0.#");
                int at = _cmbSize.Items.IndexOf(sz);
                if (at >= 0)
                {
                    _cmbSize.SelectedIndex = at;
                }
                else
                {
                    _cmbSize.Items.Add(sz);
                    _cmbSize.SelectedIndex = _cmbSize.Items.Count - 1;
                }
            }
            else
            {
                _cmbSize.SelectedIndex = 0;
            }

            _cmbFont.Text = p.FontName == null ? "" : p.FontName;
        }

        private void OnSavePresetClick(object sender, EventArgs e)
        {
            string name = _txtPresetName.Text == null ? "" : _txtPresetName.Text.Trim();
            if (name.Length == 0)
            {
                MessageBox.Show(this, "先给这个预设起个名字。", "WordColorTool",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            Preset p = new Preset();
            p.Name = name.Replace("|", " ");
            p.HexColor = _selAutoColor ? "" : RgbToHex(_selRgb);
            p.Size = ReadSelectedSize();
            p.FontName = ReadSelectedFont();

            for (int i = 0; i < _presets.Count; i++)
            {
                if (_presets[i] != null && _presets[i].Name == p.Name)
                {
                    _presets[i] = p;
                    RefreshPresetList();
                    SavePresets();
                    _txtPresetName.Text = "";
                    return;
                }
            }

            _presets.Add(p);
            RefreshPresetList();
            SavePresets();
            _txtPresetName.Text = "";
        }

        private void OnDeletePresetClick(object sender, EventArgs e)
        {
            int idx = _lstPresets.SelectedIndex;
            if (idx < 0 || idx >= _presets.Count)
            {
                return;
            }
            _presets.RemoveAt(idx);
            RefreshPresetList();
            SavePresets();
        }

        private void RefreshPresetList()
        {
            if (_lstPresets == null)
            {
                return;
            }
            _lstPresets.BeginUpdate();
            _lstPresets.Items.Clear();
            for (int i = 0; i < _presets.Count; i++)
            {
                if (_presets[i] != null)
                {
                    _lstPresets.Items.Add(DescribePreset(_presets[i]));
                }
            }
            _lstPresets.EndUpdate();
        }

        private static string DescribePreset(Preset p)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(p.Name);
            sb.Append("：");
            sb.Append(string.IsNullOrEmpty(p.HexColor) ? "自动色" : p.HexColor);
            if (p.Size > 0) sb.Append(" " + p.Size.ToString("0.#") + "pt");
            if (p.FontName != null) sb.Append(" " + p.FontName);
            return sb.ToString();
        }

        /// <summary>打开面板时把当前锁定状态反映到控件上。</summary>
        private void RefreshPanelFromLock()
        {
            if (!_lockActive)
            {
                _selAutoColor = true;
                _selRgb = 0;
                _selectedSwatch = null;
                UpdateSwatchBorders();
                UpdateSelectionPreview();
                _cmbSize.SelectedIndex = 0;
                _cmbFont.Text = "";
                UpdateStatus();
                return;
            }

            _selAutoColor = _lockColorAuto;
            _selRgb = _lockColorRgb;
            _selectedSwatch = null;
            UpdateSwatchBorders();
            UpdateSelectionPreview();

            if (_lockSize > 0)
            {
                string sz = _lockSize.ToString("0.#");
                int at = _cmbSize.Items.IndexOf(sz);
                if (at < 0)
                {
                    _cmbSize.Items.Add(sz);
                    at = _cmbSize.Items.Count - 1;
                }
                _cmbSize.SelectedIndex = at;
            }
            else
            {
                _cmbSize.SelectedIndex = 0;
            }

            _cmbFont.Text = _lockFontName == null ? "" : _lockFontName;
            UpdateStatus();
        }

        private void UpdateStatus()
        {
            if (_lblStatus == null)
            {
                return;
            }
            _lblStatus.Text = _lockActive ? ("当前锁定：" + DescribeLockValues()) : "当前：未锁定";
        }

        private string DescribeLock()
        {
            if (!_lockActive)
            {
                return "未锁定";
            }
            return DescribeLockValues();
        }

        private string DescribeLockValues()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(_lockColorAuto ? "自动颜色" : ("#" + RgbToHex(_lockColorRgb)));
            if (_lockSize > 0) sb.Append(" " + _lockSize.ToString("0.#") + "pt");
            if (_lockFontName != null) sb.Append(" " + _lockFontName);
            return sb.ToString();
        }

        private void ShowBalloon(string title, string text, ToolTipIcon icon)
        {
            if (_trayIcon == null)
            {
                return;
            }
            try
            {
                _trayIcon.BalloonTipTitle = title;
                _trayIcon.BalloonTipText = text;
                _trayIcon.BalloonTipIcon = icon;
                _trayIcon.ShowBalloonTip(4000);
            }
            catch (Exception)
            {
            }
        }

        // ================================================================
        //  Word 互操作
        // ================================================================

        /// <summary>
        /// 拿到正在运行的 Word 实例。Word 没打开时返回 null，不崩溃、不弹窗。
        /// </summary>
        private object GetWordApp()
        {
            if (_wordApp != null)
            {
                return _wordApp;
            }

            try
            {
                // 只用 dynamic + Marshal.GetActiveObject，不引用 Word 的 COM 互操作程序集。
                _wordApp = Marshal.GetActiveObject("Word.Application");
            }
            catch (Exception)
            {
                _wordApp = null;
            }

            return _wordApp;
        }

        private void InvalidateWordApp()
        {
            _wordApp = null;
        }

        /// <summary>
        /// 这个 Word 实例是否打开了文档。返回 false 说明连上的是"后台残留实例"
        /// （无窗口、无文档，却占着 Word 的进程注册项），对它设色/读色都会失败。
        /// </summary>
        private static bool WordHasDocuments(object app)
        {
            try
            {
                dynamic wordApp = app;
                return Convert.ToInt32(wordApp.Documents.Count) > 0;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>列出当前有可见窗口的 Word 进程标题（告诉用户哪个才是他正在用的）。</summary>
        private static string GetVisibleWordWindows()
        {
            StringBuilder sb = new StringBuilder();
            try
            {
                System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName("WINWORD");
                int shown = 0;
                for (int i = 0; i < procs.Length; i++)
                {
                    string title = procs[i].MainWindowTitle;
                    if (title == null || title.Trim().Length == 0)
                    {
                        continue;   // 没窗口的就是后台残留实例
                    }
                    sb.Append("  · ");
                    sb.Append(title);
                    sb.AppendLine();
                    shown++;
                }
                if (shown == 0)
                {
                    sb.AppendLine("  （没找到有窗口的 Word 进程）");
                }
            }
            catch (Exception)
            {
            }
            return sb.ToString();
        }

        /// <summary>
        /// 6 位 RRGGBB 转成 Word 的字体颜色值（OLE BGR）。
        /// 换算：值 = R + (G &lt;&lt; 8) + (B &lt;&lt; 16)。例如 "E60000" → 0x0000E6。
        /// </summary>
        private static bool TryParseHexColor(string hex, out uint colorValue)
        {
            colorValue = 0;
            if (hex == null)
            {
                return false;
            }

            string s = hex.Trim();
            if (s.StartsWith("#"))
            {
                s = s.Substring(1);
            }
            if (s.Length != 6)
            {
                return false;
            }

            int r, g, b;
            if (!int.TryParse(s.Substring(0, 2), System.Globalization.NumberStyles.HexNumber, null, out r)) return false;
            if (!int.TryParse(s.Substring(2, 2), System.Globalization.NumberStyles.HexNumber, null, out g)) return false;
            if (!int.TryParse(s.Substring(4, 2), System.Globalization.NumberStyles.HexNumber, null, out b)) return false;

            colorValue = (uint)(r | (g << 8) | (b << 16));
            return true;
        }

        private static string RgbToHex(uint v)
        {
            return string.Format("{0:X2}{1:X2}{2:X2}",
                (int)(v & 0xFF), (int)((v >> 8) & 0xFF), (int)((v >> 16) & 0xFF));
        }

        /// <summary>
        /// 设置选区字体颜色。优先 Font.ColorRGB；某些环境下通过 ROT 拿到的 Font 对象
        /// 不暴露 ColorRGB 成员（实测 64 位 Word 被 32 位进程自动化时），
        /// 动态绑定抛异常时退回 Font.Color —— 它同样接受 OLE BGR 值。
        /// </summary>
        private static void SetFontColor(dynamic sel, uint value)
        {
            try
            {
                sel.Font.ColorRGB = value;
            }
            catch (Exception)
            {
                sel.Font.Color = value;
            }
        }

        private static long GetFontColorRaw(dynamic sel)
        {
            try
            {
                return Convert.ToInt64(sel.Font.ColorRGB);
            }
            catch (Exception)
            {
                try
                {
                    return Convert.ToInt64(sel.Font.Color);
                }
                catch (Exception)
                {
                    return long.MinValue;
                }
            }
        }

        private static int GetInt(dynamic value, int fallback)
        {
            try
            {
                return Convert.ToInt32(value);
            }
            catch (Exception)
            {
                return fallback;
            }
        }

        private static float GetFloat(dynamic value, float fallback)
        {
            try
            {
                return Convert.ToSingle(value);
            }
            catch (Exception)
            {
                return fallback;
            }
        }

        private static string GetString(dynamic value)
        {
            try
            {
                return Convert.ToString(value);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void TrySet(Action action)
        {
            try
            {
                action();
            }
            catch (Exception)
            {
            }
        }

        /// <summary>试试某项设置是否成功（供「清除格式」汇报用）。</summary>
        private static bool TrySetOk(Action action)
        {
            try
            {
                action();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        // ================================================================
        //  预设存取（%APPDATA%\WordColorTool\presets.txt）
        //  行格式：名字|色值|字号|字体名
        // ================================================================

        private static string PresetFilePath()
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WordColorTool");
            try
            {
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
            }
            catch (Exception)
            {
            }
            return Path.Combine(dir, "presets.txt");
        }

        private void SavePresets()
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < _presets.Count; i++)
                {
                    Preset p = _presets[i];
                    if (p == null)
                    {
                        continue;
                    }
                    sb.Append(p.Name == null ? "" : p.Name);
                    sb.Append('|');
                    sb.Append(p.HexColor == null ? "" : p.HexColor);
                    sb.Append('|');
                    sb.Append(p.Size.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture));
                    sb.Append('|');
                    sb.Append(p.FontName == null ? "" : p.FontName.Replace('|', ' '));
                    sb.AppendLine();
                }
                File.WriteAllText(PresetFilePath(), sb.ToString(), new UTF8Encoding(false));
            }
            catch (Exception)
            {
            }
        }

        private void LoadPresets()
        {
            _presets.Clear();
            string path = PresetFilePath();
            if (!File.Exists(path))
            {
                RefreshPresetList();
                return;
            }
            try
            {
                string[] lines = File.ReadAllLines(path, Encoding.UTF8);
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i];
                    if (line == null || line.Trim().Length == 0)
                    {
                        continue;
                    }
                    string[] f = line.Split('|');
                    if (f.Length < 4)
                    {
                        continue;
                    }
                    Preset p = new Preset();
                    p.Name = f[0];
                    p.HexColor = f[1];
                    float sz;
                    if (!float.TryParse(f[2], System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out sz))
                    {
                        sz = 0f;
                    }
                    p.Size = sz > 0 ? sz : 0f;
                    p.FontName = string.IsNullOrEmpty(f[3]) ? null : f[3];
                    _presets.Add(p);
                }
            }
            catch (Exception)
            {
            }
            RefreshPresetList();
        }

        // ================================================================
        //  退出
        // ================================================================

        private void OnExitClick(object sender, EventArgs e)
        {
            Cleanup();
            Application.Exit();
        }

        private void Cleanup()
        {
            if (_shuttingDown)
            {
                return;
            }
            _shuttingDown = true;

            try { _lockTimer.Stop(); } catch (Exception) { }
            _lockTimer.Tick -= OnLockTimerTick;
            _lockActive = false;

            if (_trayIcon != null)
            {
                _trayIcon.Visible = false;
                _trayIcon.Dispose();
                _trayIcon = null;
            }
            if (_trayMenu != null)
            {
                _trayMenu.Dispose();
                _trayMenu = null;
            }

            _wordApp = null;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                Cleanup();
                _lockTimer.Dispose();
            }
            base.Dispose(disposing);
        }

        private static Label MakeLabel(string text, int x, int y, int w, int h, Font font)
        {
            Label l = new Label();
            l.Text = text;
            l.Location = new Point(x, y);
            l.Size = new Size(w, h);
            if (font != null)
            {
                l.Font = font;
            }
            return l;
        }

        private static Button MakeButton(string text, int x, int y, int w, int h, EventHandler click)
        {
            Button b = new Button();
            b.Text = text;
            b.Location = new Point(x, y);
            b.Size = new Size(w, h);
            b.Click += click;
            return b;
        }

        private static Color RgbFromHex(string hex, Color fallback)
        {
            uint v;
            if (hex != null && TryParseHexColor(hex, out v))
            {
                return Color.FromArgb((int)(v & 0xFF), (int)((v >> 8) & 0xFF), (int)((v >> 16) & 0xFF));
            }
            return fallback;
        }
    }
}
