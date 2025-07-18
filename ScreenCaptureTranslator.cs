using System;
using System.Drawing;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Tesseract;
using System.Text.Json;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using System.Security.Cryptography;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace ScreenCaptureTranslator
{
    public partial class MainForm : Form
    {
        private Point startPoint;
        private Rectangle selectionRectangle;
        private bool isSelecting;
        private Form selectionForm;
        private PictureBox previewBox;
        private Button confirmButton;
        private TextBox resultLabel;
        private Bitmap capturedImage;
        private static readonly HttpClient client = new HttpClient();
        private Panel previewPanel;
        private Panel buttonPanel;
        private Panel resultPanel;
        private ToolStripMenuItem startupMenuItem;
        private MenuStrip menuStrip;
        private BaiduConfig baiduConfig;
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;
        private const int HOTKEY_ID = 1;
        private const int WM_HOTKEY = 0x0312;

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        public MainForm()
        {
            LoadConfig();
            InitializeComponents();
            InitializeSystemTray();
            RegisterHotKey();
        }

        private void RegisterHotKey()
        {
            const uint MOD_ALT = 0x0001; // Alt
            const uint VK_Q = 0x51;      // Q键
            RegisterHotKey(this.Handle, HOTKEY_ID, MOD_ALT, VK_Q);
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                StartCaptureButton_Click(this, EventArgs.Empty);
            }
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            RegisterHotKey();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.WindowState = FormWindowState.Minimized;
                this.ShowInTaskbar = false;
                this.Visible = false;
            }
            else
            {
                trayIcon.Visible = false;
                trayIcon.Dispose();
                UnregisterHotKey(this.Handle, HOTKEY_ID);
            }
            base.OnFormClosing(e);
        }

        private void InitializeSystemTray()
        {
            trayIcon = new NotifyIcon
            {
                Text = "屏幕截图翻译器",
                Visible = true
            };

            try
            {
                string iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tray.ico");
                if (File.Exists(iconPath))
                {
                    trayIcon.Icon = new Icon(iconPath);
                }
                else
                {
                    trayIcon.Icon = SystemIcons.Application;
                    MessageBox.Show($"未找到托盘图标: {iconPath}\n请确保 tray.ico 存在于程序目录，并设置为“始终复制”。\n当前使用默认图标。",
                                    "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                trayIcon.Icon = SystemIcons.Application;
                MessageBox.Show($"加载托盘图标失败: {ex.Message}\n请检查 tray.ico 文件格式。\n当前使用默认图标。",
                                "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("进入翻译界面", null, (s, e) => ShowMainWindow());
            trayMenu.Items.Add("退出", null, (s, e) => Application.Exit());

            trayIcon.ContextMenuStrip = trayMenu;
            trayIcon.MouseClick += (s, e) =>
            {
                if (e.Button == MouseButtons.Left)
                {
                    ShowMainWindow();
                }
            };

            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Visible = false;
        }

        private void ShowMainWindow()
        {
            this.Visible = true;
            this.ShowInTaskbar = true;
            this.WindowState = FormWindowState.Normal;
            SetForegroundWindow(this.Handle);
            ShowWindow(this.Handle, 5);
            this.Activate();
        }

        private void LoadConfig()
        {
            try
            {
                string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
                bool createdNewConfig = false;

                if (!File.Exists(configPath))
                {
                    string templateJson = @"{
    ""BaiduTranslate"": {
        ""AppId"": ""YOUR_APP_ID"",
        ""SecretKey"": ""YOUR_SECRET_KEY""
    }
}";
                    File.WriteAllText(configPath, templateJson);
                    createdNewConfig = true;
                }
                else
                {
                    string json = File.ReadAllText(configPath);
                    try
                    {
                        using (var doc = JsonDocument.Parse(json))
                        {
                        }
                    }
                    catch (JsonException)
                    {
                        string templateJson = @"{
    ""BaiduTranslate"": {
        ""AppId"": ""YOUR_APP_ID"",
        ""SecretKey"": ""YOUR_SECRET_KEY""
    }
}";
                        File.WriteAllText(configPath, templateJson);
                        createdNewConfig = true;
                    }
                }

                string jsonContent = File.ReadAllText(configPath);
                using (var doc = JsonDocument.Parse(jsonContent))
                {
                    var root = doc.RootElement;
                    if (root.TryGetProperty("BaiduTranslate", out var baidu))
                    {
                        baiduConfig = new BaiduConfig
                        {
                            AppId = baidu.GetProperty("AppId").GetString(),
                            SecretKey = baidu.GetProperty("SecretKey").GetString()
                        };
                        if (string.IsNullOrEmpty(baiduConfig.AppId) || baiduConfig.AppId == "YOUR_APP_ID" ||
                            string.IsNullOrEmpty(baiduConfig.SecretKey) || baiduConfig.SecretKey == "YOUR_SECRET_KEY")
                        {
                            string message = createdNewConfig
                                ? $"已在 {configPath} 创建配置文件 config.json。\n请在文件中填写有效的 AppId 和 SecretKey。\n1. 访问 https://fanyi-api.baidu.com 注册并获取凭据。\n2. 编辑 config.json 替换 YOUR_APP_ID 和 YOUR_SECRET_KEY。\n3. 重新启动程序。"
                                : $"config.json 中的 AppId 或 SecretKey 无效，请编辑 {configPath} 并填写有效的凭据。\n1. 访问 https://fanyi-api.baidu.com 注册并获取凭据。\n2. 替换 YOUR_APP_ID 和 YOUR_SECRET_KEY。\n3. 重新启动程序。";
                            MessageBox.Show(message, "配置提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            throw new Exception("无效的配置文件");
                            Environment.Exit(1);
                        }
                    }
                    else
                    {
                        MessageBox.Show($"config.json 中未找到 BaiduTranslate 配置，请检查 {configPath} 格式！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        throw new Exception("无效的配置文件格式");
                        Environment.Exit(1);
                    }
                }

                if (createdNewConfig)
                {
                    MessageBox.Show($"已在 {configPath} 创建配置文件 config.json。\n请在文件中填写有效的 AppId 和 SecretKey。\n1. 访问 https://fanyi-api.baidu.com 注册并获取凭据。\n2. 编辑 config.json 替换 YOUR_APP_ID 和 YOUR_SECRET_KEY。\n3. 重新启动程序。", "配置提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    throw new Exception("已创建配置文件，请填写后重试");
                    Environment.Exit(1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载配置文件失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
                Environment.Exit(1);
            }
        }

        private void InitializeComponents()
        {
            this.Text = "屏幕截图翻译器";
            this.Size = new Size(1200, 640);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.BackColor = Color.WhiteSmoke;
            this.Font = new Font("微软雅黑", 12);

            try
            {
                string iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tray.ico");
                if (File.Exists(iconPath))
                {
                    this.Icon = new Icon(iconPath);
                }
                else
                {
                    this.Icon = SystemIcons.Application;
                    MessageBox.Show($"未找到窗体图标: {iconPath}\n请确保 tray.ico 存在于程序目录，并设置为“始终复制”。\n当前使用默认图标。",
                                    "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                this.Icon = SystemIcons.Application;
                MessageBox.Show($"加载窗体图标失败: {ex.Message}\n请检查 tray.ico 文件格式。\n当前使用默认图标。",
                                "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            menuStrip = new MenuStrip
            {
                BackColor = Color.WhiteSmoke,
                Font = new Font("微软雅黑", 10),
                Location = new Point(0, 0)
            };
            ToolStripMenuItem settingsMenu = new ToolStripMenuItem("设置");
            startupMenuItem = new ToolStripMenuItem("开机启动")
            {
                CheckOnClick = true,
                Checked = IsStartupEnabled()
            };
            startupMenuItem.Click += StartupMenuItem_Click;
            settingsMenu.DropDownItems.Add(startupMenuItem);

            ToolStripMenuItem helpMenu = new ToolStripMenuItem("使用帮助");
            helpMenu.Click += (s, e) =>
            {
                using (var helpForm = new HelpForm())
                {
                    helpForm.ShowDialog();
                }
            };

            menuStrip.Items.Add(settingsMenu);
            menuStrip.Items.Add(helpMenu);
            this.Controls.Add(menuStrip);

            Label titleLabel = new Label
            {
                Font = new Font("微软雅黑", 30, FontStyle.Bold),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.FromArgb(44, 62, 80),
                BackColor = Color.Transparent,
                Size = new Size(900, 60),
                Location = new Point((this.ClientSize.Width - 900) / 2, 30 + menuStrip.Height)
            };
            this.Controls.Add(titleLabel);

            previewPanel = new Panel
            {
                Location = new Point((this.ClientSize.Width - 1100) / 2, 110 + menuStrip.Height),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(8),
                AutoSize = true,
                Visible = false
            };
            previewBox = new PictureBox
            {
                Location = new Point(8, 8),
                BorderStyle = BorderStyle.None,
                SizeMode = PictureBoxSizeMode.Normal,
                BackColor = Color.Gainsboro
            };
            previewPanel.Controls.Add(previewBox);
            this.Controls.Add(previewPanel);

            buttonPanel = new Panel
            {
                Size = new Size(700, 60),
                BackColor = Color.Transparent
            };

            Button startCaptureButton = new Button
            {
                Text = "开始截图",
                Size = new Size(220, 44),
                Font = new Font("微软雅黑", 12, FontStyle.Bold),
                BackColor = Color.FromArgb(52, 152, 219),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Location = new Point(60, 8),
                Cursor = Cursors.Hand
            };
            startCaptureButton.FlatAppearance.BorderSize = 0;
            startCaptureButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(41, 128, 185);
            startCaptureButton.Click += StartCaptureButton_Click;

            confirmButton = new Button
            {
                Text = "确认并翻译",
                Size = new Size(220, 44),
                Font = new Font("微软雅黑", 12, FontStyle.Bold),
                BackColor = Color.FromArgb(176, 176, 176),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Location = new Point(420, 8),
                Cursor = Cursors.Hand,
                Enabled = false
            };
            confirmButton.FlatAppearance.BorderSize = 0;
            confirmButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(39, 174, 96);
            confirmButton.FlatAppearance.MouseDownBackColor = Color.FromArgb(52, 152, 219);
            confirmButton.Click += ConfirmButton_Click;

            confirmButton.EnabledChanged += (s, e) =>
            {
                if (confirmButton.Enabled)
                    confirmButton.BackColor = Color.FromArgb(52, 152, 219);
                else
                    confirmButton.BackColor = Color.FromArgb(220, 220, 220);
            };
            confirmButton.MouseDown += (s, e) =>
            {
                if (confirmButton.Enabled)
                    confirmButton.BackColor = Color.FromArgb(52, 152, 219);
            };
            confirmButton.MouseUp += (s, e) =>
            {
                if (confirmButton.Enabled)
                    confirmButton.BackColor = Color.FromArgb(176, 176, 176);
            };

            buttonPanel.Controls.Add(startCaptureButton);
            buttonPanel.Controls.Add(confirmButton);
            this.Controls.Add(buttonPanel);

            resultPanel = new Panel
            {
                Size = new Size(1100, 320),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(10)
            };
            resultLabel = new TextBox
            {
                Size = new Size(1080, 300),
                Multiline = true,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("微软雅黑", 14, FontStyle.Regular),
                BackColor = Color.White,
                ForeColor = Color.FromArgb(52, 73, 94)
            };
            resultPanel.Controls.Add(resultLabel);
            this.Controls.Add(resultPanel);

            buttonPanel.Location = new Point((this.ClientSize.Width - 700) / 2, 110 + menuStrip.Height);
            resultPanel.Location = new Point((this.ClientSize.Width - 1100) / 2, buttonPanel.Top + buttonPanel.Height + 20);
        }

        private bool IsStartupEnabled()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", false))
            {
                if (key != null)
                {
                    string value = key.GetValue(Application.ProductName) as string;
                    return value != null && value == Application.ExecutablePath;
                }
                return false;
            }
        }

        private void StartupMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (startupMenuItem.Checked)
                    {
                        key.SetValue(Application.ProductName, Application.ExecutablePath);
                        MessageBox.Show("已启用开机启动", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        key.DeleteValue(Application.ProductName, false);
                        MessageBox.Show("已禁用开机启动", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置开机启动失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void StartCaptureButton_Click(object sender, EventArgs e)
        {
            previewBox.Image = null;
            previewPanel.Visible = false;
            this.Visible = false;
            StartScreenCapture();
        }

        private async void StartScreenCapture()
        {
            await Task.Delay(100);
            selectionForm = new Form
            {
                FormBorderStyle = FormBorderStyle.None,
                Opacity = 0.3,
                BackColor = Color.Black,
                Bounds = Screen.PrimaryScreen.Bounds,
                Location = new Point(0, 0),
                Cursor = Cursors.Cross,
                TopMost = true,
                KeyPreview = true
            };

            isSelecting = false;
            selectionRectangle = Rectangle.Empty;

            selectionForm.MouseDown += SelectionForm_MouseDown;
            selectionForm.MouseMove += SelectionForm_MouseMove;
            selectionForm.MouseUp += SelectionForm_MouseUp;
            selectionForm.Paint += SelectionForm_Paint;
            selectionForm.Shown += (s, e) =>
            {
                SetForegroundWindow(selectionForm.Handle);
                ShowWindow(selectionForm.Handle, 5);
                selectionForm.Focus();

            };
            selectionForm.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                {
                    selectionForm.Close();
                    ShowMainWindow();
                    previewBox.Image = null;
                    previewPanel.Visible = false;
                    confirmButton.Enabled = false;
                }
            };

            selectionForm.Show();
        }

        private void SelectionForm_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isSelecting = true;
                startPoint = e.Location;
                selectionRectangle = new Rectangle(e.Location, new Size(0, 0));
            }
        }

        private void SelectionForm_MouseMove(object sender, MouseEventArgs e)
        {
            if (isSelecting)
            {
                selectionRectangle = new Rectangle(
                    Math.Min(startPoint.X, e.X),
                    Math.Min(startPoint.Y, e.Y),
                    Math.Abs(e.X - startPoint.X),
                    Math.Abs(e.Y - startPoint.Y));
                selectionForm.Invalidate();
            }
        }

        private void SelectionForm_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isSelecting = false;
                if (selectionRectangle.Width > 0 && selectionRectangle.Height > 0)
                {
                    CaptureScreenRegion();
                    selectionForm.Close();
                    ShowMainWindow();
                    confirmButton.Enabled = true;
                    previewBox.Image = capturedImage;
                    ExtractAndDisplayText();
                }
                else
                {
                    selectionForm.Close();
                    ShowMainWindow();
                    previewBox.Image = null;
                    previewPanel.Visible = false;
                    confirmButton.Enabled = false;
                }
            }
        }

        private void SelectionForm_Paint(object sender, PaintEventArgs e)
        {
            using (Pen pen = new Pen(Color.Red, 2))
            {
                e.Graphics.DrawRectangle(pen, selectionRectangle);
            }
        }

        private void CaptureScreenRegion()
        {
            capturedImage = new Bitmap(selectionRectangle.Width, selectionRectangle.Height);
            using (Graphics g = Graphics.FromImage(capturedImage))
            {
                g.CopyFromScreen(selectionRectangle.Location, Point.Empty, selectionRectangle.Size);
            }

            previewBox.Size = new Size(capturedImage.Width, capturedImage.Height);
            previewPanel.Size = new Size(capturedImage.Width + 16, capturedImage.Height + 16);
            previewPanel.Visible = true;

            resultPanel.Size = new Size(Math.Max(400, capturedImage.Width + 16), resultPanel.Height);
            resultLabel.Size = new Size(Math.Max(380, capturedImage.Width), resultLabel.Height);

            int previewTop = 110 + menuStrip.Height;
            previewPanel.Location = new Point((this.ClientSize.Width - previewPanel.Width) / 2, previewTop);
            buttonPanel.Location = new Point((this.ClientSize.Width - buttonPanel.Width) / 2, previewTop + previewPanel.Height + 30);
            resultPanel.Location = new Point((this.ClientSize.Width - resultPanel.Width) / 2, buttonPanel.Top + buttonPanel.Height + 30);

            int contentWidth = Math.Max(600, Math.Max(previewPanel.Width + 100, resultPanel.Width + 100));
            int contentHeight = resultPanel.Top + resultPanel.Height + 50;

            int maxWidth = Screen.PrimaryScreen.WorkingArea.Width;
            int maxHeight = Screen.PrimaryScreen.WorkingArea.Height;
            contentWidth = Math.Min(contentWidth, maxWidth);
            contentHeight = Math.Min(contentHeight, maxHeight);

            this.ClientSize = new Size(contentWidth, contentHeight);

            previewPanel.Location = new Point((this.ClientSize.Width - previewPanel.Width) / 2, previewTop);
            buttonPanel.Location = new Point((this.ClientSize.Width - buttonPanel.Width) / 2, previewTop + previewPanel.Height + 30);
            resultPanel.Location = new Point((this.ClientSize.Width - resultPanel.Width) / 2, buttonPanel.Top + buttonPanel.Height + 30);
        }

        private async void ExtractAndDisplayText()
        {
            int textLines;
            int newHeight;
            int contentHeight;
            int previewTop = 110 + menuStrip.Height;

            try
            {
                string extractedText = await PerformOCR(capturedImage);
                extractedText = extractedText.Normalize();
                resultLabel.Text = $"识别到的内容:\r\n{extractedText}";

                textLines = resultLabel.Text.Split('\n').Length;
                newHeight = Math.Max(100, textLines * 25 + 20);
                resultPanel.Size = new Size(resultPanel.Width, newHeight);
                resultLabel.Size = new Size(resultLabel.Width, newHeight - 20);

                contentHeight = resultPanel.Top + resultPanel.Height + 50;
                this.ClientSize = new Size(this.ClientSize.Width, Math.Min(contentHeight, Screen.PrimaryScreen.WorkingArea.Height));

                previewPanel.Location = new Point((this.ClientSize.Width - previewPanel.Width) / 2, previewTop);
                buttonPanel.Location = new Point((this.ClientSize.Width - buttonPanel.Width) / 2, previewTop + previewPanel.Height + 30);
                resultPanel.Location = new Point((this.ClientSize.Width - resultPanel.Width) / 2, buttonPanel.Top + buttonPanel.Height + 30);
            }
            catch (Exception ex)
            {
                resultLabel.Text = $"提取失败:\r\n{ex.Message}";
                textLines = resultLabel.Text.Split('\n').Length;
                newHeight = Math.Max(100, textLines * 25 + 20);
                resultPanel.Size = new Size(resultPanel.Width, newHeight);
                resultLabel.Size = new Size(resultLabel.Width, newHeight - 20);

                contentHeight = resultPanel.Top + resultPanel.Height + 50;
                this.ClientSize = new Size(this.ClientSize.Width, Math.Min(contentHeight, Screen.PrimaryScreen.WorkingArea.Height));

                previewPanel.Location = new Point((this.ClientSize.Width - previewPanel.Width) / 2, previewTop);
                buttonPanel.Location = new Point((this.ClientSize.Width - buttonPanel.Width) / 2, previewTop + previewPanel.Height + 30);
                resultPanel.Location = new Point((this.ClientSize.Width - resultPanel.Width) / 2, buttonPanel.Top + buttonPanel.Height + 30);
            }
        }

        public static string ConvertIpaToSimple(string ipa)
        {
            if (string.IsNullOrWhiteSpace(ipa)) return ipa;

            return ipa
                .Replace("ɹ", "r")
                .Replace("p̚", "p")
                .Replace("ʔ", "");
        }

        public static string ExtractEnglishLetters(string input)
        {
            return Regex.Replace(input, @"[^A-Za-z0-9 ]", " ").Replace("  ", " ").Trim();
        }

        private async Task<string> GetPhoneticFromDictionaryApi(string text)
        {
            var words = text.Split(new[] { ' ', '\t', '\r', '\n', '.', ',', '!', '?' }, StringSplitOptions.RemoveEmptyEntries);
            var results = new List<string>();

            foreach (var word in words)
            {
                string ipa = await GetSingleWordPhonetic(word);
                string simple = ConvertIpaToSimple(ipa);
                results.Add($"{word}: {simple}");
            }
            return string.Join("\r\n", results);
        }

        private async Task<string> GetSingleWordPhonetic(string word)
        {
            try
            {
                string url = $"https://api.dictionaryapi.dev/api/v2/entries/en/{Uri.EscapeDataString(word)}";
                var response = await client.GetAsync(url);
                if (!response.IsSuccessStatusCode)
                    return "(无音标)";

                var json = await response.Content.ReadAsStringAsync();
                using (var doc = JsonDocument.Parse(json))
                {
                    var root = doc.RootElement;
                    if (root.ValueKind == JsonValueKind.Array && root.GetArrayLength() > 0)
                    {
                        var entry = root[0];
                        if (entry.TryGetProperty("phonetics", out var phoneticsArr) && phoneticsArr.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var phonetic in phoneticsArr.EnumerateArray())
                            {
                                if (phonetic.TryGetProperty("text", out var textProp))
                                {
                                    string ipa = textProp.GetString();
                                    if (!string.IsNullOrEmpty(ipa) && ipa.Contains("/"))
                                        return ipa;
                                }
                            }
                            foreach (var phonetic in phoneticsArr.EnumerateArray())
                            {
                                if (phonetic.TryGetProperty("text", out var textProp))
                                {
                                    string ipa = textProp.GetString();
                                    if (!string.IsNullOrEmpty(ipa))
                                        return ipa;
                                }
                            }
                        }
                    }
                }
                return "(无音标)";
            }
            catch
            {
                return "(无音标)";
            }
        }

        private async void ConfirmButton_Click(object sender, EventArgs e)
        {
            int textLines;
            int newHeight;
            int contentHeight;
            int previewTop = 110 + menuStrip.Height;

            confirmButton.Enabled = false;
            confirmButton.BackColor = Color.FromArgb(220, 220, 220);
            resultLabel.Text = "处理中...";

            try
            {
                string extractedText = await PerformOCR(capturedImage);
                extractedText = extractedText.Normalize();
                string filteredText = ExtractEnglishLetters(extractedText);
                if (string.IsNullOrWhiteSpace(filteredText))
                {
                    resultLabel.Text = "未检测到有效英文文本。";
                    textLines = resultLabel.Text.Split('\n').Length;
                    newHeight = Math.Max(100, textLines * 25 + 20);
                    resultPanel.Size = new Size(resultPanel.Width, newHeight);
                    resultLabel.Size = new Size(resultLabel.Width, newHeight - 20);

                    contentHeight = resultPanel.Top + resultPanel.Height + 50;
                    this.ClientSize = new Size(this.ClientSize.Width, Math.Min(contentHeight, Screen.PrimaryScreen.WorkingArea.Height));

                    previewPanel.Location = new Point((this.ClientSize.Width - previewPanel.Width) / 2, previewTop);
                    buttonPanel.Location = new Point((this.ClientSize.Width - buttonPanel.Width) / 2, previewTop + previewPanel.Height + 30);
                    resultPanel.Location = new Point((this.ClientSize.Width - resultPanel.Width) / 2, buttonPanel.Top + buttonPanel.Height + 30);
                    return;
                }
                string translatedText = await TranslateWithBaidu(filteredText, "en", "zh");
                string phonetic = await GetPhoneticFromDictionaryApi(filteredText);

                resultLabel.Text = $"原文:\r\n{extractedText}\r\n\r\n翻译:\r\n{translatedText}\r\n\r\n音标:\r\n{phonetic}";

                textLines = resultLabel.Text.Split('\n').Length;
                newHeight = Math.Max(100, textLines * 25 + 20);
                resultPanel.Size = new Size(resultPanel.Width, newHeight);
                resultLabel.Size = new Size(resultLabel.Width, newHeight - 20);

                contentHeight = resultPanel.Top + resultPanel.Height + 50;
                this.ClientSize = new Size(this.ClientSize.Width, Math.Min(contentHeight, Screen.PrimaryScreen.WorkingArea.Height));

                previewPanel.Location = new Point((this.ClientSize.Width - previewPanel.Width) / 2, previewTop);
                buttonPanel.Location = new Point((this.ClientSize.Width - buttonPanel.Width) / 2, previewTop + previewPanel.Height + 30);
                resultPanel.Location = new Point((this.ClientSize.Width - resultPanel.Width) / 2, buttonPanel.Top + buttonPanel.Height + 30);
            }
            catch (Exception ex)
            {
                resultLabel.Text = $"错误:\r\n{ex.Message}\r\n可能原因：\r\n1. 网络连接问题\r\n2. 百度翻译 API 配额耗尽，请登录 https://fanyi-api.baidu.com 检查配额\r\n3. AppId 或 SecretKey 错误，请检查 config.json";
                textLines = resultLabel.Text.Split('\n').Length;
                newHeight = Math.Max(100, textLines * 25 + 20);
                resultPanel.Size = new Size(resultPanel.Width, newHeight);
                resultLabel.Size = new Size(resultLabel.Width, newHeight - 20);

                contentHeight = resultPanel.Top + resultPanel.Height + 50;
                this.ClientSize = new Size(this.ClientSize.Width, Math.Min(contentHeight, Screen.PrimaryScreen.WorkingArea.Height));

                previewPanel.Location = new Point((this.ClientSize.Width - previewPanel.Width) / 2, previewTop);
                buttonPanel.Location = new Point((this.ClientSize.Width - buttonPanel.Width) / 2, previewTop + previewPanel.Height + 30);
                resultPanel.Location = new Point((this.ClientSize.Width - resultPanel.Width) / 2, buttonPanel.Top + buttonPanel.Height + 30);
            }
            finally
            {
                confirmButton.Enabled = true;
                confirmButton.BackColor = Color.FromArgb(176, 176, 176);
            }
        }

        private async Task<string> PerformOCR(Bitmap image)
        {
            try
            {
                using (var engine = new TesseractEngine(@".\lib\Tesseract-OCR\tessdata", "eng", EngineMode.Default))
                {
                    using (var img = PixConverter.ToPix(image))
                    {
                        using (var page = engine.Process(img))
                        {
                            return page.GetText().Trim();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"OCR 失败: {ex.Message}");
            }
        }

        private async Task<string> TranslateWithBaidu(string text, string sourceLang, string targetLang)
        {
            if (string.IsNullOrEmpty(text))
            {
                throw new Exception("翻译失败: 输入文本为空。");
            }

            try
            {
                string salt = DateTime.Now.Ticks.ToString();
                string sign = GenerateBaiduSign(baiduConfig.AppId, text, salt, baiduConfig.SecretKey);
                string url = $"https://fanyi-api.baidu.com/api/trans/vip/translate?q={Uri.EscapeDataString(text)}&from={sourceLang}&to={targetLang}&appid={baiduConfig.AppId}&salt={salt}&sign={sign}";
                CopyableMessageBox.Show($"翻译请求URL: {url}");
                var response = await client.GetAsync(url);
                var responseBody = await response.Content.ReadAsStringAsync();
                CopyableMessageBox.Show($"翻译响应状态码: {response.StatusCode}\n响应体: {responseBody}");

                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception($"翻译失败: 服务器返回状态码 {response.StatusCode}");
                }

                using (var doc = JsonDocument.Parse(responseBody))
                {
                    var root = doc.RootElement;
                    if (root.TryGetProperty("error_code", out var errorCode))
                    {
                        string errorMsg = root.GetProperty("error_msg").GetString();
                        throw new Exception($"翻译失败: 错误码 {errorCode}, 错误信息 {errorMsg}\r\n可能原因：\r\n1. AppId 或 SecretKey 错误\r\n2. API 配额耗尽，请登录 https://fanyi-api.baidu.com 检查");
                    }

                    if (root.TryGetProperty("trans_result", out var transResult) && transResult.ValueKind == JsonValueKind.Array)
                    {
                        var translations = new List<string>();
                        foreach (var item in transResult.EnumerateArray())
                        {
                            if (item.TryGetProperty("dst", out var dst))
                            {
                                translations.Add(dst.GetString());
                            }
                        }
                        string translatedText = string.Join("", translations);
                        if (string.IsNullOrEmpty(translatedText))
                        {
                            throw new Exception("翻译失败: 翻译结果为空。");
                        }
                        return translatedText;
                    }
                    throw new Exception("翻译失败: 响应格式无效。");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"翻译失败: {ex.Message}");
            }
        }

        private string GenerateBaiduSign(string appId, string query, string salt, string secretKey)
        {
            string input = appId + query + salt + secretKey;
            using (var md5 = MD5.Create())
            {
                byte[] inputBytes = Encoding.UTF8.GetBytes(input);
                byte[] hashBytes = md5.ComputeHash(inputBytes);
                StringBuilder sb = new StringBuilder();
                foreach (byte b in hashBytes)
                {
                    sb.Append(b.ToString("x2"));
                }
                return sb.ToString();
            }
        }

        private string GeneratePhoneticTranscription(string text)
        {
            return $"/{text.Replace(" ", "/")}/";
        }

        private class BaiduConfig
        {
            public string AppId { get; set; }
            public string SecretKey { get; set; }
        }
    }

    public class HelpForm : Form
    {
        public HelpForm()
        {
            this.Text = "使用帮助";
            this.Size = new Size(800, 180);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = Color.WhiteSmoke;
            this.Font = new Font("微软雅黑", 12);

            TextBox helpText = new TextBox
            {
                Text = "按快捷键Alt+Q可截图，Esc键可退出截图状态",
                Multiline = true,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                Size = new Size(700, 100),
                Location = new Point(30, 30),
                Font = new Font("微软雅黑", 12, FontStyle.Regular),
                ForeColor = Color.FromArgb(52, 73, 94),
                BackColor = Color.WhiteSmoke,
                TextAlign = HorizontalAlignment.Center
            };
            try
            {
                string iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tray.ico");
                if (File.Exists(iconPath))
                {
                    this.Icon = new Icon(iconPath);
                }
                else
                {
                    this.Icon = SystemIcons.Application;
                    MessageBox.Show($"未找到窗体图标: {iconPath}\n请确保 tray.ico 存在于程序目录，并设置为“始终复制”。\n当前使用默认图标。",
                                    "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                this.Icon = SystemIcons.Application;
                MessageBox.Show($"加载窗体图标失败: {ex.Message}\n请检查 tray.ico 文件格式。\n当前使用默认图标。",
                                "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            this.Controls.Add(helpText);
        }
    }

    static class Program
    {
        [STAThread]
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        static void Main()
        {
            if (Environment.OSVersion.Version.Major >= 6) SetProcessDPIAware();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    public static class CopyableMessageBox
    {
        public static void Show(string text, string title = "调试窗口，关掉继续")
        {
            Form form = new Form
            {
                Text = title,
                Width = 600,
                Height = 400,
                StartPosition = FormStartPosition.CenterParent
            };

            TextBox textBox = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                Dock = DockStyle.Fill,
                ScrollBars = ScrollBars.Both,
                Text = text
            };

            form.Controls.Add(textBox);
            form.ShowDialog();
        }
    }

    public class FlexibleStringConverter : JsonConverter<string>
    {
        public override string Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            if (reader.TokenType == JsonTokenType.String)
            {
                return reader.GetString();
            }
            else if (reader.TokenType == JsonTokenType.Number)
            {
                if (reader.TryGetInt64(out long l))
                    return l.ToString();
                else if (reader.TryGetDouble(out double d))
                    return d.ToString();
                else
                    return null;
            }
            else if (reader.TokenType == JsonTokenType.Null)
            {
                return null;
            }
            else
            {
                return null;
            }
        }
        public override void Write(Utf8JsonWriter writer, string value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value);
        }
    }
}