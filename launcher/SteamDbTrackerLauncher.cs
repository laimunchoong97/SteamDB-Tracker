using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SteamDbTrackerLauncher
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            if (args.Length > 0 && args[0] == "--self-test")
            {
                Environment.Exit(SelfTest(baseDirectory));
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new TrackerForm(baseDirectory));
        }

        private static int SelfTest(string baseDirectory)
        {
            string[] required = {
                Path.Combine(baseDirectory, "runtime", "node.exe"),
                Path.Combine(baseDirectory, "steamdb_tracker.js"),
                Path.Combine(baseDirectory, "node_modules"),
                Path.Combine(baseDirectory, ".env")
            };

            foreach (string item in required)
            {
                if (!File.Exists(item) && !Directory.Exists(item)) return 2;
            }
            return 0;
        }
    }

    internal sealed class TrackerForm : Form
    {
        private readonly string baseDirectory;
        private readonly string envFile;
        private readonly ComboBox languageBox;
        private readonly NumericUpDown thresholdBox;
        private readonly Button runButton;
        private readonly Button rebuildButton;
        private readonly Button openEnglishButton;
        private readonly Button openChineseButton;
        private readonly Button openFolderButton;
        private readonly TextBox logBox;
        private readonly Label statusLabel;

        public TrackerForm(string baseDirectory)
        {
            this.baseDirectory = baseDirectory;
            envFile = Path.Combine(baseDirectory, ".env");

            Text = "SteamDB Tracker / SteamDB 追踪器";
            Width = 820;
            Height = 600;
            MinimumSize = new Size(720, 520);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei UI", 9F);

            Label heading = new Label();
            heading.Text = "SteamDB Release Tracker / SteamDB 发售追踪器";
            heading.Font = new Font("Microsoft YaHei UI", 17F, FontStyle.Bold);
            heading.ForeColor = Color.FromArgb(23, 50, 77);
            heading.AutoSize = true;
            heading.Location = new Point(24, 20);
            Controls.Add(heading);

            Label languageLabel = new Label();
            languageLabel.Text = "Workbook language / 工作簿语言";
            languageLabel.AutoSize = true;
            languageLabel.Location = new Point(27, 74);
            Controls.Add(languageLabel);

            languageBox = new ComboBox();
            languageBox.DropDownStyle = ComboBoxStyle.DropDownList;
            languageBox.Items.AddRange(new object[] {
                "English",
                "简体中文",
                "English + 简体中文"
            });
            languageBox.Width = 210;
            languageBox.Location = new Point(250, 70);
            Controls.Add(languageBox);

            Label thresholdLabel = new Label();
            thresholdLabel.Text = "Follower threshold / 关注人数门槛";
            thresholdLabel.AutoSize = true;
            thresholdLabel.Location = new Point(27, 111);
            Controls.Add(thresholdLabel);

            thresholdBox = new NumericUpDown();
            thresholdBox.Minimum = 1000;
            thresholdBox.Maximum = 10000000;
            thresholdBox.Increment = 500;
            thresholdBox.ThousandsSeparator = true;
            thresholdBox.Width = 210;
            thresholdBox.Location = new Point(250, 107);
            Controls.Add(thresholdBox);

            runButton = CreateButton("Run Tracker / 运行抓取", 28, 154, 215);
            runButton.BackColor = Color.FromArgb(15, 118, 110);
            runButton.ForeColor = Color.White;
            runButton.Click += async delegate { await RunTrackerAsync(); };
            Controls.Add(runButton);

            rebuildButton = CreateButton("Rebuild Excel / 重建 Excel", 253, 154, 215);
            rebuildButton.Click += async delegate { await RebuildTrackerAsync(); };
            Controls.Add(rebuildButton);

            openFolderButton = CreateButton("Open Folder / 打开文件夹", 478, 154, 215);
            openFolderButton.Click += delegate { OpenPath(baseDirectory); };
            Controls.Add(openFolderButton);

            openEnglishButton = CreateButton("Open English Excel", 28, 200, 215);
            openEnglishButton.Click += delegate {
                OpenWorkbook(Path.Combine(baseDirectory, "steamdb_upcoming_tracker.xlsx"));
            };
            Controls.Add(openEnglishButton);

            openChineseButton = CreateButton("打开中文 Excel", 253, 200, 215);
            openChineseButton.Click += delegate {
                OpenWorkbook(Path.Combine(baseDirectory, "steamdb_upcoming_tracker_zh-CN.xlsx"));
            };
            Controls.Add(openChineseButton);

            statusLabel = new Label();
            statusLabel.Text = "Ready / 就绪";
            statusLabel.AutoSize = true;
            statusLabel.ForeColor = Color.FromArgb(100, 116, 139);
            statusLabel.Location = new Point(28, 252);
            Controls.Add(statusLabel);

            logBox = new TextBox();
            logBox.Multiline = true;
            logBox.ReadOnly = true;
            logBox.ScrollBars = ScrollBars.Vertical;
            logBox.Font = new Font("Consolas", 9F);
            logBox.BackColor = Color.FromArgb(248, 250, 252);
            logBox.Location = new Point(28, 280);
            logBox.Size = new Size(745, 250);
            logBox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(logBox);

            LoadSettings();
            AppendLog("Ready. Choose settings, then click Run Tracker.");
            AppendLog("就绪。选择设置后，点击“运行抓取”。");
        }

        private static Button CreateButton(string text, int left, int top, int width)
        {
            Button button = new Button();
            button.Text = text;
            button.Width = width;
            button.Height = 34;
            button.Location = new Point(left, top);
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderColor = Color.FromArgb(215, 224, 232);
            return button;
        }

        private void LoadSettings()
        {
            string language = "both";
            decimal threshold = 1000;
            if (File.Exists(envFile))
            {
                foreach (string line in File.ReadAllLines(envFile))
                {
                    if (line.StartsWith("OUTPUT_LANGUAGE=", StringComparison.OrdinalIgnoreCase))
                    {
                        language = line.Substring(line.IndexOf('=') + 1).Trim();
                    }
                    if (line.StartsWith("FOCUSED_THRESHOLD=", StringComparison.OrdinalIgnoreCase))
                    {
                        decimal parsed;
                        if (decimal.TryParse(line.Substring(line.IndexOf('=') + 1).Trim(), out parsed))
                        {
                            threshold = Math.Max(1000, Math.Min(10000000, parsed));
                        }
                    }
                }
            }

            languageBox.SelectedIndex = language.Equals("en", StringComparison.OrdinalIgnoreCase)
                ? 0
                : language.StartsWith("zh", StringComparison.OrdinalIgnoreCase) ? 1 : 2;
            thresholdBox.Value = threshold;
        }

        private void SaveSettings()
        {
            string language = languageBox.SelectedIndex == 0
                ? "en"
                : languageBox.SelectedIndex == 1 ? "zh-CN" : "both";
            string[] source = File.Exists(envFile) ? File.ReadAllLines(envFile) : new string[0];
            bool languageWritten = false;
            bool thresholdWritten = false;
            StringBuilder builder = new StringBuilder();

            foreach (string original in source)
            {
                string line = original;
                if (line.StartsWith("OUTPUT_LANGUAGE=", StringComparison.OrdinalIgnoreCase))
                {
                    line = "OUTPUT_LANGUAGE=" + language;
                    languageWritten = true;
                }
                else if (line.StartsWith("FOCUSED_THRESHOLD=", StringComparison.OrdinalIgnoreCase))
                {
                    line = "FOCUSED_THRESHOLD=" + Decimal.ToInt32(thresholdBox.Value);
                    thresholdWritten = true;
                }
                builder.AppendLine(line);
            }

            if (!thresholdWritten) builder.AppendLine("FOCUSED_THRESHOLD=" + Decimal.ToInt32(thresholdBox.Value));
            if (!languageWritten) builder.AppendLine("OUTPUT_LANGUAGE=" + language);
            File.WriteAllText(envFile, builder.ToString(), new UTF8Encoding(false));
        }

        private async Task RunTrackerAsync()
        {
            if (!EnsureOutputFilesAvailable()) return;
            SaveSettings();
            string edge = FindEdge();
            if (edge == null)
            {
                MessageBox.Show(
                    "Microsoft Edge was not found.\n\n未找到 Microsoft Edge。",
                    Text,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            string profile = Path.Combine(baseDirectory, "edge_profile");
            ProcessStartInfo edgeStart = new ProcessStartInfo();
            edgeStart.FileName = edge;
            edgeStart.Arguments = "--remote-debugging-port=9222 --user-data-dir=\"" + profile
                + "\" \"https://steamdb.info/upcoming/?sort=followers_desc\"";
            edgeStart.UseShellExecute = true;
            Process.Start(edgeStart);

            DialogResult result = MessageBox.Show(
                "Complete the Cloudflare verification and wait for the SteamDB table to appear.\n"
                + "Then click OK to start.\n\n"
                + "请完成 Cloudflare 验证并等待 SteamDB 表格出现，然后点击“确定”开始。",
                "Browser verification / 浏览器验证",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Information
            );
            if (result != DialogResult.OK) return;
            await RunNodeAsync("--remote", "Crawling / 正在抓取");
        }

        private async Task RebuildTrackerAsync()
        {
            if (!EnsureOutputFilesAvailable()) return;
            await RunNodeAsync("--rebuild", "Rebuilding / 正在重建");
        }

        private bool EnsureOutputFilesAvailable()
        {
            string[] outputFiles = languageBox.SelectedIndex == 0
                ? new string[] { "steamdb_upcoming_tracker.xlsx" }
                : languageBox.SelectedIndex == 1
                    ? new string[] { "steamdb_upcoming_tracker_zh-CN.xlsx" }
                    : new string[] {
                        "steamdb_upcoming_tracker.xlsx",
                        "steamdb_upcoming_tracker_zh-CN.xlsx"
                    };

            foreach (string fileName in outputFiles)
            {
                string filePath = Path.Combine(baseDirectory, fileName);
                if (!File.Exists(filePath)) continue;

                try
                {
                    using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    {
                    }
                }
                catch (IOException)
                {
                    AppendLog("Workbook is open: " + fileName);
                    MessageBox.Show(
                        "Close " + fileName + " in Excel before running or rebuilding the tracker.\n\n"
                        + "请先在 Excel 中关闭 " + fileName + "，然后再运行抓取或重建。",
                        Text,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return false;
                }
                catch (UnauthorizedAccessException)
                {
                    AppendLog("Workbook cannot be written: " + fileName);
                    MessageBox.Show(
                        "The tracker cannot write to " + fileName + ". Close Excel and check file permissions.\n\n"
                        + "追踪器无法写入 " + fileName + "。请关闭 Excel 并检查文件权限。",
                        Text,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return false;
                }
            }

            return true;
        }

        private async Task RunNodeAsync(string arguments, string runningStatus)
        {
            SaveSettings();
            string node = Path.Combine(baseDirectory, "runtime", "node.exe");
            string script = Path.Combine(baseDirectory, "steamdb_tracker.js");
            if (!File.Exists(node) || !File.Exists(script))
            {
                MessageBox.Show(
                    "The packaged runtime is incomplete. Please reinstall or extract the full package.\n\n"
                    + "运行环境不完整，请重新安装或完整解压安装包。",
                    Text,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            SetBusy(true, runningStatus);
            AppendLog("----------------------------------------");
            try
            {
                int exitCode = await Task.Run<int>(delegate { return ExecuteNode(node, script, arguments); });
                if (exitCode == 0)
                {
                    statusLabel.Text = "Completed / 已完成";
                    AppendLog("Completed successfully / 已成功完成");
                    MessageBox.Show(
                        "Tracker completed successfully.\n\n追踪器已成功完成。",
                        Text,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
                else
                {
                    statusLabel.Text = "Failed / 失败";
                    AppendLog("Process failed with exit code " + exitCode);
                    MessageBox.Show(
                        "Tracker failed. Review the log for details.\n\n追踪失败，请查看日志。",
                        Text,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
            catch (Exception error)
            {
                statusLabel.Text = "Failed / 失败";
                AppendLog(error.ToString());
                MessageBox.Show(error.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(false, statusLabel.Text);
            }
        }

        private int ExecuteNode(string node, string script, string arguments)
        {
            ProcessStartInfo start = new ProcessStartInfo();
            start.FileName = node;
            start.Arguments = "\"" + script + "\" " + arguments;
            start.WorkingDirectory = baseDirectory;
            start.UseShellExecute = false;
            start.CreateNoWindow = true;
            start.RedirectStandardOutput = true;
            start.RedirectStandardError = true;
            start.StandardOutputEncoding = Encoding.UTF8;
            start.StandardErrorEncoding = Encoding.UTF8;

            using (Process process = new Process())
            {
                process.StartInfo = start;
                process.OutputDataReceived += delegate(object sender, DataReceivedEventArgs eventArgs) {
                    if (eventArgs.Data != null) AppendLog(eventArgs.Data);
                };
                process.ErrorDataReceived += delegate(object sender, DataReceivedEventArgs eventArgs) {
                    if (eventArgs.Data != null) AppendLog("ERROR: " + eventArgs.Data);
                };
                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
                process.WaitForExit();
                return process.ExitCode;
            }
        }

        private void AppendLog(string message)
        {
            if (logBox.InvokeRequired)
            {
                logBox.BeginInvoke(new Action<string>(AppendLog), message);
                return;
            }
            logBox.AppendText(message + Environment.NewLine);
            if (message.StartsWith("[Metadata "))
            {
                int end = message.IndexOf(']');
                if (end > 1) statusLabel.Text = message.Substring(1, end - 1);
            }
        }

        private void SetBusy(bool busy, string status)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<bool, string>(SetBusy), busy, status);
                return;
            }
            runButton.Enabled = !busy;
            rebuildButton.Enabled = !busy;
            openEnglishButton.Enabled = !busy;
            openChineseButton.Enabled = !busy;
            languageBox.Enabled = !busy;
            thresholdBox.Enabled = !busy;
            statusLabel.Text = status;
        }

        private static string FindEdge()
        {
            string[] candidates = {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft", "Edge", "Application", "msedge.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft", "Edge", "Application", "msedge.exe")
            };
            foreach (string candidate in candidates)
            {
                if (File.Exists(candidate)) return candidate;
            }
            return null;
        }

        private void OpenWorkbook(string file)
        {
            if (!File.Exists(file))
            {
                MessageBox.Show(
                    "Workbook does not exist yet. Run or rebuild the tracker first.\n\n"
                    + "工作簿尚未生成，请先运行抓取或重建。",
                    Text,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }
            OpenPath(file);
        }

        private static void OpenPath(string target)
        {
            ProcessStartInfo start = new ProcessStartInfo();
            start.FileName = target;
            start.UseShellExecute = true;
            Process.Start(start);
        }
    }
}
