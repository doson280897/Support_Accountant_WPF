using OfficeOpenXml;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Xml;
using MessageBox = System.Windows.MessageBox;

namespace Support_Accountant
{
    public partial class MainWindow : Window
    {
        #region Variables
        private DispatcherTimer curTimeTimer;
        private int errorCount = 0;
        private int warningCount = 0;
        private bool isInitialLoad = true;

        private readonly RenameInvoiceController renameController;
        private readonly ExtractInvoiceController extractController;
        #endregion

        #region MainWindow
        public MainWindow()
        {
            InitializeComponent();
            InitializeTimer();
            InitializeLogging();

            curTimeTimer = null;
            XmlFiles = Array.Empty<string>();
            PdfFiles = Array.Empty<string>();
            ExtractXmlFiles = Array.Empty<string>();
            FolderDestination = string.Empty;
            LastExcelFilePath = string.Empty;
            IsProcessing = false;
            CancellationTokenSource = null;

            renameController = new RenameInvoiceController(this);
            extractController = new ExtractInvoiceController(this);

            ExcelPackage.License.SetNonCommercialPersonal("SonDNT");
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            AutoResizeWindow();
            isInitialLoad = false;
            LogInfo("Window auto-sized to fit content");
        }

        private void InitializeLogging()
        {
            LogInfo("Application started successfully");
            LogInfo("Ready to process invoice files");
        }

        private void mainProgressBar_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            mainProgressBar.ToolTip = "Indicates the progress of the current operation.";
        }

        private void TabControl_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            ResetAllData();

            if (!isInitialLoad)
            {
                _ = Dispatcher.BeginInvoke(new Action(() =>
                {
                    AnimateWindowResize();
                    LogInfo($"Window resized for tab: {GetCurrentTabName()}");
                }), DispatcherPriority.ContextIdle);
            }
        }

        private void AnimateWindowResize()
        {
            try
            {
                UpdateLayout();
                Size desiredSize = CalculateDesiredSize();
                double targetWidth = Math.Max(MinWidth, Math.Min(MaxWidth, desiredSize.Width));
                double targetHeight = Math.Max(MinHeight, Math.Min(MaxHeight, desiredSize.Height));

                System.Windows.Media.Animation.DoubleAnimation widthAnimation = new System.Windows.Media.Animation.DoubleAnimation
                {
                    From = Width,
                    To = targetWidth,
                    Duration = TimeSpan.FromMilliseconds(250),
                    EasingFunction = new System.Windows.Media.Animation.QuadraticEase { EasingMode = System.Windows.Media.Animation.EasingMode.EaseInOut }
                };

                System.Windows.Media.Animation.DoubleAnimation heightAnimation = new System.Windows.Media.Animation.DoubleAnimation
                {
                    From = Height,
                    To = targetHeight,
                    Duration = TimeSpan.FromMilliseconds(250),
                    EasingFunction = new System.Windows.Media.Animation.QuadraticEase { EasingMode = System.Windows.Media.Animation.EasingMode.EaseInOut }
                };

                if (WindowStartupLocation == WindowStartupLocation.CenterScreen)
                {
                    double targetLeft = (SystemParameters.PrimaryScreenWidth - targetWidth) / 2;
                    double targetTop = (SystemParameters.PrimaryScreenHeight - targetHeight) / 2;

                    System.Windows.Media.Animation.DoubleAnimation leftAnimation = new System.Windows.Media.Animation.DoubleAnimation
                    {
                        From = Left,
                        To = targetLeft,
                        Duration = TimeSpan.FromMilliseconds(250),
                        EasingFunction = new System.Windows.Media.Animation.QuadraticEase { EasingMode = System.Windows.Media.Animation.EasingMode.EaseInOut }
                    };

                    System.Windows.Media.Animation.DoubleAnimation topAnimation = new System.Windows.Media.Animation.DoubleAnimation
                    {
                        From = Top,
                        To = targetTop,
                        Duration = TimeSpan.FromMilliseconds(250),
                        EasingFunction = new System.Windows.Media.Animation.QuadraticEase { EasingMode = System.Windows.Media.Animation.EasingMode.EaseInOut }
                    };

                    BeginAnimation(LeftProperty, leftAnimation);
                    BeginAnimation(TopProperty, topAnimation);
                }

                BeginAnimation(WidthProperty, widthAnimation);
                BeginAnimation(HeightProperty, heightAnimation);
            }
            catch (Exception ex)
            {
                LogError($"Error during animated window resize: {ex.Message}");
                AutoResizeWindow();
            }
        }

        private void AutoResizeWindow()
        {
            try
            {
                UpdateLayout();
                Size desiredSize = CalculateDesiredSize();
                Width = Math.Max(MinWidth, Math.Min(MaxWidth, desiredSize.Width));
                Height = Math.Max(MinHeight, Math.Min(MaxHeight, desiredSize.Height));

                if (WindowStartupLocation == WindowStartupLocation.CenterScreen)
                {
                    Left = (SystemParameters.PrimaryScreenWidth - Width) / 2;
                    Top = (SystemParameters.PrimaryScreenHeight - Height) / 2;
                }
            }
            catch (Exception ex)
            {
                LogError($"Error during window auto-resize: {ex.Message}");
            }
        }

        private Size CalculateDesiredSize()
        {
            TabItem selectedTab = mainTabControl.SelectedItem as TabItem;
            if (selectedTab?.Content is FrameworkElement tabContent)
            {
                tabContent.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                Size contentSize = tabContent.DesiredSize;
                double totalWidth = contentSize.Width + 50;
                double totalHeight = contentSize.Height + 120;
                return new Size(totalWidth, totalHeight);
            }

            return new Size(800, 600);
        }

        private string GetCurrentTabName()
        {
            TabItem selectedTab = mainTabControl.SelectedItem as TabItem;
            return selectedTab?.Header?.ToString() ?? "Unknown";
        }

        private void ResetAllData()
        {
            if (Dispatcher.CheckAccess())
            {
                label_TotalFiles_RenameInvoice.Content = "Files found: 0 XML / 0 PDF";
                label_TotalFiles_ExtractInvoice.Content = "Files found: 0 XML";
                mainProgressBar.Value = 0;
                progressText.Text = "";
                errorCount = 0;
                warningCount = 0;
                UpdateLogCounters();
            }
            else
            {
                Dispatcher.Invoke(() =>
                {
                    label_TotalFiles_RenameInvoice.Content = "Files found: 0 XML / 0 PDF";
                    label_TotalFiles_ExtractInvoice.Content = "Files found: 0 XML";
                    mainProgressBar.Value = 0;
                    progressText.Text = "";
                    errorCount = 0;
                    warningCount = 0;
                    UpdateLogCounters();
                });
            }
        }
        #endregion

        #region Public Properties and Methods for Controllers
        public string[] XmlFiles { get; set; }
        public string[] PdfFiles { get; set; }
        public string[] ExtractXmlFiles { get; set; }
        public string FolderDestination { get; set; }
        public string LastExcelFilePath { get; set; }
        public bool IsProcessing { get; set; }
        public CancellationTokenSource CancellationTokenSource { get; set; }
        public DateTime ProcessStartTime { get; set; }

        public void SetErrorCount(int count)
        {
            errorCount = count;
        }

        public void SetWarningCount(int count)
        {
            warningCount = count;
        }

        public void UpdateProgressBar(double value, double maximum, string text = "")
        {
            if (Dispatcher.CheckAccess())
            {
                mainProgressBar.Value = value;
                mainProgressBar.Maximum = maximum;
                progressText.Text = text;
            }
            else
            {
                Dispatcher.Invoke(() =>
                {
                    mainProgressBar.Value = value;
                    mainProgressBar.Maximum = maximum;
                    progressText.Text = text;
                });
            }
        }

        public void PopulateComboBox(System.Windows.Controls.ComboBox comboBox, string[] files)
        {
            comboBox.Items.Clear();
            foreach (string filePath in files)
            {
                _ = comboBox.Items.Add(Path.GetFileName(filePath));
            }

            if (comboBox.Items.Count > 0)
            {
                comboBox.SelectedIndex = 0;
                LogInfo($"Populated file list with {comboBox.Items.Count} files");
            }
        }

        public void ShowXMLContent(string filePath, string fileName)
        {
            try
            {
                LogInfo($"Opening XML file: {fileName}");

                string xmlContent = File.ReadAllText(filePath);
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(xmlContent);

                using (StringWriter stringWriter = new StringWriter())
                using (XmlTextWriter xmlTextWriter = new XmlTextWriter(stringWriter))
                {
                    xmlTextWriter.Formatting = Formatting.Indented;
                    doc.WriteContentTo(xmlTextWriter);
                    xmlTextWriter.Flush();
                    xmlContent = stringWriter.GetStringBuilder().ToString();
                }

                Window xmlViewer = new Window
                {
                    Title = $"XML Viewer - {fileName}",
                    Width = 800,
                    Height = 600,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Owner = this
                };

                System.Windows.Controls.TextBox textBoxXml = new System.Windows.Controls.TextBox
                {
                    Text = xmlContent,
                    FontFamily = new System.Windows.Media.FontFamily("Consolas"),
                    FontSize = 12,
                    IsReadOnly = true,
                    AcceptsReturn = true,
                    AcceptsTab = true,
                    VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto,
                    HorizontalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
                };

                xmlViewer.Content = textBoxXml;
                _ = xmlViewer.ShowDialog();

                LogSuccess($"Successfully opened XML file: {fileName}");
            }
            catch (Exception ex)
            {
                LogError($"Error reading XML file '{fileName}': {ex.Message}");
                ShowError($"Error reading or formatting XML file: {ex.Message}");
            }
        }

        public void OpenFileWithDefaultProgram(string filePath)
        {
            try
            {
                LogInfo($"Opening file with default program: {Path.GetFileName(filePath)}");
                _ = Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                LogSuccess($"Successfully opened file: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                LogError($"Error opening file '{Path.GetFileName(filePath)}': {ex.Message}");
                ShowError($"Error opening file: {ex.Message}");
            }
        }

        public void ShowError(string message)
        {
            _ = Dispatcher.CheckAccess()
                ? MessageBox.Show(this, message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                : Dispatcher.Invoke(() => MessageBox.Show(this, message, "Error", MessageBoxButton.OK, MessageBoxImage.Error));
        }

        public void ShowInfo(string message, string title = "Information")
        {
            _ = Dispatcher.CheckAccess()
                ? MessageBox.Show(this, message, title, MessageBoxButton.OK, MessageBoxImage.Information)
                : Dispatcher.Invoke(() => MessageBox.Show(this, message, title, MessageBoxButton.OK, MessageBoxImage.Information));
        }

        public string FormatTimeSpan(TimeSpan timeSpan)
        {
            return timeSpan.TotalHours >= 1
                ? $"{timeSpan.Hours:D2}h {timeSpan.Minutes:D2}m {timeSpan.Seconds:D2}s"
                : timeSpan.TotalMinutes >= 1
                    ? $"{timeSpan.Minutes:D2}m {timeSpan.Seconds:D2}s"
                    : $"{timeSpan.Seconds:D2}.{timeSpan.Milliseconds:D3}s";
        }
        #endregion

        #region Event Handlers - Delegate to Controllers
        private void btn_Browse_RenameInvoice_Click(object sender, RoutedEventArgs e)
        {
            renameController.Browse_Click();
        }

        private void btn_Openfile_RenameInvoice_Click(object sender, RoutedEventArgs e)
        {
            renameController.OpenFile_Click();
        }

        private async void btn_Renameall_RenameInvoice_Click(object sender, RoutedEventArgs e)
        {
            await renameController.RenameAll_Click();
        }

        private void btn_Stop_RenameInvoice_Click(object sender, RoutedEventArgs e)
        {
            renameController.Stop_Click();
        }

        private void btn_OpenFolder_RenameInvoice_Click(object sender, RoutedEventArgs e)
        {
            renameController.OpenFolder_Click();
        }

        private void btn_Browse_ExtractInvoice_Click(object sender, RoutedEventArgs e)
        {
            extractController.Browse_Click();
        }

        private void btn_Openfile_ExtractInvoice_Click(object sender, RoutedEventArgs e)
        {
            extractController.OpenFile_Click();
        }

        private async void btn_ExportSummary_ExtractInvoice_Click(object sender, RoutedEventArgs e)
        {
            await extractController.ExportSummary_Click();
        }

        private void btn_Stop_ExtractInvoice_Click(object sender, RoutedEventArgs e)
        {
            extractController.Stop_Click();
        }

        private void btn_OpenExcel_ExtractInvoice_Click(object sender, RoutedEventArgs e)
        {
            extractController.OpenExcel_Click();
        }
        #endregion

        #region Logging System
        public enum LogLevel
        {
            Info,
            Warning,
            Error,
            Success
        }

        private void LogMessage(LogLevel level, string message)
        {
            if (Dispatcher.CheckAccess())
            {
                WriteLogMessage(level, message, GetCurrentTabName());
            }
            else
            {
                Dispatcher.Invoke(() => WriteLogMessage(level, message, GetCurrentTabName()));
            }
        }

        private void WriteLogMessage(LogLevel level, string message, string tabName)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string prefix;

            switch (level)
            {
                case LogLevel.Info:
                    prefix = "[INFO]";
                    break;
                case LogLevel.Warning:
                    prefix = "[WARN]";
                    break;
                case LogLevel.Error:
                    prefix = "[ERROR]";
                    break;
                case LogLevel.Success:
                    prefix = "[SUCCESS]";
                    break;
                default:
                    prefix = "[INFO]";
                    break;
            }

            string FormatColumn(string text, int width)
            {
                return text.Length > width ? text.Substring(0, width - 1) + "…" : text.PadRight(width);
            }
            string logEntry = $"{FormatColumn(timestamp, 10)}{FormatColumn(tabName, 20)}{FormatColumn(prefix, 10)}{message}";

            if (!string.IsNullOrEmpty(txtLog_Shared.Text))
            {
                txtLog_Shared.Text += Environment.NewLine;
            }
            txtLog_Shared.Text += logEntry;

            if (level == LogLevel.Error)
            {
                errorCount++;
            }
            else if (level == LogLevel.Warning)
            {
                warningCount++;
            }

            UpdateLogCounters();

            if (chk_AutoScroll_Shared.IsChecked == true)
            {
                logScrollViewer.ScrollToEnd();
            }

            Debug.WriteLine(logEntry);
        }

        public void UpdateLogCounters(string tabName = "")
        {
            txtErrorCount_Shared.Text = errorCount.ToString();
            txtWarningCount_Shared.Text = warningCount.ToString();
            txtStatus_Shared.Text = IsProcessing ? "Processing..." : "Ready";
            txtStatus_Shared.Foreground = IsProcessing ?
                new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Orange) :
                new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Green);
        }

        public void LogInfo(string message)
        {
            LogMessage(LogLevel.Info, message);
        }

        public void LogWarning(string message)
        {
            LogMessage(LogLevel.Warning, message);
        }

        public void LogError(string message)
        {
            LogMessage(LogLevel.Error, message);
        }

        public void LogSuccess(string message)
        {
            LogMessage(LogLevel.Success, message);
        }

        private void btn_ClearLog_Click(object sender, RoutedEventArgs e)
        {
            txtLog_Shared.Clear();
            errorCount = 0;
            warningCount = 0;
            UpdateLogCounters();
            LogInfo("Log cleared by user");
        }

        private void btn_ExportLog_Shared_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
                    FileName = "Invoice_Management_log.txt"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    File.WriteAllText(saveFileDialog.FileName, txtLog_Shared.Text);
                    LogSuccess($"Log exported successfully to: {saveFileDialog.FileName}");
                }
            }
            catch (Exception ex)
            {
                LogError($"Error exporting log: {ex.Message}");
                ShowError($"Error exporting log: {ex.Message}");
            }
        }

        private void btn_ClearLog_Shared_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            btn_ClearLog_Shared.ToolTip = "Clear the shared log display area.";
        }

        private void btn_ExportLog_Shared_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            btn_ExportLog_Shared.ToolTip = "Export the shared log to a text file.";
        }
        #endregion

        #region Timer Management
        private void InitializeTimer()
        {
            curTimeTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(200)
            };
            curTimeTimer.Tick += update_GUI;
            curTimeTimer.Start();
        }

        private void update_GUI(object sender, EventArgs e)
        {
            label_CurTime.Content = DateTime.Now.ToString("ddd dd/MM/yyyy - HH:mm:ss", CultureInfo.InvariantCulture);
        }
        #endregion

        #region Menu Management
        private void Help_About_Click(object sender, RoutedEventArgs e)
        {
            LogInfo("About dialog opened");
            _ = MessageBox.Show(this, "Support Accountant v1.0\nDeveloped by SonDNT\n© 2025", "About Support Accountant", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void File_Exit_Click(object sender, RoutedEventArgs e)
        {
            LogInfo("Application exit requested");
            System.Windows.Application.Current.Shutdown();
        }
        #endregion

        #region ToolTips
        private void btn_OpenFolder_RenameInvoice_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            btn_OpenFolder_RenameInvoice.ToolTip = "Open the folder containing renamed invoices.";
        }

        private void btn_Renameall_RenameInvoice_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            btn_Renameall_RenameInvoice.ToolTip = "Start renaming all invoices in the selected folder.";
        }

        private void btn_Stop_RenameInvoice_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            btn_Stop_RenameInvoice.ToolTip = "Stop the current rename operation.";
        }

        private void btn_Openfile_RenameInvoice_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            btn_Openfile_RenameInvoice.ToolTip = "Preview the selected invoice file.";
        }

        private void ComboBox_RenameInvoice_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            ComboBox_RenameInvoice.ToolTip = "Select invoice file to open or rename.";
        }

        private void txtBox_Browse_RenameInvoice_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            txtBox_Browse_RenameInvoice.ToolTip = "Path to the folder containing XML/PDF invoices.";
        }

        private void btn_Browse_RenameInvoice_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            btn_Browse_RenameInvoice.ToolTip = "Browse to select the folder containing XML/PDF invoices.";
        }

        private void txtBox_Browse_ExtractInvoice_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            txtBox_Browse_ExtractInvoice.ToolTip = "Path to the folder containing XML invoices for extraction.";
        }

        private void btn_Browse_ExtractInvoice_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            btn_Browse_ExtractInvoice.ToolTip = "Browse to select the folder containing XML invoices for extraction.";
        }

        private void ComboBox_ExtractInvoice_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            ComboBox_ExtractInvoice.ToolTip = "Select XML invoice file to preview.";
        }

        private void btn_Openfile_ExtractInvoice_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            btn_Openfile_ExtractInvoice.ToolTip = "Preview the selected XML invoice file.";
        }

        private void btn_ExportSummary_ExtractInvoice_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            btn_ExportSummary_ExtractInvoice.ToolTip = "Export invoice summary to Excel with selected fields.";
        }

        private void btn_Stop_ExtractInvoice_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            btn_Stop_ExtractInvoice.ToolTip = "Stop the current extraction operation.";
        }

        private void btn_OpenExcel_ExtractInvoice_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            btn_OpenExcel_ExtractInvoice.ToolTip = "Open the last exported Excel file.";
        }

        private void btn_ExportLog_ExtractInvoice_Click(object sender, RoutedEventArgs e)
        {
            btn_ExportLog_Shared_Click(sender, e);
        }
        #endregion
    }
}