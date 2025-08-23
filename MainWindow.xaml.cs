using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Threading;
using System.Xml;
using MessageBox = System.Windows.MessageBox;
using WpfCheckBox = System.Windows.Controls.CheckBox;

namespace Support_Accountant
{
    public partial class MainWindow : Window
    {
        #region Variables
        private DispatcherTimer curTimeTimer;
        private string[] xmlFiles;
        private string[] pdfFiles;
        private string folderDestination;
        private bool isProcessing;
        private int errorCount = 0;
        private int warningCount = 0;
        private string[] extractXmlFiles;
        private string lastExcelFilePath;
        private bool isInitialLoad = true;
        private CancellationTokenSource cancellationTokenSource;
        #endregion

        #region MainWindow
        public MainWindow()
        {
            InitializeComponent();
            InitializeTimer();
            InitializeLogging();

            curTimeTimer = null;
            xmlFiles = Array.Empty<string>();
            pdfFiles = Array.Empty<string>();
            extractXmlFiles = Array.Empty<string>();
            folderDestination = string.Empty;
            lastExcelFilePath = string.Empty;
            isProcessing = false;
            cancellationTokenSource = null;

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
                Dispatcher.BeginInvoke(new Action(() =>
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

                var widthAnimation = new System.Windows.Media.Animation.DoubleAnimation
                {
                    From = Width,
                    To = targetWidth,
                    Duration = TimeSpan.FromMilliseconds(250),
                    EasingFunction = new System.Windows.Media.Animation.QuadraticEase { EasingMode = System.Windows.Media.Animation.EasingMode.EaseInOut }
                };

                var heightAnimation = new System.Windows.Media.Animation.DoubleAnimation
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

                    var leftAnimation = new System.Windows.Media.Animation.DoubleAnimation
                    {
                        From = Left,
                        To = targetLeft,
                        Duration = TimeSpan.FromMilliseconds(250),
                        EasingFunction = new System.Windows.Media.Animation.QuadraticEase { EasingMode = System.Windows.Media.Animation.EasingMode.EaseInOut }
                    };

                    var topAnimation = new System.Windows.Media.Animation.DoubleAnimation
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

        private void UpdateProgressBar(double value, double maximum, string text = "")
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
        #endregion

        #region Logging System
        public enum LogLevel
        {
            Info,
            Warning,
            Error,
            Success
        }

        private void LogMessage(LogLevel level, string message, string tabName = "RenameInvoice")
        {
            if (Dispatcher.CheckAccess())
            {
                WriteLogMessage(level, message, tabName);
            }
            else
            {
                Dispatcher.Invoke(() => WriteLogMessage(level, message, tabName));
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

            string logEntry = $"{timestamp} {prefix} {message}";
            System.Windows.Controls.TextBox logTextBox = tabName == "ExtractInvoice" ? txtLog_ExtractInvoice : txtLog_RenameInvoice;
            ScrollViewer scrollViewer = tabName == "ExtractInvoice" ? logScrollViewer_Extract : logScrollViewer;
            WpfCheckBox autoScrollCheckBox = tabName == "ExtractInvoice" ? chk_AutoScroll_ExtractInvoice : chk_AutoScroll_RenameInvoice;

            if (!string.IsNullOrEmpty(logTextBox.Text))
            {
                logTextBox.Text += Environment.NewLine;
            }
            logTextBox.Text += logEntry;

            if (level == LogLevel.Error)
            {
                errorCount++;
            }
            else if (level == LogLevel.Warning)
            {
                warningCount++;
            }

            UpdateLogCounters(tabName);

            if (autoScrollCheckBox.IsChecked == true)
            {
                scrollViewer.ScrollToEnd();
            }

            Debug.WriteLine(logEntry);
        }

        private void UpdateLogCounters(string tabName = "RenameInvoice")
        {
            if (tabName == "ExtractInvoice")
            {
                txtErrorCount_ExtractInvoice.Text = errorCount.ToString();
                txtWarningCount_ExtractInvoice.Text = warningCount.ToString();
                txtStatus_ExtractInvoice.Text = isProcessing ? "Processing..." : "Ready";
                txtStatus_ExtractInvoice.Foreground = isProcessing ?
                    new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Orange) :
                    new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Green);
            }
            else
            {
                txtErrorCount_RenameInvoice.Text = errorCount.ToString();
                txtWarningCount_RenameInvoice.Text = warningCount.ToString();
                txtStatus_RenameInvoice.Text = isProcessing ? "Processing..." : "Ready";
                txtStatus_RenameInvoice.Foreground = isProcessing ?
                    new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Orange) :
                    new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Green);
            }
        }

        private void LogInfo(string message, string tabName = "RenameInvoice")
        {
            LogMessage(LogLevel.Info, message, tabName);
        }

        private void LogWarning(string message, string tabName = "RenameInvoice")
        {
            LogMessage(LogLevel.Warning, message, tabName);
        }

        private void LogError(string message, string tabName = "RenameInvoice")
        {
            LogMessage(LogLevel.Error, message, tabName);
        }

        private void LogSuccess(string message, string tabName = "RenameInvoice")
        {
            LogMessage(LogLevel.Success, message, tabName);
        }

        private void btn_ClearLog_Click(object sender, RoutedEventArgs e)
        {
            string tabName = "RenameInvoice";
            if (sender is System.Windows.Controls.Button button)
            {
                if (button.Name.Contains("Extract"))
                {
                    tabName = "ExtractInvoice";
                    txtLog_ExtractInvoice.Clear();
                }
                else
                {
                    txtLog_RenameInvoice.Clear();
                }
            }

            errorCount = 0;
            warningCount = 0;
            UpdateLogCounters(tabName);
            LogInfo("Log cleared by user", tabName);
        }
        #endregion

        #region File Operations
        private bool LoadFilesFromFolder(string folderPath, out string[] xmlFilesFound, out string[] pdfFilesFound)
        {
            LogInfo($"Scanning folder: {folderPath}");

            xmlFilesFound = Directory.GetFiles(folderPath, "*.xml", SearchOption.TopDirectoryOnly);
            pdfFilesFound = Directory.GetFiles(folderPath, "*.pdf", SearchOption.TopDirectoryOnly);

            LogInfo($"Found {xmlFilesFound.Length} XML files and {pdfFilesFound.Length} PDF files");

            if (xmlFilesFound.Length == 0 && pdfFilesFound.Length == 0)
            {
                LogWarning("No XML or PDF files found in the selected folder");
            }

            return xmlFilesFound.Length > 0 || pdfFilesFound.Length > 0;
        }

        private void PopulateComboBox(System.Windows.Controls.ComboBox comboBox, string[] files)
        {
            comboBox.Items.Clear();
            foreach (var filePath in files)
            {
                comboBox.Items.Add(Path.GetFileName(filePath));
            }

            if (comboBox.Items.Count > 0)
            {
                comboBox.SelectedIndex = 0;
                LogInfo($"Populated file list with {comboBox.Items.Count} files");
            }
        }

        private void ShowXMLContent(string filePath, string fileName)
        {
            try
            {
                LogInfo($"Opening XML file: {fileName}");

                string xmlContent = File.ReadAllText(filePath);
                var doc = new XmlDocument();
                doc.LoadXml(xmlContent);

                using (var stringWriter = new StringWriter())
                using (var xmlTextWriter = new XmlTextWriter(stringWriter))
                {
                    xmlTextWriter.Formatting = Formatting.Indented;
                    doc.WriteContentTo(xmlTextWriter);
                    xmlTextWriter.Flush();
                    xmlContent = stringWriter.GetStringBuilder().ToString();
                }

                var xmlViewer = new Window
                {
                    Title = $"XML Viewer - {fileName}",
                    Width = 800,
                    Height = 600,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Owner = this
                };

                var textBoxXml = new System.Windows.Controls.TextBox
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
                xmlViewer.ShowDialog();

                LogSuccess($"Successfully opened XML file: {fileName}");
            }
            catch (Exception ex)
            {
                LogError($"Error reading XML file '{fileName}': {ex.Message}");
                ShowError($"Error reading or formatting XML file: {ex.Message}");
            }
        }

        private void OpenFileWithDefaultProgram(string filePath)
        {
            try
            {
                LogInfo($"Opening file with default program: {Path.GetFileName(filePath)}");
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                LogSuccess($"Successfully opened file: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                LogError($"Error opening file '{Path.GetFileName(filePath)}': {ex.Message}");
                ShowError($"Error opening file: {ex.Message}");
            }
        }

        private void ShowError(string message)
        {
            if (Dispatcher.CheckAccess())
            {
                MessageBox.Show(this, message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                Dispatcher.Invoke(() => MessageBox.Show(this, message, "Error", MessageBoxButton.OK, MessageBoxImage.Error));
            }
        }

        private void ShowInfo(string message, string title = "Information")
        {
            if (Dispatcher.CheckAccess())
            {
                MessageBox.Show(this, message, title, MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                Dispatcher.Invoke(() => MessageBox.Show(this, message, title, MessageBoxButton.OK, MessageBoxImage.Information));
            }
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
            MessageBox.Show(this, "Support Accountant v1.0\nDeveloped by SonDNT\n© 2025", "About Support Accountant", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void File_Exit_Click(object sender, RoutedEventArgs e)
        {
            LogInfo("Application exit requested");
            System.Windows.Application.Current.Shutdown();
        }
        #endregion

        #region Rename Invoice Management
        private void btn_Browse_RenameInvoice_Click(object sender, RoutedEventArgs e)
        {
            LogInfo("Browse button clicked for folder selection");

            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select XML/PDF Folder";
                folderDialog.ShowNewFolderButton = false;

                var result = folderDialog.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    txtBox_Browse_RenameInvoice.Text = folderDialog.SelectedPath;
                    LogInfo($"Selected folder: {folderDialog.SelectedPath}");
                }
                else
                {
                    LogInfo("Folder selection cancelled by user");
                    return;
                }

                if (!LoadFilesFromFolder(folderDialog.SelectedPath, out string[] xmlFilesFound, out string[] pdfFilesFound))
                {
                    LogError("No XML or PDF files found in the selected folder");
                    ShowError("No XML or PDF files found in the selected folder.");
                    label_TotalFiles_RenameInvoice.Content = "Files found: 0 XML / 0 PDF";
                    ComboBox_RenameInvoice.Items.Clear();
                    return;
                }

                xmlFiles = xmlFilesFound;
                pdfFiles = pdfFilesFound;
                PopulateComboBox(ComboBox_RenameInvoice, xmlFiles.Concat(pdfFiles).ToArray());
                label_TotalFiles_RenameInvoice.Content = $"Files found: {xmlFiles.Length} XML / {pdfFiles.Length} PDF";
                UpdateProgressBar(0, 1, "");

                LogSuccess($"Successfully loaded {xmlFiles.Length} XML and {pdfFiles.Length} PDF files");
                ShowInfo($"Loaded {xmlFiles.Length} XML and {pdfFiles.Length} PDF files from the selected folder.");
            }
        }

        private void btn_Openfile_RenameInvoice_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtBox_Browse_RenameInvoice.Text) || ComboBox_RenameInvoice.SelectedItem == null)
            {
                LogWarning("No folder or file selected for opening");
                ShowError("Please select a folder and a file.");
                return;
            }

            string folderPath = txtBox_Browse_RenameInvoice.Text;
            string fileName = ComboBox_RenameInvoice.SelectedItem.ToString();
            string filePath = Path.Combine(folderPath, fileName);

            if (!File.Exists(filePath))
            {
                LogError($"Selected file does not exist: {fileName}");
                ShowError("Selected file does not exist.");
                return;
            }

            string extension = Path.GetExtension(fileName).ToLower();

            switch (extension)
            {
                case ".xml":
                    ShowXMLContent(filePath, fileName);
                    break;
                case ".pdf":
                    OpenFileWithDefaultProgram(filePath);
                    break;
                default:
                    LogError($"Unsupported file type: {extension}");
                    ShowError("Unsupported file type.");
                    break;
            }
        }

        private async void btn_Renameall_RenameInvoice_Click(object sender, RoutedEventArgs e)
        {
            if (isProcessing)
            {
                LogWarning("Rename process already in progress - request ignored");
                ShowInfo("Processing is already in progress. Please wait for it to complete.");
                return;
            }

            if (xmlFiles == null || pdfFiles == null || (xmlFiles.Length == 0 && pdfFiles.Length == 0))
            {
                LogError("No files loaded for renaming process");
                ShowError("No files loaded. Please select a folder with XML/PDF files first.");
                return;
            }

            LogInfo("Starting rename process - selecting destination folder");

            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select folder to store renamed files";
                folderDialog.ShowNewFolderButton = true;

                if (folderDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                {
                    LogInfo("Destination folder selection cancelled");
                    return;
                }

                folderDestination = folderDialog.SelectedPath;
                LogInfo($"Selected destination folder: {folderDestination}");

                string renamedFolder = Path.Combine(folderDestination, "Renamed");
                string failedFolder = Path.Combine(folderDestination, "Renamed_failed");

                try
                {
                    if (Directory.Exists(renamedFolder))
                    {
                        Directory.Delete(renamedFolder, true);
                        LogInfo("Removed existing 'Renamed' folder");
                    }

                    if (Directory.Exists(failedFolder))
                    {
                        Directory.Delete(failedFolder, true);
                        LogInfo("Removed existing 'Renamed_failed' folder");
                    }

                    Directory.CreateDirectory(renamedFolder);
                    Directory.CreateDirectory(failedFolder);

                    LogSuccess($"Created folders: Renamed and Renamed_failed");
                    ShowInfo($"Folders created successfully:\n- {renamedFolder}\n- {failedFolder}", "Folders Created");
                }
                catch (Exception ex)
                {
                    LogError($"Error creating destination folders: {ex.Message}");
                    ShowError($"Error creating folders: {ex.Message}");
                    return;
                }

                cancellationTokenSource = new CancellationTokenSource();
                btn_Renameall_RenameInvoice.IsEnabled = false;
                btn_Stop_RenameInvoice.IsEnabled = true;
                ExportXml_Tab.Visibility = Visibility.Collapsed;
                isProcessing = true;
                UpdateLogCounters();

                LogInfo("Starting file renaming process...");

                try
                {
                    await ProcessFilesForRenamingAsync(renamedFolder, failedFolder, cancellationTokenSource.Token);
                }
                catch (OperationCanceledException)
                {
                    LogWarning("Rename operation was cancelled by user");
                    UpdateProgressBar(0, 1, "Operation Cancelled");
                    ShowInfo("Rename operation was cancelled.", "Operation Cancelled");
                }
                finally
                {
                    btn_Renameall_RenameInvoice.IsEnabled = true;
                    btn_Stop_RenameInvoice.IsEnabled = false;
                    ExportXml_Tab.Visibility = Visibility.Visible;
                    isProcessing = false;
                    UpdateLogCounters();
                    LogInfo("Rename process completed");
                    cancellationTokenSource?.Dispose();
                    cancellationTokenSource = null;
                }
            }
        }

        private void btn_Stop_RenameInvoice_Click(object sender, RoutedEventArgs e)
        {
            if (cancellationTokenSource != null && !cancellationTokenSource.Token.IsCancellationRequested)
            {
                LogWarning("Stop button clicked - cancelling rename operation");
                cancellationTokenSource.Cancel();
                btn_Stop_RenameInvoice.IsEnabled = false;
                UpdateProgressBar(0, 1, "Stopping...");
            }
        }

        public async Task ProcessFilesForRenamingAsync(string renamedFolder, string failedFolder, CancellationToken cancellationToken)
        {
            int totalFiles = (xmlFiles?.Length ?? 0) + (pdfFiles?.Length ?? 0);

            if (totalFiles == 0)
            {
                LogWarning("No files to process");
                ShowInfo("No files to process.", "Information");
                return;
            }

            LogInfo($"Processing {totalFiles} total files ({xmlFiles?.Length ?? 0} XML, {pdfFiles?.Length ?? 0} PDF)");
            UpdateProgressBar(0, totalFiles, "Starting...");

            int xmlSuccessCount = 0;
            int xmlFailedCount = 0;
            int pdfSuccessCount = 0;
            int pdfFailedCount = 0;

            try
            {
                if (xmlFiles != null && xmlFiles.Length > 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    LogInfo($"Starting XML processing: {xmlFiles.Length} files");
                    UpdateProgressBar(0, totalFiles, "Processing XML files...");

                    var xmlResults = await ProcessXmlFilesAsync(xmlFiles, renamedFolder, failedFolder, cancellationToken);
                    xmlSuccessCount = xmlResults.successCount;
                    xmlFailedCount = xmlResults.failedCount;

                    LogInfo($"XML processing completed: {xmlSuccessCount} success, {xmlFailedCount} failed");
                }

                if (pdfFiles != null && pdfFiles.Length > 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    LogInfo($"Starting PDF processing: {pdfFiles.Length} files");
                    UpdateProgressBar(xmlFiles?.Length ?? 0, totalFiles, "Processing PDF files...");

                    var pdfResults = await ProcessPdfFilesBatchAsync(pdfFiles, renamedFolder, failedFolder, cancellationToken);
                    pdfSuccessCount = pdfResults.successCount;
                    pdfFailedCount = pdfResults.failedCount;

                    LogInfo($"PDF processing completed: {pdfSuccessCount} success, {pdfFailedCount} failed");
                }

                UpdateProgressBar(totalFiles, totalFiles, "Processing Complete!");
                await Task.Delay(1000, cancellationToken);

                int totalSuccess = xmlSuccessCount + pdfSuccessCount;
                int totalFailed = xmlFailedCount + pdfFailedCount;

                LogSuccess($"All processing completed - Total Success: {totalSuccess}, Total Failed: {totalFailed}");

                ShowInfo($"Processing completed:\n" +
                        $"XML Files - Success: {xmlSuccessCount}, Failed: {xmlFailedCount}\n" +
                        $"PDF Files - Success: {pdfSuccessCount}, Failed: {pdfFailedCount}\n" +
                        $"Total - Success: {totalSuccess}, Failed: {totalFailed}",
                        "Rename Process Complete");
            }
            catch (OperationCanceledException)
            {
                LogWarning("Processing operation was cancelled");
                throw;
            }
            catch (Exception ex)
            {
                LogError($"Critical error during processing: {ex.Message}");
                ShowError($"Error during processing: {ex.Message}");
            }
            finally
            {
                UpdateProgressBar(0, 1, "");
            }
        }

        private async Task<(int successCount, int failedCount)> ProcessXmlFilesAsync(string[] xmlFilePaths, string renamedFolder, string failedFolder, CancellationToken cancellationToken)
        {
            int successCount = 0;
            int failedCount = 0;

            foreach (string xmlFilePath in xmlFilePaths)
            {
                cancellationToken.ThrowIfCancellationRequested();

                string fileName = Path.GetFileName(xmlFilePath);
                LogInfo($"Processing XML file: {fileName}");
                UpdateProgressBar(successCount + failedCount, xmlFilePaths.Length + (pdfFiles?.Length ?? 0), $"Processing XML: {fileName}");

                bool success = await Task.Run(() =>
                {
                    try
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        return ProcessSingleXmlFile(xmlFilePath, renamedFolder, failedFolder);
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        LogError($"Exception processing XML file '{fileName}': {ex.Message}");
                        CopyToFailedFolder(xmlFilePath, failedFolder, $"Processing error: {ex.Message}");
                        return false;
                    }
                }, cancellationToken);

                if (success)
                {
                    successCount++;
                    LogSuccess($"Successfully processed XML: {fileName}");
                }
                else
                {
                    failedCount++;
                    LogError($"Failed to process XML: {fileName}");
                }

                await Task.Delay(50, cancellationToken);
            }

            return (successCount, failedCount);
        }

        private async Task<(int successCount, int failedCount)> ProcessPdfFilesBatchAsync(string[] pdfFilePaths, string renamedFolder, string failedFolder, CancellationToken cancellationToken)
        {
            try
            {
                LogInfo("Starting PDF batch processing with Python script");

                string pythonExe = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "portable/python", "python.exe");
                string scriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "portable/app", "pdf_rename.py");

                if (!File.Exists(pythonExe))
                {
                    LogError($"Python executable not found: {pythonExe}");
                    ShowError($"Python executable not found: {pythonExe}");
                    return (0, pdfFilePaths.Length);
                }

                if (!File.Exists(scriptPath))
                {
                    LogError($"Python script not found: {scriptPath}");
                    ShowError($"Python script not found: {scriptPath}");
                    return (0, pdfFilePaths.Length);
                }

                var arguments = new List<string>
                {
                    $"\"{scriptPath}\"",
                    "-i"
                };

                foreach (string pdfPath in pdfFilePaths)
                {
                    arguments.Add($"\"{pdfPath}\"");
                }

                arguments.Add("-s");
                arguments.Add($"\"{renamedFolder}\"");
                arguments.Add("-f");
                arguments.Add($"\"{failedFolder}\"");

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = pythonExe,
                    Arguments = string.Join(" ", arguments),
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                LogInfo($"Executing Python command: {pythonExe} {psi.Arguments}");

                return await Task.Run(() =>
                {
                    using (Process process = Process.Start(psi))
                    {
                        int successCount = 0;
                        int failedCount = 0;
                        int processedCount = 0;
                        int baseProgress = xmlFiles?.Length ?? 0;

                        string line;
                        while ((line = process.StandardOutput.ReadLine()) != null)
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            LogInfo($"Python output: {line}");

                            if (line.StartsWith("PROGRESS:"))
                            {
                                processedCount++;
                                string fileName = ExtractFileNameFromProgressLine(line);

                                if (line.Contains("SUCCESS"))
                                {
                                    successCount++;
                                    Dispatcher.Invoke(() =>
                                    {
                                        UpdateProgressBar(baseProgress + processedCount,
                                            (xmlFiles?.Length ?? 0) + pdfFilePaths.Length,
                                            $"Completed PDF: {fileName}");
                                        LogSuccess($"PDF processed successfully: {fileName}");
                                    });
                                }
                                else if (line.Contains("FAILED") || line.Contains("ERROR"))
                                {
                                    failedCount++;
                                    Dispatcher.Invoke(() =>
                                    {
                                        UpdateProgressBar(baseProgress + processedCount,
                                            (xmlFiles?.Length ?? 0) + pdfFilePaths.Length,
                                            $"Failed PDF: {fileName}");
                                        LogError($"PDF processing failed: {fileName}");
                                    });
                                }
                            }
                            else if (line.StartsWith("SUMMARY:"))
                            {
                                var summaryMatch = System.Text.RegularExpressions.Regex.Match(line,
                                    @"SUCCESS=(\d+), FAILED=(\d+)");
                                if (summaryMatch.Success)
                                {
                                    successCount = int.Parse(summaryMatch.Groups[1].Value);
                                    failedCount = int.Parse(summaryMatch.Groups[2].Value);
                                    LogInfo($"Python script summary: {successCount} success, {failedCount} failed");
                                }
                            }
                        }

                        string errors = process.StandardError.ReadToEnd();
                        process.WaitForExit();

                        if (!string.IsNullOrEmpty(errors))
                        {
                            LogError($"Python script errors: {errors}");
                            Debug.WriteLine($"Python script errors: {errors}");
                        }

                        LogInfo($"Python script completed with exit code: {process.ExitCode}");
                        return (successCount, failedCount);
                    }
                }, cancellationToken);
            }
            catch (OperationCanceledException)
            {
                LogWarning("PDF processing was cancelled");
                throw;
            }
            catch (Exception ex)
            {
                LogError($"Error executing Python batch script: {ex.Message}");
                Debug.WriteLine($"Error executing Python batch script: {ex.Message}");
                ShowError($"Error processing PDF files: {ex.Message}");
                return (0, pdfFilePaths.Length);
            }
        }

        private string ExtractFileNameFromProgressLine(string progressLine)
        {
            var match = System.Text.RegularExpressions.Regex.Match(progressLine, @"PROGRESS:\s+(.+?)\s+->");
            return match.Success ? match.Groups[1].Value : "Unknown file";
        }

        private bool ProcessSingleXmlFile(string xmlFilePath, string renamedFolder, string failedFolder)
        {
            try
            {
                var doc = new XmlDocument();
                doc.Load(xmlFilePath);
                string sHDon = doc.SelectSingleNode("//SHDon")?.InnerText?.Trim();
                string nLap = doc.SelectSingleNode("//NLap")?.InnerText?.Trim();

                if (string.IsNullOrEmpty(sHDon) || string.IsNullOrEmpty(nLap))
                {
                    LogWarning($"Missing SHDon or NLap data in file: {Path.GetFileName(xmlFilePath)}");
                    CopyToFailedFolder(xmlFilePath, failedFolder, "Missing SHDon or NLap data");
                    return false;
                }

                if (!DateTime.TryParse(nLap, out DateTime parsedDate))
                {
                    LogWarning($"Invalid date format in NLap for file: {Path.GetFileName(xmlFilePath)} (Date: {nLap})");
                    CopyToFailedFolder(xmlFilePath, failedFolder, "Invalid date format in NLap");
                    return false;
                }

                string datePrefix = parsedDate.ToString("yyMMdd");
                string newFileName = $"{datePrefix}_{sHDon}.xml";
                string finalFileName = GetUniqueFileName(renamedFolder, newFileName);
                string destinationPath = Path.Combine(renamedFolder, finalFileName);
                File.Copy(xmlFilePath, destinationPath, false);

                LogInfo($"Renamed {Path.GetFileName(xmlFilePath)} to {finalFileName}");
                return true;
            }
            catch (Exception ex)
            {
                LogError($"Error processing XML file {Path.GetFileName(xmlFilePath)}: {ex.Message}");
                CopyToFailedFolder(xmlFilePath, failedFolder, "General processing error");
                return false;
            }
        }

        private void CopyToFailedFolder(string sourceFilePath, string failedFolder, string reason)
        {
            try
            {
                string fileName = Path.GetFileName(sourceFilePath);
                string failedFilePath = Path.Combine(failedFolder, fileName);
                string uniqueFailedPath = GetUniqueFileName(failedFolder, fileName);
                failedFilePath = Path.Combine(failedFolder, uniqueFailedPath);
                File.Copy(sourceFilePath, failedFilePath, false);

                LogWarning($"Copied to failed folder: {fileName} (Reason: {reason})");
            }
            catch (Exception ex)
            {
                LogError($"Error copying to failed folder: {ex.Message}");
                ShowError($"Error copying to failed folder: {ex.Message}");
            }
        }

        private string GetUniqueFileName(string directory, string fileName)
        {
            string nameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
            string extension = Path.GetExtension(fileName);
            string finalFileName = fileName;
            int counter = 1;

            while (File.Exists(Path.Combine(directory, finalFileName)))
            {
                finalFileName = $"{nameWithoutExtension}({counter}){extension}";
                counter++;
            }

            return finalFileName;
        }

        private void btn_OpenFolder_RenameInvoice_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(folderDestination) || !Directory.Exists(folderDestination))
            {
                LogWarning("No valid destination folder to open");
                ShowError("No valid destination folder to open. Please complete a rename operation first.");
                return;
            }
            try
            {
                LogInfo($"Opening destination folder: {folderDestination}");
                Process.Start(new ProcessStartInfo
                {
                    FileName = folderDestination,
                    UseShellExecute = true,
                    Verb = "open"
                });
                LogSuccess("Successfully opened destination folder");
            }
            catch (Exception ex)
            {
                LogError($"Error opening destination folder: {ex.Message}");
                ShowError($"Error opening destination folder: {ex.Message}");
            }
        }
        #endregion

        #region Extract Invoice Management
        private void btn_Browse_ExtractInvoice_Click(object sender, RoutedEventArgs e)
        {
            LogInfo("Browse button clicked for extract invoice folder selection", "ExtractInvoice");

            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select XML Folder for Invoice Extraction";
                folderDialog.ShowNewFolderButton = false;

                var result = folderDialog.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    txtBox_Browse_ExtractInvoice.Text = folderDialog.SelectedPath;
                    LogInfo($"Selected folder: {folderDialog.SelectedPath}", "ExtractInvoice");
                }
                else
                {
                    LogInfo("Folder selection cancelled by user", "ExtractInvoice");
                    return;
                }

                if (!LoadExtractFilesFromFolder(folderDialog.SelectedPath, out string[] xmlFilesFound))
                {
                    LogError("No XML files found in the selected folder", "ExtractInvoice");
                    ShowError("No XML files found in the selected folder.");
                    label_TotalFiles_ExtractInvoice.Content = "Files found: 0 XML";
                    ComboBox_ExtractInvoice.Items.Clear();
                    return;
                }

                extractXmlFiles = xmlFilesFound;
                PopulateComboBox(ComboBox_ExtractInvoice, extractXmlFiles);
                label_TotalFiles_ExtractInvoice.Content = $"Files found: {extractXmlFiles.Length} XML";
                UpdateProgressBar(0, 1, "");

                LogSuccess($"Successfully loaded {extractXmlFiles.Length} XML files for extraction", "ExtractInvoice");
                ShowInfo($"Loaded {extractXmlFiles.Length} XML files from the selected folder.");
            }
        }

        private bool LoadExtractFilesFromFolder(string folderPath, out string[] xmlFilesFound)
        {
            LogInfo($"Scanning folder for XML files: {folderPath}", "ExtractInvoice");

            xmlFilesFound = Directory.GetFiles(folderPath, "*.xml", SearchOption.TopDirectoryOnly);

            LogInfo($"Found {xmlFilesFound.Length} XML files", "ExtractInvoice");

            if (xmlFilesFound.Length == 0)
            {
                LogWarning("No XML files found in the selected folder", "ExtractInvoice");
            }

            return xmlFilesFound.Length > 0;
        }

        private void btn_Openfile_ExtractInvoice_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtBox_Browse_ExtractInvoice.Text) || ComboBox_ExtractInvoice.SelectedItem == null)
            {
                LogWarning("No folder or file selected for opening", "ExtractInvoice");
                ShowError("Please select a folder and an XML file.");
                return;
            }

            string folderPath = txtBox_Browse_ExtractInvoice.Text;
            string fileName = ComboBox_ExtractInvoice.SelectedItem.ToString();
            string filePath = Path.Combine(folderPath, fileName);

            if (!File.Exists(filePath))
            {
                LogError($"Selected file does not exist: {fileName}", "ExtractInvoice");
                ShowError("Selected XML file does not exist.");
                return;
            }

            ShowXMLContent(filePath, fileName);
        }

        private async void btn_ExportSummary_ExtractInvoice_Click(object sender, RoutedEventArgs e)
        {
            if (isProcessing)
            {
                LogWarning("Extract process already in progress - request ignored", "ExtractInvoice");
                ShowInfo("Processing is already in progress. Please wait for it to complete.");
                return;
            }

            if (extractXmlFiles == null || extractXmlFiles.Length == 0)
            {
                LogError("No XML files loaded for extraction process", "ExtractInvoice");
                ShowError("No XML files loaded. Please select a folder with XML files first.");
                return;
            }

            LogInfo("Starting invoice extraction process - selecting output Excel file", "ExtractInvoice");

            using (var saveDialog = new SaveFileDialog())
            {
                saveDialog.Title = "Export Invoice Summary to Excel";
                saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                saveDialog.FileName = $"invoice_summary.xlsx";

                if (saveDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                {
                    LogInfo("Excel file selection cancelled", "ExtractInvoice");
                    return;
                }

                lastExcelFilePath = saveDialog.FileName;
                LogInfo($"Selected output Excel file: {lastExcelFilePath}", "ExtractInvoice");

                cancellationTokenSource = new CancellationTokenSource();
                btn_ExportSummary_ExtractInvoice.IsEnabled = false;
                btn_Stop_ExtractInvoice.IsEnabled = true;
                rename_tab.Visibility = Visibility.Collapsed;
                isProcessing = true;
                UpdateLogCounters("ExtractInvoice");

                LogInfo("Starting Excel export process...", "ExtractInvoice");

                try
                {
                    await ExportToExcelAsync(lastExcelFilePath, cancellationTokenSource.Token);
                }
                catch (OperationCanceledException)
                {
                    LogWarning("Export operation was cancelled by user", "ExtractInvoice");
                    UpdateProgressBar(0, 1, "Operation Cancelled");
                    ShowInfo("Export operation was cancelled.", "Operation Cancelled");
                }
                finally
                {
                    btn_ExportSummary_ExtractInvoice.IsEnabled = true;
                    btn_Stop_ExtractInvoice.IsEnabled = false;
                    rename_tab.Visibility = Visibility.Visible;
                    isProcessing = false;
                    UpdateLogCounters("ExtractInvoice");
                    LogInfo("Export process completed", "ExtractInvoice");
                    cancellationTokenSource?.Dispose();
                    cancellationTokenSource = null;
                }
            }
        }

        private void btn_Stop_ExtractInvoice_Click(object sender, RoutedEventArgs e)
        {
            if (cancellationTokenSource != null && !cancellationTokenSource.Token.IsCancellationRequested)
            {
                LogWarning("Stop button clicked - cancelling extract operation", "ExtractInvoice");
                cancellationTokenSource.Cancel();
                btn_Stop_ExtractInvoice.IsEnabled = false;
                UpdateProgressBar(0, 1, "Stopping...");
            }
        }

        private void btn_OpenExcel_ExtractInvoice_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(lastExcelFilePath) || !File.Exists(lastExcelFilePath))
            {
                LogWarning("No valid Excel file to open", "ExtractInvoice");
                ShowError("No valid Excel file to open. Please complete an export operation first.");
                return;
            }

            try
            {
                LogInfo($"Opening Excel file: {lastExcelFilePath}", "ExtractInvoice");
                Process.Start(new ProcessStartInfo
                {
                    FileName = lastExcelFilePath,
                    UseShellExecute = true
                });
                LogSuccess("Successfully opened Excel file", "ExtractInvoice");
            }
            catch (Exception ex)
            {
                LogError($"Error opening Excel file: {ex.Message}", "ExtractInvoice");
                ShowError($"Error opening Excel file: {ex.Message}");
            }
        }

        private async Task ExportToExcelAsync(string excelPath, CancellationToken cancellationToken)
        {
            LogInfo($"Starting Excel export to: {excelPath}", "ExtractInvoice");
            UpdateProgressBar(0, extractXmlFiles.Length, "Initializing...");

            try
            {
                using (var package = new ExcelPackage())
                {
                    var summarySheet = package.Workbook.Worksheets.Add("Summary");
                    CreateDynamicSummaryHeaders(summarySheet);

                    int row = 2;
                    int processedCount = 0;

                    foreach (var xmlFile in extractXmlFiles)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        string fileName = Path.GetFileName(xmlFile);
                        LogInfo($"Processing XML file: {fileName}", "ExtractInvoice");

                        Dispatcher.Invoke(() =>
                        {
                            UpdateProgressBar(processedCount, extractXmlFiles.Length, $"Processing: {fileName}");
                        });

                        try
                        {
                            ProcessXmlFileForExtraction(package, summarySheet, xmlFile, ref row);
                            LogSuccess($"Successfully processed: {fileName}", "ExtractInvoice");
                        }
                        catch (OperationCanceledException)
                        {
                            LogWarning("Operation was cancelled during file processing", "ExtractInvoice");
                            throw;
                        }
                        catch (Exception ex)
                        {
                            LogError($"Error processing file '{fileName}': {ex.Message}", "ExtractInvoice");
                            continue;
                        }

                        processedCount++;

                        if (processedCount % 10 == 0)
                        {
                            await Task.Delay(50, cancellationToken);
                        }
                    }

                    Dispatcher.Invoke(() =>
                    {
                        UpdateProgressBar(extractXmlFiles.Length, extractXmlFiles.Length, "Saving Excel file...");
                    });

                    LogInfo("Finalizing Excel file...", "ExtractInvoice");

                    summarySheet.Cells[summarySheet.Dimension.Address].AutoFitColumns();
                    var fileInfo = new FileInfo(excelPath);
                    package.SaveAs(fileInfo);
                    Dispatcher.Invoke(() =>
                    {
                        UpdateProgressBar(extractXmlFiles.Length, extractXmlFiles.Length, "Export Complete!");
                    });

                    LogSuccess($"Excel export completed successfully: {processedCount} files processed", "ExtractInvoice");

                    Dispatcher.Invoke(() =>
                    {
                        ShowInfo($"Excel export completed successfully!\nProcessed: {processedCount} files\nSaved to: {excelPath}", "Export Complete");
                    });
                }
            }
            catch (OperationCanceledException)
            {
                LogWarning("Excel export was cancelled", "ExtractInvoice");
                throw;
            }
            catch (Exception ex)
            {
                LogError($"Critical error during Excel export: {ex.Message}", "ExtractInvoice");
                Dispatcher.Invoke(() =>
                {
                    ShowError($"Error exporting to Excel: {ex.Message}");
                });
                throw;
            }
            finally
            {
                Dispatcher.Invoke(() =>
                {
                    UpdateProgressBar(0, 1, "");
                });
            }
        }

        private void CreateDynamicSummaryHeaders(ExcelWorksheet summarySheet)
        {
            var headers = new List<string>
            {
                "Tên Sheet", "Số Hóa Đơn", "Ngày Lập"
            };

            if (checkBox_Seller.IsChecked == true)
            {
                headers.AddRange(new[] { "Tên Người Bán", "MST Người Bán", "Địa Chỉ Người Bán" });
            }
            if (checkBox_Buyer.IsChecked == true)
            {
                headers.AddRange(new[] { "Tên Người Mua", "MST Người Mua", "Địa Chỉ Người Mua" });
            }
            if (checkBox_TongSL.IsChecked == true)
            {
                headers.AddRange(new[] { "Tổng Số Lượng" });
            }
            if (checkBox_TongTien.IsChecked == true)
            {
                headers.AddRange(new[] { "Tổng Tiền" });
            }
            if (checkBox_TienThue.IsChecked == true)
            {
                headers.AddRange(new[] { "Tiền Thuế" });
            }
            if (checkBox_ThanhTien.IsChecked == true)
            {
                headers.AddRange(new[] { "Thành Tiền" });
            }
            if (checkBox_Currency.IsChecked == true)
            {
                headers.AddRange(new[] { "Đơn Vị Tiền Tệ" });
            }

            for (int i = 0; i < headers.Count; i++)
            {
                summarySheet.Cells[1, i + 1].Value = headers[i];
                summarySheet.Cells[1, i + 1].Style.Font.Bold = true;
            }

            LogInfo($"Created dynamic headers with {headers.Count} columns", "ExtractInvoice");
        }

        private void ProcessXmlFileForExtraction(ExcelPackage package, ExcelWorksheet summarySheet, string xmlFile, ref int row)
        {
            var doc = new XmlDocument();
            doc.Load(xmlFile);

            string fileName = Path.GetFileNameWithoutExtension(xmlFile);
            string sheetName = CreateSafeSheetName(fileName);

            ExcelWorksheet fileSheet = null;
            fileSheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName);
            if (fileSheet == null)
            {
                fileSheet = package.Workbook.Worksheets.Add(sheetName);
                CreateDetailSheet(fileSheet, doc);
            }

            var values = ExtractDynamicSummaryData(doc, fileName, sheetName);
            PopulateDynamicSummaryRow(summarySheet, fileSheet, values, row);
            row++;
        }

        private string CreateSafeSheetName(string name)
        {
            string safeName = name.Length > 31 ? name.Substring(0, 31) : name;
            char[] invalidChars = { ':', '\\', '/', '?', '*', '[', ']' };
            foreach (char c in invalidChars)
            {
                safeName = safeName.Replace(c, '_');
            }
            return safeName;
        }

        private void CreateDetailSheet(ExcelWorksheet sheet, XmlDocument doc)
        {
            var detailHeaders = new List<string>
            {
                "STT", "THHDVu (Tên hàng hóa/dịch vụ)", "DVTinh (Đơn vị tính)",
                "SLuong (Số lượng)", "DGia (Đơn giá)", "ThTien (Thành tiền)",
                "TSuat (Thuế suất)", "DVTTe (Đơn vị tiền tệ)"
            };

            if (checkBox_Seller2.IsChecked == true)
            {
                detailHeaders.AddRange(new[] { "Tên Người Bán", "MST Người Bán", "Địa Chỉ Người Bán" });
            }
            if (checkBox_Buyer2.IsChecked == true)
            {
                detailHeaders.AddRange(new[] { "Tên Người Mua", "MST Người Mua", "Địa Chỉ Người Mua" });
            }

            for (int i = 0; i < detailHeaders.Count; i++)
            {
                sheet.Cells[1, i + 1].Value = detailHeaders[i];
                sheet.Cells[1, i + 1].Style.Font.Bold = true;
            }

            var nodes = doc.SelectNodes("//HHDVu");
            int detailRow = 2;
            string currency = doc.SelectSingleNode("//DVTTe")?.InnerText ?? "";

            foreach (XmlNode node in nodes)
            {
                PopulateDetailRow(sheet, node, doc, currency, detailRow);
                detailRow++;
            }

            AddSummaryTable(sheet, doc, detailRow + 2);
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
        }

        private void PopulateDetailRow(ExcelWorksheet sheet, XmlNode node, XmlDocument doc, string currency, int row)
        {
            int col = 1;
            sheet.Cells[row, col++].Value = node.SelectSingleNode("STT")?.InnerText ?? "";
            sheet.Cells[row, col++].Value = node.SelectSingleNode("THHDVu")?.InnerText ?? "";
            sheet.Cells[row, col++].Value = node.SelectSingleNode("DVTinh")?.InnerText ?? "";
            sheet.Cells[row, col++].Value = FormatDecimalString(node.SelectSingleNode("SLuong")?.InnerText ?? "");

            string dGia = node.SelectSingleNode("DGia")?.InnerText ?? "";
            string thTien = node.SelectSingleNode("ThTien")?.InnerText ?? "";
            string tSuat = node.SelectSingleNode("TSuat")?.InnerText ?? "";

            sheet.Cells[row, col++].Value = string.IsNullOrEmpty(dGia) ? "" : $"{FormatDecimalString(dGia)} {currency}";
            sheet.Cells[row, col++].Value = string.IsNullOrEmpty(thTien) ? "" : $"{FormatDecimalString(thTien)} {currency}";
            sheet.Cells[row, col++].Value = string.IsNullOrEmpty(tSuat) ? "" : FormatDecimalString(tSuat);
            sheet.Cells[row, col++].Value = currency;

            if (checkBox_Seller2.IsChecked == true)
            {
                sheet.Cells[row, col++].Value = doc.SelectSingleNode("//NBan/Ten")?.InnerText ?? "";
                sheet.Cells[row, col++].Value = doc.SelectSingleNode("//NBan/MST")?.InnerText ?? "";
                sheet.Cells[row, col++].Value = doc.SelectSingleNode("//NBan/DChi")?.InnerText ?? "";
            }
            if (checkBox_Buyer2.IsChecked == true)
            {
                sheet.Cells[row, col++].Value = doc.SelectSingleNode("//NMua/Ten")?.InnerText ?? "";
                sheet.Cells[row, col++].Value = doc.SelectSingleNode("//NMua/MST")?.InnerText ?? "";
                sheet.Cells[row, col++].Value = doc.SelectSingleNode("//NMua/DChi")?.InnerText ?? "";
            }
        }

        private void AddSummaryTable(ExcelWorksheet sheet, XmlDocument doc, int startRow)
        {
            var tToanNode = doc.SelectSingleNode("//TToan");
            if (tToanNode == null) return;

            sheet.Cells[startRow, 1].Value = "Thành tiền";
            sheet.Cells[startRow, 2].Value = "Thuế suất";
            sheet.Cells[startRow, 3].Value = "Tiền thuế";

            for (int i = 1; i <= 3; i++)
            {
                sheet.Cells[startRow, i].Style.Font.Bold = true;
            }

            var ltsuatNodes = tToanNode.SelectNodes("THTTLTSuat/LTSuat");
            int tRow = startRow + 1;
            foreach (XmlNode ltsuat in ltsuatNodes)
            {
                sheet.Cells[tRow, 1].Value = FormatDecimalString(ltsuat.SelectSingleNode("ThTien")?.InnerText ?? "");
                sheet.Cells[tRow, 2].Value = ltsuat.SelectSingleNode("TSuat")?.InnerText ?? "";
                sheet.Cells[tRow, 3].Value = FormatDecimalString(ltsuat.SelectSingleNode("TThue")?.InnerText ?? "");
                tRow++;
            }

            int summaryRow = tRow + 1;
            sheet.Cells[summaryRow, 1].Value = "Tổng cộng (chưa thuế):";
            sheet.Cells[summaryRow, 2].Value = FormatDecimalString(tToanNode.SelectSingleNode("TgTCThue")?.InnerText ?? "");
            sheet.Cells[summaryRow + 1, 1].Value = "Tổng tiền thuế:";
            sheet.Cells[summaryRow + 1, 2].Value = FormatDecimalString(tToanNode.SelectSingleNode("TgTThue")?.InnerText ?? "");
            sheet.Cells[summaryRow + 2, 1].Value = "Tổng cộng (đã thuế):";
            sheet.Cells[summaryRow + 2, 2].Value = FormatDecimalString(tToanNode.SelectSingleNode("TgTTTBSo")?.InnerText ?? "");
            sheet.Cells[summaryRow + 3, 1].Value = "Bằng chữ:";
            sheet.Cells[summaryRow + 3, 2].Value = tToanNode.SelectSingleNode("TgTTTBChu")?.InnerText ?? "";

            for (int r = summaryRow; r <= summaryRow + 3; r++)
            {
                sheet.Cells[r, 1].Style.Font.Bold = true;
            }
        }

        private List<object> ExtractDynamicSummaryData(XmlDocument doc, string fileName, string sheetName)
        {
            var values = new List<object>
            {
                sheetName,
                doc.SelectSingleNode("//SHDon")?.InnerText ?? "",
                doc.SelectSingleNode("//NLap")?.InnerText ?? ""
            };

            if (checkBox_Seller.IsChecked == true)
            {
                values.Add(doc.SelectSingleNode("//NBan/Ten")?.InnerText ?? "");
                values.Add(doc.SelectSingleNode("//NBan/MST")?.InnerText ?? "");
                values.Add(doc.SelectSingleNode("//NBan/DChi")?.InnerText ?? "");
            }

            if (checkBox_Buyer.IsChecked == true)
            {
                values.Add(doc.SelectSingleNode("//NMua/Ten")?.InnerText ?? "");
                values.Add(doc.SelectSingleNode("//NMua/MST")?.InnerText ?? "");
                values.Add(doc.SelectSingleNode("//NMua/DChi")?.InnerText ?? "");
            }

            if (checkBox_TongSL.IsChecked == true)
            {
                values.Add(doc.SelectNodes("//HHDVu/STT")?.Count ?? 0);
            }
            if (checkBox_TongTien.IsChecked == true)
            {
                values.Add(FormatDecimalString(doc.SelectSingleNode("//TgTCThue")?.InnerText ?? ""));
            }
            if (checkBox_TienThue.IsChecked == true)
            {
                values.Add(FormatDecimalString(doc.SelectSingleNode("//TgTThue")?.InnerText ?? ""));
            }
            if (checkBox_ThanhTien.IsChecked == true)
            {
                values.Add(FormatDecimalString(doc.SelectSingleNode("//TgTTTBSo")?.InnerText ?? ""));
            }
            if (checkBox_Currency.IsChecked == true)
            {
                values.Add(doc.SelectSingleNode("//DVTTe")?.InnerText ?? "");
            }
            return values;
        }

        private void PopulateDynamicSummaryRow(ExcelWorksheet summarySheet, ExcelWorksheet fileSheet, List<object> values, int row)
        {
            for (int col = 0; col < values.Count; col++)
            {
                if (col == 0 && fileSheet != null)
                {
                    summarySheet.Cells[row, col + 1].Hyperlink = new ExcelHyperLink($"'{fileSheet.Name}'!A1", fileSheet.Name);
                    summarySheet.Cells[row, col + 1].Value = fileSheet.Name;
                    summarySheet.Cells[row, col + 1].Style.Font.UnderLine = true;
                }
                else
                {
                    summarySheet.Cells[row, col + 1].Value = values[col];
                }
            }
        }

        private string FormatDecimalString(string value)
        {
            if (decimal.TryParse(value, out decimal result))
            {
                if (result == Math.Truncate(result))
                    return result.ToString("#,##0", CultureInfo.InvariantCulture);
                return result.ToString("#,##0.###", CultureInfo.InvariantCulture);
            }
            return value;
        }
        #endregion

        #region Rename Invoice ToolTips
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
        #endregion

        #region Extract Invoice ToolTips
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
        #endregion
    }
}