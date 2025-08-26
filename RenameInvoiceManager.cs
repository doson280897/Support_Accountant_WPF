using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Xml;

namespace Support_Accountant
{
    public class RenameInvoiceController
    {
        private readonly MainWindow mainWindow;

        public RenameInvoiceController(MainWindow window)
        {
            mainWindow = window;
        }

        public void Browse_Click()
        {
            mainWindow.LogInfo("Browse button clicked for folder selection");

            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select XML/PDF Folder";
                folderDialog.ShowNewFolderButton = false;

                DialogResult result = folderDialog.ShowDialog();
                if (result != DialogResult.OK)
                {
                    mainWindow.LogInfo("Folder selection cancelled by user");
                    return;
                }

                mainWindow.txtBox_Browse_RenameInvoice.Text = folderDialog.SelectedPath;
                mainWindow.LogInfo($"Selected folder: {folderDialog.SelectedPath}");

                if (!LoadFilesFromFolder(folderDialog.SelectedPath, out string[] xmlFilesFound, out string[] pdfFilesFound))
                {
                    mainWindow.LogError("No XML or PDF files found in the selected folder");
                    mainWindow.ShowError("No XML or PDF files found in the selected folder.");
                    mainWindow.label_TotalFiles_RenameInvoice.Content = "Files found: 0 XML / 0 PDF";
                    mainWindow.ComboBox_RenameInvoice.Items.Clear();
                    return;
                }

                mainWindow.XmlFiles = xmlFilesFound;
                mainWindow.PdfFiles = pdfFilesFound;
                mainWindow.PopulateComboBox(mainWindow.ComboBox_RenameInvoice, mainWindow.XmlFiles.Concat(mainWindow.PdfFiles).ToArray());
                mainWindow.label_TotalFiles_RenameInvoice.Content = $"Files found: {mainWindow.XmlFiles.Length} XML / {mainWindow.PdfFiles.Length} PDF";
                mainWindow.UpdateProgressBar(0, 1, "");

                mainWindow.LogSuccess($"Successfully loaded {mainWindow.XmlFiles.Length} XML and {mainWindow.PdfFiles.Length} PDF files");
                mainWindow.ShowInfo($"Loaded {mainWindow.XmlFiles.Length} XML and {mainWindow.PdfFiles.Length} PDF files from the selected folder.");
            }
        }

        public void OpenFile_Click()
        {
            if (string.IsNullOrWhiteSpace(mainWindow.txtBox_Browse_RenameInvoice.Text) || mainWindow.ComboBox_RenameInvoice.SelectedItem == null)
            {
                mainWindow.LogWarning("No folder or file selected for opening");
                mainWindow.ShowError("Please select a folder and a file.");
                return;
            }

            string folderPath = mainWindow.txtBox_Browse_RenameInvoice.Text;
            string fileName = mainWindow.ComboBox_RenameInvoice.SelectedItem.ToString();
            string filePath = Path.Combine(folderPath, fileName);

            if (!File.Exists(filePath))
            {
                mainWindow.LogError($"Selected file does not exist: {fileName}");
                mainWindow.ShowError("Selected file does not exist.");
                return;
            }

            string extension = Path.GetExtension(fileName).ToLower();
            switch (extension)
            {
                case ".xml":
                    mainWindow.ShowXMLContent(filePath, fileName);
                    break;
                case ".pdf":
                    mainWindow.OpenFileWithDefaultProgram(filePath);
                    break;
                default:
                    mainWindow.LogError($"Unsupported file type: {extension}");
                    mainWindow.ShowError("Unsupported file type.");
                    break;
            }
        }

        public async Task RenameAll_Click()
        {
            if (mainWindow.IsProcessing)
            {
                mainWindow.LogWarning("Rename process already in progress - request ignored");
                mainWindow.ShowInfo("Processing is already in progress. Please wait for it to complete.");
                return;
            }

            if (IsNoFilesLoaded())
            {
                mainWindow.LogError("No files loaded for renaming process");
                mainWindow.ShowError("No files loaded. Please select a folder with XML/PDF files first.");
                return;
            }

            mainWindow.LogInfo("Starting rename process - selecting destination folder");

            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select folder to store renamed files";
                folderDialog.ShowNewFolderButton = true;

                if (folderDialog.ShowDialog() != DialogResult.OK)
                {
                    mainWindow.LogInfo("Destination folder selection cancelled");
                    return;
                }

                mainWindow.FolderDestination = folderDialog.SelectedPath;
                mainWindow.LogInfo($"Selected destination folder: {mainWindow.FolderDestination}");

                if (!PrepareDestinationFolders())
                {
                    return;
                }

                await ProcessRenameOperation();
            }
        }

        private bool IsNoFilesLoaded()
        {
            return mainWindow.XmlFiles?.Length == 0 && mainWindow.PdfFiles?.Length == 0;
        }

        private bool PrepareDestinationFolders()
        {
            string renamedFolder = Path.Combine(mainWindow.FolderDestination, "Renamed");
            string failedFolder = Path.Combine(mainWindow.FolderDestination, "Renamed_failed");

            try
            {
                CleanupExistingFolders(renamedFolder, failedFolder);
                CreateDestinationFolders(renamedFolder, failedFolder);
                return true;
            }
            catch (Exception ex)
            {
                mainWindow.LogError($"Error creating destination folders: {ex.Message}");
                mainWindow.ShowError($"Error creating folders: {ex.Message}");
                return false;
            }
        }

        private void CleanupExistingFolders(string renamedFolder, string failedFolder)
        {
            if (Directory.Exists(renamedFolder))
            {
                Directory.Delete(renamedFolder, true);
                mainWindow.LogInfo("Removed existing 'Renamed' folder");
            }

            if (Directory.Exists(failedFolder))
            {
                Directory.Delete(failedFolder, true);
                mainWindow.LogInfo("Removed existing 'Renamed_failed' folder");
            }
        }

        private void CreateDestinationFolders(string renamedFolder, string failedFolder)
        {
            _ = Directory.CreateDirectory(renamedFolder);
            _ = Directory.CreateDirectory(failedFolder);

            mainWindow.LogSuccess("Created folders: Renamed and Renamed_failed");
            mainWindow.ShowInfo($"Folders created successfully:\n- {renamedFolder}\n- {failedFolder}", "Folders Created");
        }

        private async Task ProcessRenameOperation()
        {
            string renamedFolder = Path.Combine(mainWindow.FolderDestination, "Renamed");
            string failedFolder = Path.Combine(mainWindow.FolderDestination, "Renamed_failed");

            mainWindow.CancellationTokenSource = new CancellationTokenSource();
            SetProcessingState(true);

            mainWindow.LogInfo($"Process started at: {mainWindow.ProcessStartTime:yyyy-MM-dd HH:mm:ss}");

            try
            {
                await ProcessFilesForRenamingAsync(renamedFolder, failedFolder, mainWindow.CancellationTokenSource.Token);
            }
            catch (OperationCanceledException)
            {
                mainWindow.LogWarning("Rename operation was cancelled by user");
                mainWindow.UpdateProgressBar(0, 1, "Operation Cancelled");
                mainWindow.ShowInfo("Rename operation was cancelled.", "Operation Cancelled");
            }
            finally
            {
                FinishProcessing();
            }
        }

        private void SetProcessingState(bool isProcessing)
        {
            mainWindow.btn_Renameall_RenameInvoice.IsEnabled = !isProcessing;
            mainWindow.btn_Stop_RenameInvoice.IsEnabled = isProcessing;
            mainWindow.ExportXml_Tab.Visibility = isProcessing ? Visibility.Collapsed : Visibility.Visible;
            mainWindow.IsProcessing = isProcessing;

            if (isProcessing)
            {
                mainWindow.ProcessStartTime = DateTime.Now;
                mainWindow.UpdateLogCounters();
            }
        }

        private void FinishProcessing()
        {
            DateTime processEndTime = DateTime.Now;
            TimeSpan totalTime = processEndTime - mainWindow.ProcessStartTime;

            SetProcessingState(false);
            mainWindow.UpdateLogCounters();

            mainWindow.LogInfo($"Process ended at: {processEndTime:yyyy-MM-dd HH:mm:ss}");
            mainWindow.LogSuccess($"Rename process completed - Total time: {mainWindow.FormatTimeSpan(totalTime)}");

            mainWindow.CancellationTokenSource?.Dispose();
            mainWindow.CancellationTokenSource = null;
        }

        public void Stop_Click()
        {
            if (mainWindow.CancellationTokenSource?.Token.IsCancellationRequested == false)
            {
                mainWindow.LogWarning("Stop button clicked - cancelling rename operation");
                mainWindow.CancellationTokenSource.Cancel();
                mainWindow.btn_Stop_RenameInvoice.IsEnabled = false;
                mainWindow.UpdateProgressBar(0, 1, "Stopping...");
            }
        }

        public void OpenFolder_Click()
        {
            if (string.IsNullOrWhiteSpace(mainWindow.FolderDestination) || !Directory.Exists(mainWindow.FolderDestination))
            {
                mainWindow.LogWarning("No valid destination folder to open");
                mainWindow.ShowError("No valid destination folder to open. Please complete a rename operation first.");
                return;
            }

            try
            {
                mainWindow.LogInfo($"Opening destination folder: {mainWindow.FolderDestination}");
                _ = Process.Start(new ProcessStartInfo
                {
                    FileName = mainWindow.FolderDestination,
                    UseShellExecute = true,
                    Verb = "open"
                });
                mainWindow.LogSuccess("Successfully opened destination folder");
            }
            catch (Exception ex)
            {
                mainWindow.LogError($"Error opening destination folder: {ex.Message}");
                mainWindow.ShowError($"Error opening destination folder: {ex.Message}");
            }
        }

        private bool LoadFilesFromFolder(string folderPath, out string[] xmlFilesFound, out string[] pdfFilesFound)
        {
            mainWindow.LogInfo($"Scanning folder: {folderPath}");

            xmlFilesFound = Directory.GetFiles(folderPath, "*.xml", SearchOption.TopDirectoryOnly);
            pdfFilesFound = Directory.GetFiles(folderPath, "*.pdf", SearchOption.TopDirectoryOnly);

            mainWindow.LogInfo($"Found {xmlFilesFound.Length} XML files and {pdfFilesFound.Length} PDF files");

            if (xmlFilesFound.Length == 0 && pdfFilesFound.Length == 0)
            {
                mainWindow.LogWarning("No XML or PDF files found in the selected folder");
            }

            return xmlFilesFound.Length > 0 || pdfFilesFound.Length > 0;
        }

        public async Task ProcessFilesForRenamingAsync(string renamedFolder, string failedFolder, CancellationToken cancellationToken)
        {
            int totalFiles = (mainWindow.XmlFiles?.Length ?? 0) + (mainWindow.PdfFiles?.Length ?? 0);

            if (totalFiles == 0)
            {
                mainWindow.LogWarning("No files to process");
                mainWindow.ShowInfo("No files to process.", "Information");
                return;
            }

            mainWindow.LogInfo($"Processing {totalFiles} total files ({mainWindow.XmlFiles?.Length ?? 0} XML, {mainWindow.PdfFiles?.Length ?? 0} PDF)");
            mainWindow.UpdateProgressBar(0, totalFiles, "Starting...");

            ProcessingResults results = new ProcessingResults();

            try
            {
                await ProcessXmlFiles(renamedFolder, failedFolder, cancellationToken, results);
                await ProcessPdfFiles(renamedFolder, failedFolder, cancellationToken, results);

                _ = ShowFinalResults(results, totalFiles);
            }
            catch (OperationCanceledException)
            {
                mainWindow.LogWarning("Processing operation was cancelled");
                throw;
            }
            catch (Exception ex)
            {
                mainWindow.LogError($"Critical error during processing: {ex.Message}");
                mainWindow.ShowError($"Error during processing: {ex.Message}");
            }
            finally
            {
                mainWindow.UpdateProgressBar(0, 1, "");
            }
        }

        private async Task ProcessXmlFiles(string renamedFolder, string failedFolder, CancellationToken cancellationToken, ProcessingResults results)
        {
            if (mainWindow.XmlFiles?.Length > 0)
            {
                cancellationToken.ThrowIfCancellationRequested();

                DateTime xmlStartTime = DateTime.Now;
                mainWindow.LogInfo($"Starting XML processing: {mainWindow.XmlFiles.Length} files");
                mainWindow.UpdateProgressBar(0, (mainWindow.XmlFiles?.Length ?? 0) + (mainWindow.PdfFiles?.Length ?? 0), "Processing XML files...");

                (int successCount, int failedCount) = await ProcessXmlFilesAsync(mainWindow.XmlFiles, renamedFolder, failedFolder, cancellationToken);
                results.XmlSuccessCount = successCount;
                results.XmlFailedCount = failedCount;

                TimeSpan xmlProcessTime = DateTime.Now - xmlStartTime;
                mainWindow.LogInfo($"XML processing completed: {successCount} success, {failedCount} failed - Time: {mainWindow.FormatTimeSpan(xmlProcessTime)}");
            }
        }

        private async Task ProcessPdfFiles(string renamedFolder, string failedFolder, CancellationToken cancellationToken, ProcessingResults results)
        {
            if (mainWindow.PdfFiles?.Length > 0)
            {
                cancellationToken.ThrowIfCancellationRequested();

                DateTime pdfStartTime = DateTime.Now;
                mainWindow.LogInfo($"Starting PDF processing: {mainWindow.PdfFiles.Length} files");
                mainWindow.UpdateProgressBar(mainWindow.XmlFiles?.Length ?? 0, (mainWindow.XmlFiles?.Length ?? 0) + mainWindow.PdfFiles.Length, "Processing PDF files...");

                (int successCount, int failedCount) = await ProcessPdfFilesBatchAsync(mainWindow.PdfFiles, renamedFolder, failedFolder, cancellationToken);
                results.PdfSuccessCount = successCount;
                results.PdfFailedCount = failedCount;

                TimeSpan pdfProcessTime = DateTime.Now - pdfStartTime;
                mainWindow.LogInfo($"PDF processing completed: {successCount} success, {failedCount} failed - Time: {mainWindow.FormatTimeSpan(pdfProcessTime)}");
            }
        }

        private async Task ShowFinalResults(ProcessingResults results, int totalFiles)
        {
            mainWindow.UpdateProgressBar(totalFiles, totalFiles, "Processing Complete!");
            await Task.Delay(1000);

            int totalSuccess = results.XmlSuccessCount + results.PdfSuccessCount;
            int totalFailed = results.XmlFailedCount + results.PdfFailedCount;

            mainWindow.LogSuccess($"All processing completed - Total Success: {totalSuccess}, Total Failed: {totalFailed}");

            mainWindow.ShowInfo($"Processing completed:\n" +
                    $"XML Files - Success: {results.XmlSuccessCount}, Failed: {results.XmlFailedCount}\n" +
                    $"PDF Files - Success: {results.PdfSuccessCount}, Failed: {results.PdfFailedCount}\n" +
                    $"Total - Success: {totalSuccess}, Failed: {totalFailed}",
                    "Rename Process Complete");
        }

        private async Task<(int successCount, int failedCount)> ProcessXmlFilesAsync(string[] xmlFilePaths, string renamedFolder, string failedFolder, CancellationToken cancellationToken)
        {
            int successCount = 0;
            int failedCount = 0;

            foreach (string xmlFilePath in xmlFilePaths)
            {
                cancellationToken.ThrowIfCancellationRequested();

                string fileName = Path.GetFileName(xmlFilePath);
                mainWindow.LogInfo($"Processing XML file: {fileName}");
                mainWindow.UpdateProgressBar(successCount + failedCount, xmlFilePaths.Length + (mainWindow.PdfFiles?.Length ?? 0), $"Processing XML: {fileName}");

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
                        mainWindow.LogError($"Exception processing XML file '{fileName}': {ex.Message}");
                        CopyToFailedFolder(xmlFilePath, failedFolder, $"Processing error: {ex.Message}");
                        return false;
                    }
                }, cancellationToken);

                if (success)
                {
                    successCount++;
                    mainWindow.LogSuccess($"Successfully processed XML: {fileName}");
                }
                else
                {
                    failedCount++;
                    mainWindow.LogError($"Failed to process XML: {fileName}");
                }

                await Task.Delay(50, cancellationToken);
            }

            return (successCount, failedCount);
        }

        private async Task<(int successCount, int failedCount)> ProcessPdfFilesBatchAsync(string[] pdfFilePaths, string renamedFolder, string failedFolder, CancellationToken cancellationToken)
        {
            try
            {
                mainWindow.LogInfo("Starting PDF batch processing with Python script");

                (string pythonExe, string scriptPath) = GetPythonPaths();
                if (!ValidatePythonEnvironment(pythonExe, scriptPath))
                {
                    return (0, pdfFilePaths.Length);
                }

                List<string> arguments = BuildPythonArguments(scriptPath, pdfFilePaths, renamedFolder, failedFolder);
                ProcessStartInfo psi = CreateProcessStartInfo(pythonExe, arguments);

                mainWindow.LogInfo($"Executing Python command: {pythonExe} {psi.Arguments}");

                return await ExecutePythonScript(psi, pdfFilePaths, cancellationToken);
            }
            catch (OperationCanceledException)
            {
                mainWindow.LogWarning("PDF processing was cancelled");
                throw;
            }
            catch (Exception ex)
            {
                mainWindow.LogError($"Error executing Python batch script: {ex.Message}");
                mainWindow.ShowError($"Error processing PDF files: {ex.Message}");
                return (0, pdfFilePaths.Length);
            }
        }

        private (string pythonExe, string scriptPath) GetPythonPaths()
        {
            string pythonExe = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "portable/python", "python.exe");
            string scriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "portable/app", "pdf_rename.py");
            return (pythonExe, scriptPath);
        }

        private bool ValidatePythonEnvironment(string pythonExe, string scriptPath)
        {
            if (!File.Exists(pythonExe))
            {
                mainWindow.LogError($"Python executable not found: {pythonExe}");
                mainWindow.ShowError($"Python executable not found: {pythonExe}");
                return false;
            }

            if (!File.Exists(scriptPath))
            {
                mainWindow.LogError($"Python script not found: {scriptPath}");
                mainWindow.ShowError($"Python script not found: {scriptPath}");
                return false;
            }

            return true;
        }

        private List<string> BuildPythonArguments(string scriptPath, string[] pdfFilePaths, string renamedFolder, string failedFolder)
        {
            List<string> arguments = new List<string> { $"\"{scriptPath}\"", "-i" };
            arguments.AddRange(pdfFilePaths.Select(path => $"\"{path}\""));
            arguments.AddRange(new[] { "-s", $"\"{renamedFolder}\"", "-f", $"\"{failedFolder}\"" });
            return arguments;
        }

        private ProcessStartInfo CreateProcessStartInfo(string pythonExe, List<string> arguments)
        {
            return new ProcessStartInfo
            {
                FileName = pythonExe,
                Arguments = string.Join(" ", arguments),
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
        }

        private async Task<(int successCount, int failedCount)> ExecutePythonScript(ProcessStartInfo psi, string[] pdfFilePaths, CancellationToken cancellationToken)
        {
            return await Task.Run(() =>
            {
                using (Process process = Process.Start(psi))
                {
                    PythonProcessResults results = new PythonProcessResults();
                    int baseProgress = mainWindow.XmlFiles?.Length ?? 0;

                    ProcessPythonOutput(process, results, baseProgress, pdfFilePaths.Length, cancellationToken);

                    string errors = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    if (!string.IsNullOrEmpty(errors))
                    {
                        mainWindow.LogError($"Python script errors: {errors}");
                        Debug.WriteLine($"Python script errors: {errors}");
                    }

                    mainWindow.LogInfo($"Python script completed with exit code: {process.ExitCode}");
                    return (results.SuccessCount, results.FailedCount);
                }
            }, cancellationToken);
        }

        private void ProcessPythonOutput(Process process, PythonProcessResults results, int baseProgress, int totalPdfFiles, CancellationToken cancellationToken)
        {
            string line;
            while ((line = process.StandardOutput.ReadLine()) != null)
            {
                cancellationToken.ThrowIfCancellationRequested();
                mainWindow.LogInfo($"Python output: {line}");

                ProcessProgressLine(line, results, baseProgress, totalPdfFiles);
                ProcessSummaryLine(line, results);
            }
        }

        private void ProcessProgressLine(string line, PythonProcessResults results, int baseProgress, int totalPdfFiles)
        {
            if (!line.StartsWith("PROGRESS:"))
            {
                return;
            }

            results.ProcessedCount++;
            string fileName = ExtractFileNameFromProgressLine(line);

            if (line.Contains("SUCCESS"))
            {
                results.SuccessCount++;
                UpdateProgressSuccess(baseProgress, results.ProcessedCount, totalPdfFiles, fileName);
            }
            else if (line.Contains("FAILED") || line.Contains("ERROR"))
            {
                results.FailedCount++;
                UpdateProgressFailure(baseProgress, results.ProcessedCount, totalPdfFiles, fileName);
            }
        }

        private void UpdateProgressSuccess(int baseProgress, int processedCount, int totalPdfFiles, string fileName)
        {
            mainWindow.Dispatcher.Invoke(() =>
            {
                mainWindow.UpdateProgressBar(baseProgress + processedCount,
                    (mainWindow.XmlFiles?.Length ?? 0) + totalPdfFiles,
                    $"Completed PDF: {fileName}");
                mainWindow.LogSuccess($"PDF processed successfully: {fileName}");
            });
        }

        private void UpdateProgressFailure(int baseProgress, int processedCount, int totalPdfFiles, string fileName)
        {
            mainWindow.Dispatcher.Invoke(() =>
            {
                mainWindow.UpdateProgressBar(baseProgress + processedCount,
                    (mainWindow.XmlFiles?.Length ?? 0) + totalPdfFiles,
                    $"Failed PDF: {fileName}");
                mainWindow.LogError($"PDF processing failed: {fileName}");
            });
        }

        private void ProcessSummaryLine(string line, PythonProcessResults results)
        {
            if (!line.StartsWith("SUMMARY:"))
            {
                return;
            }

            Match summaryMatch = Regex.Match(line, @"SUCCESS=(\d+), FAILED=(\d+)");
            if (summaryMatch.Success)
            {
                results.SuccessCount = int.Parse(summaryMatch.Groups[1].Value);
                results.FailedCount = int.Parse(summaryMatch.Groups[2].Value);
                mainWindow.LogInfo($"Python script summary: {results.SuccessCount} success, {results.FailedCount} failed");
            }
        }

        private string ExtractFileNameFromProgressLine(string progressLine)
        {
            Match match = Regex.Match(progressLine, @"PROGRESS:\s+(.+?)\s+->");
            return match.Success ? match.Groups[1].Value : "Unknown file";
        }

        private bool ProcessSingleXmlFile(string xmlFilePath, string renamedFolder, string failedFolder)
        {
            try
            {
                (string sHDon, string nLap) = ExtractXmlData(xmlFilePath);
                if (string.IsNullOrEmpty(sHDon) || string.IsNullOrEmpty(nLap))
                {
                    mainWindow.LogWarning($"Missing SHDon or NLap data in file: {Path.GetFileName(xmlFilePath)}");
                    CopyToFailedFolder(xmlFilePath, failedFolder, "Missing SHDon or NLap data");
                    return false;
                }

                if (!ValidateAndFormatDate(nLap, out string datePrefix))
                {
                    mainWindow.LogWarning($"Invalid date format in NLap for file: {Path.GetFileName(xmlFilePath)} (Date: {nLap})");
                    CopyToFailedFolder(xmlFilePath, failedFolder, "Invalid date format in NLap");
                    return false;
                }

                return CopyRenamedFile(xmlFilePath, renamedFolder, datePrefix, sHDon);
            }
            catch (Exception ex)
            {
                mainWindow.LogError($"Error processing XML file {Path.GetFileName(xmlFilePath)}: {ex.Message}");
                CopyToFailedFolder(xmlFilePath, failedFolder, "General processing error");
                return false;
            }
        }

        private (string sHDon, string nLap) ExtractXmlData(string xmlFilePath)
        {
            string xmlContent = File.ReadAllText(xmlFilePath);
            xmlContent = CleanXmlString(xmlContent);
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlContent);

            string sHDon = doc.SelectSingleNode("//SHDon")?.InnerText?.Trim();
            string nLap = doc.SelectSingleNode("//NLap")?.InnerText?.Trim();

            return (sHDon, nLap);
        }

        private bool ValidateAndFormatDate(string nLap, out string datePrefix)
        {
            datePrefix = null;
            if (!DateTime.TryParse(nLap, out DateTime parsedDate))
            {
                return false;
            }

            datePrefix = parsedDate.ToString("yyMMdd");
            return true;
        }

        private bool CopyRenamedFile(string xmlFilePath, string renamedFolder, string datePrefix, string sHDon)
        {
            string newFileName = $"{datePrefix}_{sHDon}.xml";
            string finalFileName = GetUniqueFileName(renamedFolder, newFileName);
            string destinationPath = Path.Combine(renamedFolder, finalFileName);
            File.Copy(xmlFilePath, destinationPath, false);

            mainWindow.LogInfo($"Renamed {Path.GetFileName(xmlFilePath)} to {finalFileName}");
            return true;
        }

        private string CleanXmlString(string xmlContent)
        {
            return new string(xmlContent.Where(c =>
                c == 0x9 || c == 0xA || c == 0xD ||
                (c >= 0x20 && c <= 0xD7FF) ||
                (c >= 0xE000 && c <= 0xFFFD)
            ).ToArray());
        }

        private void CopyToFailedFolder(string sourceFilePath, string failedFolder, string reason)
        {
            try
            {
                string fileName = Path.GetFileName(sourceFilePath);
                string uniqueFailedPath = GetUniqueFileName(failedFolder, fileName);
                string failedFilePath = Path.Combine(failedFolder, uniqueFailedPath);
                File.Copy(sourceFilePath, failedFilePath, false);

                mainWindow.LogWarning($"Copied to failed folder: {fileName} (Reason: {reason})");
            }
            catch (Exception ex)
            {
                mainWindow.LogError($"Error copying to failed folder: {ex.Message}");
                mainWindow.ShowError($"Error copying to failed folder: {ex.Message}");
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

        private class ProcessingResults
        {
            public int XmlSuccessCount { get; set; }
            public int XmlFailedCount { get; set; }
            public int PdfSuccessCount { get; set; }
            public int PdfFailedCount { get; set; }
        }

        private class PythonProcessResults
        {
            public int SuccessCount { get; set; }
            public int FailedCount { get; set; }
            public int ProcessedCount { get; set; }
        }
    }
}