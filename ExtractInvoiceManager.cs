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
using System.Windows.Forms;
using System.Xml;

namespace Support_Accountant
{
    public class ExtractInvoiceController
    {
        private readonly MainWindow mainWindow;

        public ExtractInvoiceController(MainWindow window)
        {
            mainWindow = window;
        }

        public void Browse_Click()
        {
            mainWindow.LogInfo("Browse button clicked for extract invoice folder selection");

            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select XML Folder for Invoice Extraction";
                folderDialog.ShowNewFolderButton = false;

                DialogResult result = folderDialog.ShowDialog();
                if (result != DialogResult.OK)
                {
                    mainWindow.LogInfo("Folder selection cancelled by user");
                    return;
                }

                mainWindow.txtBox_Browse_ExtractInvoice.Text = folderDialog.SelectedPath;
                mainWindow.LogInfo($"Selected folder: {folderDialog.SelectedPath}");

                if (!LoadExtractFilesFromFolder(folderDialog.SelectedPath, out string[] xmlFilesFound))
                {
                    mainWindow.LogError("No XML files found in the selected folder");
                    mainWindow.ShowError("No XML files found in the selected folder.");
                    mainWindow.label_TotalFiles_ExtractInvoice.Content = "Files found: 0 XML";
                    mainWindow.ComboBox_ExtractInvoice.Items.Clear();
                    return;
                }

                mainWindow.ExtractXmlFiles = xmlFilesFound;
                mainWindow.PopulateComboBox(mainWindow.ComboBox_ExtractInvoice, mainWindow.ExtractXmlFiles);
                mainWindow.label_TotalFiles_ExtractInvoice.Content = $"Files found: {mainWindow.ExtractXmlFiles.Length} XML";
                mainWindow.UpdateProgressBar(0, 1, "");

                mainWindow.LogSuccess($"Successfully loaded {mainWindow.ExtractXmlFiles.Length} XML files for extraction");
                mainWindow.ShowInfo($"Loaded {mainWindow.ExtractXmlFiles.Length} XML files from the selected folder.");
            }
        }

        public void OpenFile_Click()
        {
            if (string.IsNullOrWhiteSpace(mainWindow.txtBox_Browse_ExtractInvoice.Text) || mainWindow.ComboBox_ExtractInvoice.SelectedItem == null)
            {
                mainWindow.LogWarning("No folder or file selected for opening");
                mainWindow.ShowError("Please select a folder and an XML file.");
                return;
            }

            string folderPath = mainWindow.txtBox_Browse_ExtractInvoice.Text;
            string fileName = mainWindow.ComboBox_ExtractInvoice.SelectedItem.ToString();
            string filePath = Path.Combine(folderPath, fileName);

            if (!File.Exists(filePath))
            {
                mainWindow.LogError($"Selected file does not exist: {fileName}");
                mainWindow.ShowError("Selected XML file does not exist.");
                return;
            }

            mainWindow.ShowXMLContent(filePath, fileName);
        }

        public async Task ExportSummary_Click()
        {
            if (mainWindow.IsProcessing)
            {
                mainWindow.LogWarning("Extract process already in progress - request ignored");
                mainWindow.ShowInfo("Processing is already in progress. Please wait for it to complete.");
                return;
            }

            if (mainWindow.ExtractXmlFiles?.Length == 0)
            {
                mainWindow.LogError("No XML files loaded for extraction process");
                mainWindow.ShowError("No XML files loaded. Please select a folder with XML files first.");
                return;
            }

            mainWindow.LogInfo("Starting invoice extraction process - selecting output Excel file");

            using (SaveFileDialog saveDialog = new SaveFileDialog())
            {
                saveDialog.Title = "Export Invoice Summary to Excel";
                saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                saveDialog.FileName = "invoice_summary.xlsx";

                if (saveDialog.ShowDialog() != DialogResult.OK)
                {
                    mainWindow.LogInfo("Excel file selection cancelled");
                    return;
                }

                mainWindow.LastExcelFilePath = saveDialog.FileName;
                mainWindow.LogInfo($"Selected output Excel file: {mainWindow.LastExcelFilePath}");

                await ProcessExport();
            }
        }

        private async Task ProcessExport()
        {
            mainWindow.CancellationTokenSource = new CancellationTokenSource();
            mainWindow.btn_ExportSummary_ExtractInvoice.IsEnabled = false;
            mainWindow.btn_Stop_ExtractInvoice.IsEnabled = true;
            mainWindow.rename_tab.Visibility = Visibility.Collapsed;
            mainWindow.IsProcessing = true;
            mainWindow.ProcessStartTime = DateTime.Now;
            mainWindow.UpdateLogCounters("ExtractInvoice");

            mainWindow.LogInfo($"Process started at: {mainWindow.ProcessStartTime:yyyy-MM-dd HH:mm:ss}");

            try
            {
                await ExportToExcelAsync(mainWindow.LastExcelFilePath, mainWindow.CancellationTokenSource.Token);
            }
            catch (OperationCanceledException)
            {
                mainWindow.LogWarning("Export operation was cancelled by user");
                mainWindow.UpdateProgressBar(0, 1, "Operation Cancelled");
                mainWindow.ShowInfo("Export operation was cancelled.", "Operation Cancelled");
            }
            finally
            {
                DateTime processEndTime = DateTime.Now;
                TimeSpan totalTime = processEndTime - mainWindow.ProcessStartTime;

                mainWindow.btn_ExportSummary_ExtractInvoice.IsEnabled = true;
                mainWindow.btn_Stop_ExtractInvoice.IsEnabled = false;
                mainWindow.rename_tab.Visibility = Visibility.Visible;
                mainWindow.IsProcessing = false;
                mainWindow.UpdateLogCounters("ExtractInvoice");
                mainWindow.LogInfo($"Process ended at: {processEndTime:yyyy-MM-dd HH:mm:ss}");
                mainWindow.LogSuccess($"Extract process completed - Total time: {mainWindow.FormatTimeSpan(totalTime)}");

                mainWindow.CancellationTokenSource?.Dispose();
                mainWindow.CancellationTokenSource = null;
            }
        }

        public void Stop_Click()
        {
            if (mainWindow.CancellationTokenSource != null && !mainWindow.CancellationTokenSource.Token.IsCancellationRequested)
            {
                mainWindow.LogWarning("Stop button clicked - cancelling extract operation");
                mainWindow.CancellationTokenSource.Cancel();
                mainWindow.btn_Stop_ExtractInvoice.IsEnabled = false;
                mainWindow.UpdateProgressBar(0, 1, "Stopping...");
            }
        }

        public void OpenExcel_Click()
        {
            if (string.IsNullOrWhiteSpace(mainWindow.LastExcelFilePath) || !File.Exists(mainWindow.LastExcelFilePath))
            {
                mainWindow.LogWarning("No valid Excel file to open");
                mainWindow.ShowError("No valid Excel file to open. Please complete an export operation first.");
                return;
            }

            try
            {
                mainWindow.LogInfo($"Opening Excel file: {mainWindow.LastExcelFilePath}");
                _ = Process.Start(new ProcessStartInfo
                {
                    FileName = mainWindow.LastExcelFilePath,
                    UseShellExecute = true
                });
                mainWindow.LogSuccess("Successfully opened Excel file");
            }
            catch (Exception ex)
            {
                mainWindow.LogError($"Error opening Excel file: {ex.Message}");
                mainWindow.ShowError($"Error opening Excel file: {ex.Message}");
            }
        }

        private bool LoadExtractFilesFromFolder(string folderPath, out string[] xmlFilesFound)
        {
            mainWindow.LogInfo($"Scanning folder for XML files: {folderPath}");

            xmlFilesFound = Directory.GetFiles(folderPath, "*.xml", SearchOption.TopDirectoryOnly);
            mainWindow.LogInfo($"Found {xmlFilesFound.Length} XML files");

            if (xmlFilesFound.Length == 0)
            {
                mainWindow.LogWarning("No XML files found in the selected folder");
            }

            return xmlFilesFound.Length > 0;
        }

        private async Task ExportToExcelAsync(string excelPath, CancellationToken cancellationToken)
        {
            mainWindow.LogInfo($"Starting Excel export to: {excelPath}");
            mainWindow.UpdateProgressBar(0, mainWindow.ExtractXmlFiles.Length, "Initializing...");

            try
            {
                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet summarySheet = package.Workbook.Worksheets.Add("Summary");
                    CreateDynamicSummaryHeaders(summarySheet);

                    ExcelWorksheet detailSheet = package.Workbook.Worksheets.Add("Detail");
                    CreateDetailSheetHeaders(detailSheet);

                    int summaryRow = 2;
                    int detailRow = 2;
                    int processedCount = 0;

                    Dictionary<string, Dictionary<string, (decimal amount, decimal tax)>> totalsByCurrency = new Dictionary<string, Dictionary<string, (decimal amount, decimal tax)>>();
                    Dictionary<string, (decimal beforeTax, decimal tax, decimal afterTax)> grandTotalsByCurrency = new Dictionary<string, (decimal beforeTax, decimal tax, decimal afterTax)>();

                    foreach (string xmlFile in mainWindow.ExtractXmlFiles)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        string fileName = Path.GetFileName(xmlFile);
                        mainWindow.LogInfo($"Processing XML file: {fileName}");

                        mainWindow.Dispatcher.Invoke(() =>
                        {
                            mainWindow.UpdateProgressBar(processedCount, mainWindow.ExtractXmlFiles.Length, $"Processing: {fileName}");
                        });

                        try
                        {
                            ProcessXmlFileForExtraction(package, summarySheet, xmlFile, ref summaryRow);
                            ProcessXmlFileForDetailSheetWithCurrencyTotals(detailSheet, xmlFile, ref detailRow, totalsByCurrency, grandTotalsByCurrency);
                            mainWindow.LogSuccess($"Successfully processed: {fileName}");
                        }
                        catch (OperationCanceledException)
                        {
                            mainWindow.LogWarning("Operation was cancelled during file processing");
                            throw;
                        }
                        catch (Exception ex)
                        {
                            mainWindow.LogError($"Error processing file '{fileName}': {ex.Message}");
                            continue;
                        }

                        processedCount++;

                        if (processedCount % 10 == 0)
                        {
                            await Task.Delay(50, cancellationToken);
                        }
                    }

                    AddDetailSheetSummaryByCurrency(detailSheet, totalsByCurrency, grandTotalsByCurrency, detailRow + 2);

                    mainWindow.Dispatcher.Invoke(() =>
                    {
                        mainWindow.UpdateProgressBar(mainWindow.ExtractXmlFiles.Length, mainWindow.ExtractXmlFiles.Length, "Saving Excel file...");
                    });

                    mainWindow.LogInfo("Finalizing Excel file...");

                    summarySheet.Cells[summarySheet.Dimension.Address].AutoFitColumns();
                    detailSheet.Cells[detailSheet.Dimension.Address].AutoFitColumns();
                    FileInfo fileInfo = new FileInfo(excelPath);
                    package.SaveAs(fileInfo);

                    mainWindow.Dispatcher.Invoke(() =>
                    {
                        mainWindow.UpdateProgressBar(mainWindow.ExtractXmlFiles.Length, mainWindow.ExtractXmlFiles.Length, "Export Complete!");
                        mainWindow.ShowInfo($"Excel export completed successfully!\nProcessed: {processedCount} files\nSaved to: {excelPath}", "Export Complete");
                    });

                    mainWindow.LogSuccess($"Excel export completed successfully: {processedCount} files processed");
                }
            }
            catch (OperationCanceledException)
            {
                mainWindow.LogWarning("Excel export was cancelled");
                throw;
            }
            catch (Exception ex)
            {
                mainWindow.LogError($"Critical error during Excel export: {ex.Message}");
                mainWindow.Dispatcher.Invoke(() =>
                {
                    mainWindow.ShowError($"Error exporting to Excel: {ex.Message}");
                });
                throw;
            }
            finally
            {
                mainWindow.Dispatcher.Invoke(() =>
                {
                    mainWindow.UpdateProgressBar(0, 1, "");
                });
            }
        }

        private void ProcessXmlFileForDetailSheetWithCurrencyTotals(ExcelWorksheet detailSheet, string xmlFile, ref int row,
                                                           Dictionary<string, Dictionary<string, (decimal amount, decimal tax)>> totalsByCurrency,
                                                           Dictionary<string, (decimal beforeTax, decimal tax, decimal afterTax)> grandTotalsByCurrency)
        {
            string xmlContent = File.ReadAllText(xmlFile);
            xmlContent = CleanXmlString(xmlContent);
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlContent);

            string sHDon = doc.SelectSingleNode("//SHDon")?.InnerText ?? "";
            string nLap = doc.SelectSingleNode("//NLap")?.InnerText ?? "";
            string currency = doc.SelectSingleNode("//DVTTe")?.InnerText ?? "";

            string sellerName = doc.SelectSingleNode("//NBan/Ten")?.InnerText ?? "";
            string sellerMST = doc.SelectSingleNode("//NBan/MST")?.InnerText ?? "";
            string sellerAddress = doc.SelectSingleNode("//NBan/DChi")?.InnerText ?? "";
            string buyerName = doc.SelectSingleNode("//NMua/Ten")?.InnerText ?? "";
            string buyerMST = doc.SelectSingleNode("//NMua/MST")?.InnerText ?? "";
            string buyerAddress = doc.SelectSingleNode("//NMua/DChi")?.InnerText ?? "";

            string fileName = Path.GetFileNameWithoutExtension(xmlFile);
            string sheetName = CreateSafeSheetName(fileName);

            XmlNodeList nodes = doc.SelectNodes("//HHDVu");
            int startRow = row;
            int itemCount = nodes.Count;

            foreach (XmlNode node in nodes)
            {
                PopulateDetailSheetRowWithHyperlink(detailSheet, node, sHDon, nLap, currency,
                                                   sellerName, sellerMST, sellerAddress,
                                                   buyerName, buyerMST, buyerAddress, row, sheetName,
                                                   row == startRow);
                row++;
            }

            if (itemCount > 1)
            {
                MergeCellsForMultipleItems(detailSheet, startRow, itemCount);
            }

            InitializeCurrencyDictionaries(currency, totalsByCurrency, grandTotalsByCurrency);
            ProcessInvoiceTotals(doc, currency, totalsByCurrency, grandTotalsByCurrency);
        }

        private void MergeCellsForMultipleItems(ExcelWorksheet detailSheet, int startRow, int itemCount)
        {
            detailSheet.Cells[startRow, 1, startRow + itemCount - 1, 1].Merge = true;
            detailSheet.Cells[startRow, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            detailSheet.Cells[startRow, 2, startRow + itemCount - 1, 2].Merge = true;
            detailSheet.Cells[startRow, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            if (mainWindow.checkBox_Seller2.IsChecked == true && mainWindow.checkBox_Buyer2.IsChecked == true)
            {
                for (int col = 4; col <= 9; col++)
                {
                    detailSheet.Cells[startRow, col, startRow + itemCount - 1, col].Merge = true;
                    detailSheet.Cells[startRow, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }
            }
            else if (mainWindow.checkBox_Seller2.IsChecked == true || mainWindow.checkBox_Buyer2.IsChecked == true)
            {
                for (int col = 4; col <= 6; col++)
                {
                    detailSheet.Cells[startRow, col, startRow + itemCount - 1, col].Merge = true;
                    detailSheet.Cells[startRow, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }
            }
        }

        private void InitializeCurrencyDictionaries(string currency,
                                                   Dictionary<string, Dictionary<string, (decimal amount, decimal tax)>> totalsByCurrency,
                                                   Dictionary<string, (decimal beforeTax, decimal tax, decimal afterTax)> grandTotalsByCurrency)
        {
            if (!totalsByCurrency.ContainsKey(currency))
            {
                totalsByCurrency[currency] = new Dictionary<string, (decimal amount, decimal tax)>();
            }

            if (!grandTotalsByCurrency.ContainsKey(currency))
            {
                grandTotalsByCurrency[currency] = (0, 0, 0);
            }
        }

        private void ProcessInvoiceTotals(XmlDocument doc, string currency,
                                         Dictionary<string, Dictionary<string, (decimal amount, decimal tax)>> totalsByCurrency,
                                         Dictionary<string, (decimal beforeTax, decimal tax, decimal afterTax)> grandTotalsByCurrency)
        {
            XmlNode tToanNode = doc.SelectSingleNode("//TToan");
            if (tToanNode == null)
            {
                return;
            }

            ProcessGrandTotals(tToanNode, currency, grandTotalsByCurrency);
            ProcessTaxRates(tToanNode, currency, totalsByCurrency);
        }

        private void ProcessGrandTotals(XmlNode tToanNode, string currency,
                                       Dictionary<string, (decimal beforeTax, decimal tax, decimal afterTax)> grandTotalsByCurrency)
        {
            (decimal beforeTax, decimal tax, decimal afterTax) = grandTotalsByCurrency[currency];

            if (decimal.TryParse(tToanNode.SelectSingleNode("TgTCThue")?.InnerText ?? "0", out decimal invoiceBeforeTax))
            {
                beforeTax += invoiceBeforeTax;
            }

            if (decimal.TryParse(tToanNode.SelectSingleNode("TgTThue")?.InnerText ?? "0", out decimal invoiceTax))
            {
                tax += invoiceTax;
            }

            if (decimal.TryParse(tToanNode.SelectSingleNode("TgTTTBSo")?.InnerText ?? "0", out decimal invoiceAfterTax))
            {
                afterTax += invoiceAfterTax;
            }

            grandTotalsByCurrency[currency] = (beforeTax, tax, afterTax);
        }

        private void ProcessTaxRates(XmlNode tToanNode, string currency,
                                    Dictionary<string, Dictionary<string, (decimal amount, decimal tax)>> totalsByCurrency)
        {
            XmlNodeList ltsuatNodes = tToanNode.SelectNodes("THTTLTSuat/LTSuat");
            foreach (XmlNode ltsuat in ltsuatNodes)
            {
                string taxRate = ltsuat.SelectSingleNode("TSuat")?.InnerText ?? "0%";
                if (decimal.TryParse(ltsuat.SelectSingleNode("ThTien")?.InnerText ?? "0", out decimal amount) &&
                    decimal.TryParse(ltsuat.SelectSingleNode("TThue")?.InnerText ?? "0", out decimal tax))
                {
                    if (totalsByCurrency[currency].ContainsKey(taxRate))
                    {
                        (decimal amount, decimal tax) existing = totalsByCurrency[currency][taxRate];
                        totalsByCurrency[currency][taxRate] = (existing.amount + amount, existing.tax + tax);
                    }
                    else
                    {
                        totalsByCurrency[currency][taxRate] = (amount, tax);
                    }
                }
            }
        }

        private void AddDetailSheetSummaryByCurrency(ExcelWorksheet detailSheet,
                                                     Dictionary<string, Dictionary<string, (decimal amount, decimal tax)>> totalsByCurrency,
                                                     Dictionary<string, (decimal beforeTax, decimal tax, decimal afterTax)> grandTotalsByCurrency,
                                                     int startRow)
        {
            mainWindow.LogInfo("Adding currency-based summary section to Detail sheet");

            detailSheet.Cells[startRow, 1].Value = "TỔNG KẾT THEO TỪNG LOẠI TIỀN TỆ";
            detailSheet.Cells[startRow, 1].Style.Font.Bold = true;
            detailSheet.Cells[startRow, 1].Style.Font.Size = 14;
            startRow += 2;

            int currentRow = startRow;

            foreach (KeyValuePair<string, Dictionary<string, (decimal amount, decimal tax)>> currencyGroup in totalsByCurrency.OrderBy(x => x.Key))
            {
                string currency = currencyGroup.Key;
                Dictionary<string, (decimal amount, decimal tax)> taxRateTotals = currencyGroup.Value;

                detailSheet.Cells[currentRow, 1].Value = $"TIỀN TỆ: {currency}";
                detailSheet.Cells[currentRow, 1].Style.Font.Bold = true;
                detailSheet.Cells[currentRow, 1].Style.Font.Size = 12;
                currentRow += 1;

                currentRow = AddTaxBreakdownTable(detailSheet, taxRateTotals, currency, currentRow);
                currentRow = AddCurrencyTotals(detailSheet, grandTotalsByCurrency, currency, currentRow);
            }

            ExcelRange summaryRange = detailSheet.Cells[startRow - 2, 1, currentRow - 1, 3];
            summaryRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            summaryRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            summaryRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            summaryRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            mainWindow.LogInfo("Currency-based summary section added successfully to Detail sheet");
        }

        private int AddTaxBreakdownTable(ExcelWorksheet detailSheet, Dictionary<string, (decimal amount, decimal tax)> taxRateTotals, string currency, int currentRow)
        {
            detailSheet.Cells[currentRow, 1].Value = "Thành tiền";
            detailSheet.Cells[currentRow, 2].Value = "Thuế suất";
            detailSheet.Cells[currentRow, 3].Value = "Tiền thuế";
            detailSheet.Cells[currentRow, 4].Value = "Đơn vị tiền tệ";

            for (int i = 1; i <= 4; i++)
            {
                detailSheet.Cells[currentRow, i].Style.Font.Bold = true;
            }

            currentRow++;

            foreach (KeyValuePair<string, (decimal amount, decimal tax)> kvp in taxRateTotals.OrderBy(x => x.Key))
            {
                detailSheet.Cells[currentRow, 1].Value = FormatDecimalString(kvp.Value.amount.ToString());
                detailSheet.Cells[currentRow, 2].Value = kvp.Key;
                detailSheet.Cells[currentRow, 3].Value = FormatDecimalString(kvp.Value.tax.ToString());
                detailSheet.Cells[currentRow, 4].Value = currency;
                currentRow++;
            }

            return currentRow;
        }

        private int AddCurrencyTotals(ExcelWorksheet detailSheet,
                                     Dictionary<string, (decimal beforeTax, decimal tax, decimal afterTax)> grandTotalsByCurrency,
                                     string currency, int currentRow)
        {
            if (grandTotalsByCurrency.ContainsKey(currency))
            {
                (decimal beforeTax, decimal tax, decimal afterTax) = grandTotalsByCurrency[currency];

                currentRow++;
                detailSheet.Cells[currentRow, 1].Value = "Tổng cộng (chưa thuế):";
                detailSheet.Cells[currentRow, 2].Value = FormatDecimalString(beforeTax.ToString());
                detailSheet.Cells[currentRow, 3].Value = currency;
                detailSheet.Cells[currentRow, 1].Style.Font.Bold = true;

                currentRow++;
                detailSheet.Cells[currentRow, 1].Value = "Tổng tiền thuế:";
                detailSheet.Cells[currentRow, 2].Value = FormatDecimalString(tax.ToString());
                detailSheet.Cells[currentRow, 3].Value = currency;
                detailSheet.Cells[currentRow, 1].Style.Font.Bold = true;

                currentRow++;
                detailSheet.Cells[currentRow, 1].Value = "Tổng cộng (đã thuế):";
                detailSheet.Cells[currentRow, 2].Value = FormatDecimalString(afterTax.ToString());
                detailSheet.Cells[currentRow, 3].Value = currency;
                detailSheet.Cells[currentRow, 1].Style.Font.Bold = true;
            }

            return currentRow + 2;
        }

        private void CreateDetailSheetHeaders(ExcelWorksheet detailSheet)
        {
            List<string> headers = new List<string>
            {
                "Số Hóa Đơn", "Ngày Lập", "STT"
            };

            if (mainWindow.checkBox_Seller2.IsChecked == true)
            {
                headers.AddRange(new[] { "Tên Người Bán", "MST Người Bán", "Địa Chỉ Người Bán" });
            }
            if (mainWindow.checkBox_Buyer2.IsChecked == true)
            {
                headers.AddRange(new[] { "Tên Người Mua", "MST Người Mua", "Địa Chỉ Người Mua" });
            }

            headers.AddRange(new[] {"THHDVu (Tên hàng hóa/dịch vụ)",
                "DVTinh (Đơn vị tính)", "SLuong (Số lượng)", "DGia (Đơn giá)",
                "ThTien (Tiền trước thuế)", "TSuat (Thuế suất)", "TgTien (Tiền sau thuế)", "DVTTe (Đơn vị tiền tệ)"});

            for (int i = 0; i < headers.Count; i++)
            {
                detailSheet.Cells[1, i + 1].Value = headers[i];
                detailSheet.Cells[1, i + 1].Style.Font.Bold = true;
            }

            mainWindow.LogInfo($"Created Detail sheet headers with {headers.Count} columns");
        }

        private void PopulateDetailSheetRowWithHyperlink(ExcelWorksheet detailSheet, XmlNode node, string sHDon, string nLap,
                                               string currency, string sellerName, string sellerMST, string sellerAddress,
                                               string buyerName, string buyerMST, string buyerAddress, int row, string sheetName,
                                               bool addHyperlink)
        {
            int col = 1;
            if (addHyperlink)
            {
                detailSheet.Cells[row, col].Hyperlink = new ExcelHyperLink($"'{sheetName}'!A1", sHDon);
                detailSheet.Cells[row, col].Value = sHDon;
                detailSheet.Cells[row, col].Style.Font.UnderLine = true;
            }
            else
            {
                detailSheet.Cells[row, col].Value = sHDon;
            }
            col++;

            detailSheet.Cells[row, col++].Value = nLap;
            detailSheet.Cells[row, col++].Value = node.SelectSingleNode("STT")?.InnerText ?? "";

            if (mainWindow.checkBox_Seller2.IsChecked == true)
            {
                detailSheet.Cells[row, col++].Value = sellerName;
                detailSheet.Cells[row, col++].Value = sellerMST;
                detailSheet.Cells[row, col++].Value = sellerAddress;
            }

            if (mainWindow.checkBox_Buyer2.IsChecked == true)
            {
                detailSheet.Cells[row, col++].Value = buyerName;
                detailSheet.Cells[row, col++].Value = buyerMST;
                detailSheet.Cells[row, col++].Value = buyerAddress;
            }

            detailSheet.Cells[row, col++].Value = node.SelectSingleNode("THHDVu")?.InnerText ?? "";
            detailSheet.Cells[row, col++].Value = node.SelectSingleNode("DVTinh")?.InnerText ?? "";
            detailSheet.Cells[row, col++].Value = FormatDecimalString(node.SelectSingleNode("SLuong")?.InnerText ?? "");

            string dGia = node.SelectSingleNode("DGia")?.InnerText ?? "";
            string thTien = node.SelectSingleNode("ThTien")?.InnerText ?? "";
            string tSuat = node.SelectSingleNode("TSuat")?.InnerText ?? "";

            detailSheet.Cells[row, col++].Value = string.IsNullOrEmpty(dGia) ? "" : $"{FormatDecimalString(dGia)} {currency}";
            detailSheet.Cells[row, col++].Value = string.IsNullOrEmpty(thTien) ? "" : $"{FormatDecimalString(thTien)} {currency}";
            detailSheet.Cells[row, col++].Value = string.IsNullOrEmpty(tSuat) ? "" : FormatDecimalString(tSuat);

            string tgTien = CalculateAfterTaxAmount(thTien, tSuat, currency);
            detailSheet.Cells[row, col++].Value = tgTien;
            detailSheet.Cells[row, col++].Value = currency;
        }

        private string CalculateAfterTaxAmount(string thTien, string tSuat, string currency)
        {
            if (string.IsNullOrEmpty(thTien) || string.IsNullOrEmpty(tSuat))
            {
                return "";
            }

            if (decimal.TryParse(thTien, out decimal thTienValue) && decimal.TryParse(tSuat.Replace("%", ""), out decimal taxRate))
            {
                decimal taxAmount = thTienValue * (taxRate / 100);
                decimal afterTaxAmount = thTienValue + taxAmount;
                return $"{FormatDecimalString(afterTaxAmount.ToString())} {currency}";
            }

            return "";
        }

        private void CreateDynamicSummaryHeaders(ExcelWorksheet summarySheet)
        {
            List<string> headers = new List<string>
            {
                "Tên Sheet", "Số Hóa Đơn", "Ngày Lập"
            };

            AddConditionalHeaders(headers);

            for (int i = 0; i < headers.Count; i++)
            {
                summarySheet.Cells[1, i + 1].Value = headers[i];
                summarySheet.Cells[1, i + 1].Style.Font.Bold = true;
            }

            mainWindow.LogInfo($"Created dynamic headers with {headers.Count} columns");
        }

        private void AddConditionalHeaders(List<string> headers)
        {
            if (mainWindow.checkBox_Seller.IsChecked == true)
            {
                headers.AddRange(new[] { "Tên Người Bán", "MST Người Bán", "Địa Chỉ Người Bán" });
            }
            if (mainWindow.checkBox_Buyer.IsChecked == true)
            {
                headers.AddRange(new[] { "Tên Người Mua", "MST Người Mua", "Địa Chỉ Người Mua" });
            }
            if (mainWindow.checkBox_TongSL.IsChecked == true)
            {
                headers.Add("Tổng Số Lượng");
            }
            if (mainWindow.checkBox_TongTien.IsChecked == true)
            {
                headers.Add("Tổng Tiền");
            }
            if (mainWindow.checkBox_TienThue.IsChecked == true)
            {
                headers.Add("Tiền Thuế");
            }
            if (mainWindow.checkBox_ThanhTien.IsChecked == true)
            {
                headers.Add("Thành Tiền");
            }
            if (mainWindow.checkBox_Currency.IsChecked == true)
            {
                headers.Add("Đơn Vị Tiền Tệ");
            }
        }

        private void ProcessXmlFileForExtraction(ExcelPackage package, ExcelWorksheet summarySheet, string xmlFile, ref int row)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlFile);

            string fileName = Path.GetFileNameWithoutExtension(xmlFile);
            string sheetName = CreateSafeSheetName(fileName);

            ExcelWorksheet fileSheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName);
            if (fileSheet == null)
            {
                fileSheet = package.Workbook.Worksheets.Add(sheetName);
                CreateDetailSheet(fileSheet, doc);
            }

            List<object> values = ExtractDynamicSummaryData(doc, fileName, sheetName);
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
            List<string> detailHeaders = new List<string>
            {
                "STT", "THHDVu (Tên hàng hóa/dịch vụ)", "DVTinh (Đơn vị tính)",
                "SLuong (Số lượng)", "DGia (Đơn giá)", "ThTien (Tiền trước thuế)",
                "TSuat (Thuế suất)", "TgTien (Tiền sau thuế)", "DVTTe (Đơn vị tiền tệ)"
            };

            if (mainWindow.checkBox_Seller2.IsChecked == true)
            {
                detailHeaders.AddRange(new[] { "Tên Người Bán", "MST Người Bán", "Địa Chỉ Người Bán" });
            }
            if (mainWindow.checkBox_Buyer2.IsChecked == true)
            {
                detailHeaders.AddRange(new[] { "Tên Người Mua", "MST Người Mua", "Địa Chỉ Người Mua" });
            }

            for (int i = 0; i < detailHeaders.Count; i++)
            {
                sheet.Cells[1, i + 1].Value = detailHeaders[i];
                sheet.Cells[1, i + 1].Style.Font.Bold = true;
            }

            XmlNodeList nodes = doc.SelectNodes("//HHDVu");
            int detailRow = 2;
            string currency = doc.SelectSingleNode("//DVTTe")?.InnerText ?? "";

            foreach (XmlNode node in nodes)
            {
                PopulateDetailRowWithAfterTax(sheet, node, doc, currency, detailRow);
                detailRow++;
            }

            AddSummaryTable(sheet, doc, detailRow + 2);
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
        }

        private void PopulateDetailRowWithAfterTax(ExcelWorksheet sheet, XmlNode node, XmlDocument doc, string currency, int row)
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

            string tgTien = CalculateAfterTaxAmount(thTien, tSuat, currency);
            sheet.Cells[row, col++].Value = tgTien;
            sheet.Cells[row, col++].Value = currency;

            if (mainWindow.checkBox_Seller2.IsChecked == true)
            {
                sheet.Cells[row, col++].Value = doc.SelectSingleNode("//NBan/Ten")?.InnerText ?? "";
                sheet.Cells[row, col++].Value = doc.SelectSingleNode("//NBan/MST")?.InnerText ?? "";
                sheet.Cells[row, col++].Value = doc.SelectSingleNode("//NBan/DChi")?.InnerText ?? "";
            }
            if (mainWindow.checkBox_Buyer2.IsChecked == true)
            {
                sheet.Cells[row, col++].Value = doc.SelectSingleNode("//NMua/Ten")?.InnerText ?? "";
                sheet.Cells[row, col++].Value = doc.SelectSingleNode("//NMua/MST")?.InnerText ?? "";
                sheet.Cells[row, col++].Value = doc.SelectSingleNode("//NMua/DChi")?.InnerText ?? "";
            }
        }

        private void AddSummaryTable(ExcelWorksheet sheet, XmlDocument doc, int startRow)
        {
            XmlNode tToanNode = doc.SelectSingleNode("//TToan");
            if (tToanNode == null)
            {
                return;
            }

            sheet.Cells[startRow, 1].Value = "Thành tiền";
            sheet.Cells[startRow, 2].Value = "Thuế suất";
            sheet.Cells[startRow, 3].Value = "Tiền thuế";

            for (int i = 1; i <= 3; i++)
            {
                sheet.Cells[startRow, i].Style.Font.Bold = true;
            }

            XmlNodeList ltsuatNodes = tToanNode.SelectNodes("THTTLTSuat/LTSuat");
            int tRow = startRow + 1;
            foreach (XmlNode ltsuat in ltsuatNodes)
            {
                sheet.Cells[tRow, 1].Value = FormatDecimalString(ltsuat.SelectSingleNode("ThTien")?.InnerText ?? "");
                sheet.Cells[tRow, 2].Value = ltsuat.SelectSingleNode("TSuat")?.InnerText ?? "";
                sheet.Cells[tRow, 3].Value = FormatDecimalString(ltsuat.SelectSingleNode("TThue")?.InnerText ?? "");
                tRow++;
            }

            int summaryRow = tRow + 1;
            (string, string)[] summaryData = new[]
            {
                ("Tổng cộng (chưa thuế):", tToanNode.SelectSingleNode("TgTCThue")?.InnerText ?? ""),
                ("Tổng tiền thuế:", tToanNode.SelectSingleNode("TgTThue")?.InnerText ?? ""),
                ("Tổng cộng (đã thuế):", tToanNode.SelectSingleNode("TgTTTBSo")?.InnerText ?? ""),
                ("Bằng chữ:", tToanNode.SelectSingleNode("TgTTTBChu")?.InnerText ?? "")
            };

            for (int i = 0; i < summaryData.Length; i++)
            {
                sheet.Cells[summaryRow + i, 1].Value = summaryData[i].Item1;
                sheet.Cells[summaryRow + i, 2].Value = i < 3 ? FormatDecimalString(summaryData[i].Item2) : summaryData[i].Item2;
                sheet.Cells[summaryRow + i, 1].Style.Font.Bold = true;
            }
        }

        private List<object> ExtractDynamicSummaryData(XmlDocument doc, string fileName, string sheetName)
        {
            List<object> values = new List<object>
            {
                sheetName,
                doc.SelectSingleNode("//SHDon")?.InnerText ?? "",
                doc.SelectSingleNode("//NLap")?.InnerText ?? ""
            };

            AddConditionalSummaryData(doc, values);
            return values;
        }

        private void AddConditionalSummaryData(XmlDocument doc, List<object> values)
        {
            if (mainWindow.checkBox_Seller.IsChecked == true)
            {
                values.Add(doc.SelectSingleNode("//NBan/Ten")?.InnerText ?? "");
                values.Add(doc.SelectSingleNode("//NBan/MST")?.InnerText ?? "");
                values.Add(doc.SelectSingleNode("//NBan/DChi")?.InnerText ?? "");
            }

            if (mainWindow.checkBox_Buyer.IsChecked == true)
            {
                values.Add(doc.SelectSingleNode("//NMua/Ten")?.InnerText ?? "");
                values.Add(doc.SelectSingleNode("//NMua/MST")?.InnerText ?? "");
                values.Add(doc.SelectSingleNode("//NMua/DChi")?.InnerText ?? "");
            }

            if (mainWindow.checkBox_TongSL.IsChecked == true)
            {
                values.Add(doc.SelectNodes("//HHDVu/STT")?.Count ?? 0);
            }
            if (mainWindow.checkBox_TongTien.IsChecked == true)
            {
                values.Add(FormatDecimalString(doc.SelectSingleNode("//TgTCThue")?.InnerText ?? ""));
            }
            if (mainWindow.checkBox_TienThue.IsChecked == true)
            {
                values.Add(FormatDecimalString(doc.SelectSingleNode("//TgTThue")?.InnerText ?? ""));
            }
            if (mainWindow.checkBox_ThanhTien.IsChecked == true)
            {
                values.Add(FormatDecimalString(doc.SelectSingleNode("//TgTTTBSo")?.InnerText ?? ""));
            }
            if (mainWindow.checkBox_Currency.IsChecked == true)
            {
                values.Add(doc.SelectSingleNode("//DVTTe")?.InnerText ?? "");
            }
        }

        private string CleanXmlString(string xmlContent)
        {
            return new string(xmlContent.Where(c =>
                c == 0x9 || c == 0xA || c == 0xD ||
                (c >= 0x20 && c <= 0xD7FF) ||
                (c >= 0xE000 && c <= 0xFFFD)
            ).ToArray());
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
            return decimal.TryParse(value, out decimal result)
                ? result == Math.Truncate(result)
                    ? result.ToString("#,##0", CultureInfo.InvariantCulture)
                    : result.ToString("#,##0.###", CultureInfo.InvariantCulture)
                : value;
        }
    }
}