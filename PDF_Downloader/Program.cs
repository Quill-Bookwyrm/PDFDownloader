using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;
using System.Net;
using System.Net.Http;
using ClosedXML.Excel;
using System.ComponentModel.DataAnnotations;
using System.Threading;


namespace PDF_Downloader
{
    class Program
    {
        private static string inputFilePath = "C:\\PDF\\Input\\GRI_2017_2020.xlsx";
        private static string metaDataPath = "C:\\PDF\\Output\\meta\\MetaData.xlsx";
        private string outputFilePath = "C:\\PDF\\Output";
        private string downloadedFilePath = "C:\\PDF\\Output\\down";

        static void Main(string[] args)
        {
            List<Document> docs = new List<Document>(GetDocuments());
            docs = CheckMetadata(docs);

            Console.WriteLine("\nDownload finished, press any key to close the program");
            Console.Read();
        }
        
        public static List<Document> CheckMetadata(List<Document> documents)
        {
            List<Document> fixedDocuments = new List<Document>();
            if (!File.Exists(metaDataPath))
            {
                var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Meta Data");
                ws.Cell("A1").Value = "BRnum";
                ws.Cell("B1").Value = "Downloaded";
                ws.Cell("C1").Value = "Pdf_URL";
                ws.Cell("D1").Value = "Report Html Address";
                ws.Cell("A1").SetActive();
                var rangeTable = ws.Range(1, 1, 1, 4);
                rangeTable.FirstCell().Style
                    .Font.SetBold()
                    .Fill.SetBackgroundColor(XLColor.Aqua);

                int i = 2;
                foreach (Document doc in documents)
                {
                    ws.Cell($"A{i}").Value = doc.BrNumber;
                    ws.Cell($"B{i}").Value = "No";
                    ws.Cell($"C{i}").Value = doc.Url;
                    ws.Cell($"D{i}").Value = doc.BackupUrl;
                    i++;
                }

                ws.Columns().AdjustToContents(1, 4);
                wb.SaveAs("MetaData.xlsx");
                File.Move("MetaData.xlsx", metaDataPath);
            }
            else
            {
                List<int> brnums = new List<int>();
                var workbook = new XLWorkbook(metaDataPath);
                var ws = workbook.Worksheet("Meta Data");
                var totalRows = ws.RowsUsed().Count();
                for (int rowNum = 2; rowNum <= totalRows; rowNum++)
                {
                    string s;
                    s = ws.Cell($"A{rowNum}").Value
                        .ToString()
                        .Substring(2);
                    brnums.Add(Int32.Parse(s));
                }

                for (int i = 2; i <= totalRows; i++)
                {
                    foreach (Document doc in documents)
                    {
                        if (!brnums.Contains(Int32.Parse(doc.BrNumber.Substring(2))) && ws.Cell($"B{i}").Value.ToString().Contains("No"))
                        {
                            fixedDocuments.Add(doc);
                        }
                    }
                }
                return fixedDocuments;
            }

            return documents;
        }
        public async Task DownloadDocumentsAsync(List<Document> documents)
        {
            var httpClient = new HttpClient();
            var allTasks = new List<Task>();
            var throttler = new SemaphoreSlim(initialCount: 10);

            var wb = new XLWorkbook(metaDataPath);
            var ws = wb.Worksheet("Meta Data");
            var range = ws.RangeUsed();
            var table = range.AsTable();

            foreach (Document doc in documents)
            {
                await throttler.WaitAsync();

                allTasks.Add(
                    Task.Run(async () =>
                    {
                        try
                        {
                            if (!httpClient.GetStreamAsync(doc.Url).IsFaulted)
                            {
                                var responseStream = await httpClient.GetStreamAsync(doc.Url);
                                var fileStream = new FileStream(outputFilePath, FileMode.Create);
                                responseStream.CopyTo(fileStream);

                                var cell = table.FindRow(c => c.FirstCell().Value.ToString() == $"{doc.BrNumber}");
                                cell.Cell($"B{cell.RowNumber()}").Value = "Yes";
                            }
                            else if (!httpClient.GetStreamAsync(doc.BackupUrl).IsFaulted)
                            {
                                var responseStream = await httpClient.GetStreamAsync(doc.BackupUrl);
                                var fileStream = new FileStream(outputFilePath, FileMode.Create);
                                responseStream.CopyTo(fileStream);

                                var cell = table.FindRow(c => c.FirstCell().Value.ToString() == $"{doc.BrNumber}");
                                cell.Cell($"B{cell.RowNumber()}").Value = "Yes";
                            }
                            else
                            {
                                Console.WriteLine($"Failed to download {doc.BrNumber}");
                            }
                        }
                        finally
                        {
                            throttler.Release();
                        }
                    }));
            }
            await Task.WhenAll(allTasks);
        }
        public static List<Document> GetDocuments()
        {
            List<Document> documents = new List<Document>();
            using (XLWorkbook xlwb = new XLWorkbook(inputFilePath))
            {
                var myWorksheet = xlwb.Worksheet("0");
                var totalRows = myWorksheet.RowsUsed().Count();
                var range = myWorksheet.RangeUsed();
                var table = range.AsTable();

                for (int rowNum = 2; rowNum <= totalRows; rowNum++)
                {
                    var url = "";
                    var backupUrl = "";
                    Document doc = new Document();

                    var cellOne = table.FindColumn(c => c.FirstCell().Value.ToString() == "Pdf_URL");
                    if (cellOne != null)
                    {
                        var columnLetter = cellOne.RangeAddress.FirstAddress.ColumnLetter;
                        var brNum = myWorksheet.Cell(rowNum, 1).Value.ToString();
                        url = myWorksheet.Cell(rowNum, columnLetter).Value.ToString();
                        doc.Url = url;
                        doc.BrNumber = brNum;

                        var cellTwo = table.FindColumn(c => c.FirstCell().Value.ToString() == "Report Html Address");
                        if (cellTwo != null)
                        {
                            columnLetter = cellTwo.RangeAddress.FirstAddress.ColumnLetter;
                            backupUrl = myWorksheet.Cell(rowNum, columnLetter).Value.ToString();
                            doc.BackupUrl = backupUrl;
                        }
                    }

                    documents.Add(doc);
                }
                //var myWorksheet = xlPackage.Workbook.Worksheets.First();
                //var totalRows = myWorksheet.Dimension.End.Row;
                //var totalColumns = myWorksheet.Dimension.End.Column;
                //int urlIndex = myWorksheet
                //.Cells["1:1"]
                //.First(c => c.Value.ToString() == "Pdf_URL")
                //.Start
                //.Column;
                //int backupUrlIndex = urlIndex + 1;

                //for (int rowNum = 1; rowNum <= totalRows; rowNum++)
                //{
                //    var brNum = myWorksheet.Cells[rowNum, 0, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                //    var url = myWorksheet.Cells[rowNum, urlIndex, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                //    var backupUrl = myWorksheet.Cells[rowNum, backupUrlIndex, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());

                //    Document doc = new Document();
                //    doc.BrNumber = string.Join("", brNum);
                //    doc.Url = string.Join("", url);
                //    if (backupUrl != null)
                //    {
                //        doc.BackupUrl = string.Join("", backupUrl);
                //    }

                //    documents.Add(doc);
            }
            return documents;
        }
    }
    public class Document
    {
        string _brNumber;
        string _url;
        string _backupUrl;
        public Document(string brNumber, string url, string backupUrl)
        {
            _brNumber = brNumber;
            _url = url;
            _backupUrl = backupUrl;
        }
        public Document(string brNumber, string url)
        {
            _brNumber = brNumber;
            _url = url;
        }
        public Document()
        {
        }
        public string BrNumber { get => _brNumber; set => _brNumber = value; }
        public string Url { get => _url; set => _url = value; }
        public string BackupUrl { get => _backupUrl; set => _backupUrl = value; }
    }
}
