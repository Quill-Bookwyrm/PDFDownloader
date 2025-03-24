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


namespace PDF_Downloader
{
    class Program
    {
        private static string inputFilePath = "C:\\PDF\\Input\\GRI_2017_2020.xlsx";
        private string outputFilePath = "C:\\PDF\\Output";
        private string downloadedFilePath = "C:\\PDF\\Output\\down";

        static void Main(string[] args)
        {
            List<Document> docs = new List<Document>(GetDocuments());

            Console.Read();
        }
        public async Task DownloadDocumentsAsync(List<Document> documents)
        {
            //using (WebClient client = new WebClient())
            //{
            //    foreach (Document doc in documents)
            //    {
            //        client.DownloadFile(doc.Url, doc.BrNumber);

            //    }
            var httpClient = new HttpClient();
            foreach (Document doc in documents)
            {
                var responseStream = await httpClient.GetStreamAsync(doc.Url);
                var fileStream = new FileStream(outputFilePath, FileMode.Create);
                responseStream.CopyTo(fileStream);
            }
        }
        public static List<Document> GetDocuments()
        {
            List<Document> documents = new List<Document>();
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo($@"{inputFilePath}")))
            {
                
                var myWorksheet = xlPackage.Workbook.Worksheets.First();
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;
                int urlIndex = myWorksheet
                .Cells["1:1"]
                .First(c => c.Value.ToString() == "Pdf_URL")
                .Start
                .Column;
                int backupUrlIndex = urlIndex + 1;

                for (int rowNum = 1; rowNum <= totalRows; rowNum++)
                {
                    var brNum = myWorksheet.Cells[rowNum, 0, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                    var url = myWorksheet.Cells[rowNum, urlIndex, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                    var backupUrl = myWorksheet.Cells[rowNum, backupUrlIndex, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());

                    Document doc = new Document();
                    doc.BrNumber = string.Join("", brNum);
                    doc.Url = string.Join("", url);
                    if (backupUrl != null)
                    {
                        doc.BackupUrl = string.Join("", backupUrl);
                    }

                    documents.Add(doc);
                }
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
