using HtmlAgilityPack;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Net;
using System.Text;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System;

namespace Scrapper
{
    class Program
    {
        static void Main(string[] args)
        {
            string url = "http://bluepages.com.sa/Home/ClientDetail/";

            for (int i = 2500; i < 235000; i++)
            {
                string newurl = url + i;
                try
                {
                    var response = CallUrl(newurl).Result;
                    ParseHtml(response);
                }
                catch
                { 
                }
            }
           
        }
        private static async Task<string> CallUrl(string fullUrl)
        {
            HttpClient client = new HttpClient();
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            client.DefaultRequestHeaders.Accept.Clear();
            var response = client.GetStringAsync(fullUrl);
            return await response;
        }

        private static  void ParseHtml(string html)
        {
            string PhoneNumber = "";
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(html);
            var MainDiv = htmlDoc.DocumentNode.Descendants("div").ToList()
                    .Where(node => node.GetAttributeValue("class", "").Contains("ClientDetailFont")).ToList().FirstOrDefault();
            if (MainDiv != null)
            {
                var SubDiv = MainDiv.Descendants("div").ToList()
                        .Where(node => node.GetAttributeValue("class", "").Contains("col-lg-6")).ToList().ToList(); 
                if(SubDiv[1]!=null)
                {
                    var Span= SubDiv[1].Descendants("span").ToList()
                        .Where(node => node.GetAttributeValue("class", "").Contains("span-left")).FirstOrDefault();

                    if (Span != null) 
                    {
                         PhoneNumber = Span.InnerText;
                        //WriteSample("Test",PhoneNumber);
                        
                    }
                }
            }
             var ClientInfo = htmlDoc.DocumentNode.Descendants("span").ToList()
        .Where(node => node.GetAttributeValue("class", "").Contains("ClientInfoFont")).ToList();        


          
            string Info = "";
            foreach (var link in ClientInfo)
            {
                Info += link.InnerText + " , ";
            }
            Info += " PhoneNumber = " + PhoneNumber;
            AddLog(Info);
           

        }
        public static void AddLog(string Log)
        {

            #region write error log
            var fileName = "ErrorLogs.txt";
            String filepath = "D:\\";

            string dir = filepath + "\\";
            var filePath = Path.Combine(dir, fileName);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            //File.AppendAllText(filePath, "--- Entry Dated " + DateTime.Now.ToLongTimeString() + " ---" + Environment.NewLine);
            File.AppendAllText(filePath, Log);
            File.AppendAllText(filePath, Environment.NewLine);
            #endregion



        }
        //public static void WriteSample(string CompanyName="",string Phone="")
        //{
        //    Excel.Application excelApp = new Excel.Application();
        //    if (excelApp != null)
        //    {
        //        Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
        //        Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets.Add();

        //        excelWorksheet.Cells[1, 1] = CompanyName;
        //        excelWorksheet.Cells[1, 2] = Phone;
        //        //excelWorksheet.Cells[3, 1] = "Value3";
        //        //excelWorksheet.Cells[4, 1] = "Value4";

        //        excelApp.ActiveWorkbook.SaveAs(@"C:\SaudiCompany.xlsx", Excel.XlFileFormat.xlWorkbookNormal);

        //        excelWorkbook.Close();
        //        excelApp.Quit();

        //        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
        //        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
        //        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
        //        GC.Collect();
        //        GC.WaitForPendingFinalizers();
        //    }
        //}
    }
}
