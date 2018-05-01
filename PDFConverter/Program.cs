using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.IO;
using Bytescout.PDFExtractor;
using System.Diagnostics;

namespace PDFConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            //SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();
            //f.OpenPdf(@"C:\Users\i.upadhyay\Desktop\AV42104576.pdf");

            //if (f.PageCount > 0)
            //{
            //    f.ToWord(@"C:\Users\i.upadhyay\Desktop\AV42104576.docx");
            //    Console.WriteLine("Finised");
            //    Console.ReadKey();
            //}
            //string fileName = "myfile.ext";
            //string path1 = @"mydir";
            //string path2 = @"\mydir";
            string fullPath;

            CSVExtractor extractor = new CSVExtractor();
            extractor.RegistrationName = "demo";
            extractor.RegistrationKey = "demo";

            // Load sample PDF document
            extractor.LoadDocumentFromFile("AV42104576.pdf");

            //extractor.CSVSeparatorSymbol = ","; // you can change CSV separator symbol (if needed) from "," symbol to another if needed for non-US locales

            extractor.SaveCSVToFile("output1.csv");

            Console.WriteLine();
            Console.WriteLine("Data has been extracted to 'output.csv' file.");
            Console.WriteLine();
            Console.WriteLine("Press any key to continue and open CSV in default CSV viewer (or Excel)...");
            Console.ReadKey();

            Process.Start("output.csv");


            //fullPath = System.IO.Path.GetFullPath(@"C:\Users\i.upadhyay\Desktop\AV42104576.pdf");
            //Program n = new Program();
            //n.ExportPDFToExcel(fullPath);
        }
      
        private void ExportPDFToExcel(string fileName)
        {
            StringBuilder text = new StringBuilder();
            PdfReader pdfReader = new PdfReader(fileName);
            for (int page = 1; page <= pdfReader.NumberOfPages; page++)
            {
                ITextExtractionStrategy strategy = new LocationTextExtractionStrategy();
                string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);
                currentText = Encoding.UTF8.GetString(Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.UTF8.GetBytes(currentText)));
                text.Append(currentText);
                pdfReader.Close();
            }
            Console.Clear();
           // Console.Buffer = true;
           // Console.AddHeader("content-disposition", "attachment;filename=ReceiptExport.xls");
          //  Console.Charset = "";
           // Console.ContentType = "application/vnd.ms-excel";
            Console.Write(text);
            //Console.Flush();
           // Console.End();
        }
    }
}
