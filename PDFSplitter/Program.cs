using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace PDFSplitter
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("File Name:");
            string sourcePdfPath = Console.ReadLine().Replace("\"", "");
            //string outputPdfPath = $"{sourcePdfPath.Replace(".pdf",)
            Console.Write("Start:");
            int start =int.Parse( Console.ReadLine());
            start = start * 6 - 5;
            Console.Write("End:");
            int end = int.Parse(Console.ReadLine());
            end = end * 6;
            ExtractPages(sourcePdfPath, start, end);

        }

        public static void ExtractPages(string sourcePdfPath,
    int startPage, int endPage)
        {
            string outputPdfPath = sourcePdfPath.Replace(".pdf",$"_pages { (startPage +5)/6}-{ endPage/6}.pdf");

            PdfReader reader = null;
            Document sourceDocument = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage = null;

            try
            {
                // Intialize a new PdfReader instance with the contents of the source Pdf file:
                reader = new PdfReader(sourcePdfPath);

                // For simplicity, I am assuming all the pages share the same size
                // and rotation as the first page:
                sourceDocument = new Document(reader.GetPageSizeWithRotation(startPage));

                // Initialize an instance of the PdfCopyClass with the source 
                // document and an output file stream:
                pdfCopyProvider = new PdfCopy(sourceDocument,
                    new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));

                sourceDocument.Open();

                // Walk the specified range and add the page copies to the output file:
                for (int i = startPage; i <= endPage; i++)
                {
                    importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                    pdfCopyProvider.AddPage(importedPage);
                }
                sourceDocument.Close();
                reader.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
