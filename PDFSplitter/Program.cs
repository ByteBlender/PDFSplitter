using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HelperLibrary;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace PDFSplitter
{
    class Program
    {

        static void Main(string[] args)
        {
            Console.WriteLine("1.ExtractIDs\n2.PdfRename\n3.MatchIDs\n4.VerifyFileNames");
            string sw = Console.ReadLine();

            switch (sw)
            {
                case "1":
                    ExtractIDs();
                    break;
                case "2":
                    PdfRename();
                    break;
                case "3":
                    MatchIDs();
                    break;
                case "4":
                    VerifyFileNames();
                    break;
                default:
                    break;
            }
            // MatchIDs();
            // PdfRename();
            // ExtractIDs();
            // VerifyFileNames();

            //string sourcePdfPath = @"\\MS-WS-HP01\Shared\Sample.pdf";
            //DataTable dt = TextFileRW.readTextFileToTable(@"\\MS-WS-HP01\Shared\FileNames.txt", "\t");

            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    string fileName = dt.Rows[i]["FileName"].ToString();
            //    string path = @"\\MS-WS-HP01\Shared\";
            //    SplitAndSaveInterval(sourcePdfPath, path, i+1, 1, fileName);

            //}
        }

        public static void SplitAndSaveInterval(string pdfFilePath, string outputPath, int startPage, int interval, string pdfFileName)
        {
            using (PdfReader reader = new PdfReader(pdfFilePath))
            {
                Document document = new Document();
                PdfCopy copy = new PdfCopy(document, new FileStream(outputPath + "\\" + pdfFileName + ".pdf", FileMode.Create));
                document.Open();

                for (int pagenumber = startPage; pagenumber < (startPage + interval); pagenumber++)
                {
                    if (reader.NumberOfPages >= pagenumber)
                    {
                        copy.AddPage(copy.GetImportedPage(reader, pagenumber));
                    }
                    else
                    {
                        break;
                    }
                }
                document.Close();
            }
        }

        public static void ExtractIDs()
        {
            PdfReader reader = new PdfReader(@"S:\DATABASES\AAA MISCELLANEOUS\MAIL MAKERS\JOB 207472 ECU RESULTS MAILING\PDF Extraction\Sem 2 2020 Main BoE Transcripts.pdf");
            FileStream fs = new FileStream(@"S:\DATABASES\AAA MISCELLANEOUS\MAIL MAKERS\JOB 207472 ECU RESULTS MAILING\PDF Extraction\extracted_test.txt", FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            PdfReaderContentParser parser = new PdfReaderContentParser(reader);
            ITextExtractionStrategy strategy;
            TextMarginFinder finder;
            string previousVal = "";
            string currentVal = "";
            int count = 0;

            sw.WriteLine("Index\tID\tPageNumber\tFileName");

            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                try
                {                  
                   // finder = parser.ProcessContent(i, new TextMarginFinder());
                    //Rectangle area = new Rectangle(finder.GetLlx(), finder.GetLly(), finder.GetWidth() / 2, finder.GetHeight() / 2);
                    Rectangle area = new Rectangle(414, 660, 522, 689);
                    RenderFilter filter = new RegionTextRenderFilter(area);
                    strategy = new FilteredTextRenderListener(new LocationTextExtractionStrategy(), filter);
                    currentVal = PdfTextExtractor.GetTextFromPage(reader, i, strategy);

                    if(previousVal != currentVal)
                    {
                        count = 0;
                    }
                    count++;
                    previousVal = currentVal;
                    sw.WriteLine($"{i}\t{currentVal}\t{count}\t{currentVal}-{count}");
                }
                catch (Exception)
                {
                    sw.WriteLine($"{i}\tfailed");
                }      
            }
            sw.Flush();
            sw.Close();
        }

        public static void MatchIDs()
        {
            DataTable source = TextFileRW.readTextFileToTable(@"S:\DATABASES\AAA MISCELLANEOUS\MAIL MAKERS\JOB 207024 ECU RESULTS MAILING\JOB207024DR - 201 Course Complete Address details by Date.txt", "\t");
            source.Columns.Add("PageCount");
            source.Columns.Add("Index");
            DataTable sample = TextFileRW.readTextFileToTable(@"S:\DATABASES\AAA MISCELLANEOUS\MAIL MAKERS\JOB 207024 ECU RESULTS MAILING\PDF Extraction\Page counts.txt", "\t");
            sample.Columns.Add("Matched");

            foreach (DataRow sourceRow in source.Rows)
            {
                foreach (DataRow sampleRow in sample.Rows)
                {
                    if(sourceRow["Person Id"].ToString() == sampleRow["ID"].ToString())
                    {
                        sourceRow["PageCount"] = sampleRow["PageCount"];
                        sourceRow["Index"] = sampleRow["Index"];
                        sampleRow["Matched"] = "True";
                    }
                }

            }

            TextFileRW.writeTableToTxtFile(source, @"S:\DATABASES\AAA MISCELLANEOUS\MAIL MAKERS\JOB 207024 ECU RESULTS MAILING\JOB207024DR - 201 Course Complete Address_ID_merged.txt", "\t");
            TextFileRW.writeTableToTxtFile(sample, @"S:\DATABASES\AAA MISCELLANEOUS\MAIL MAKERS\JOB 207024 ECU RESULTS MAILING\Transcripts List.txt", "\t");

        }

        public static void PdfRename()
        {


            string dir = @"S:\DATABASES\AAA MISCELLANEOUS\MAIL MAKERS\JOB 207024 ECU RESULTS MAILING\PDF Extraction\Transcripts\";

            string[] files = Directory.GetFiles(dir);
            DataTable dt = TextFileRW.readTextFileToTable(@"S:\DATABASES\AAA MISCELLANEOUS\MAIL MAKERS\JOB 207024 ECU RESULTS MAILING\PDF Extraction\extracted.txt", "\t");
            Dictionary<int, string> names = new Dictionary<int, string>();
            foreach (DataRow r in dt.Rows)
            {
                names.Add(int.Parse(r["Index"].ToString()), r["FileName"].ToString());
            }

            foreach (var file in files)
            {

                string f = System.IO.Path.GetFileName(file);
                int index = int.Parse(f.Replace("201 Course Completed Transcripts Version 2_CLEANED_Part", "").Replace(".pdf", ""));

                Directory.Move(file, $@"{dir}\{names[index]}.pdf");
            }

        }

        public static string ReadID(string fileName)
        {
            PdfReader reader = new PdfReader(fileName);
            PdfReaderContentParser parser = new PdfReaderContentParser(reader);
            ITextExtractionStrategy strategy;
            //TextMarginFinder finder;
            //finder = parser.ProcessContent(1, new TextMarginFinder());
            Rectangle area = new Rectangle(414, 660, 522, 689);
            RenderFilter filter = new RegionTextRenderFilter(area);
            strategy = new FilteredTextRenderListener(new LocationTextExtractionStrategy(), filter);
            return PdfTextExtractor.GetTextFromPage(reader, 1, strategy);    
        }

        public static void VerifyFileNames()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("FileName");
            dt.Columns.Add("ExtractedID");
            dt.Columns.Add("Match");

            string dir = @"S:\DATABASES\AAA MISCELLANEOUS\MAIL MAKERS\JOB 207024 ECU RESULTS MAILING\PDF Extraction\Transcripts\";
            string[] files = Directory.GetFiles(dir);
            string id = "";
            foreach (var file in files)
            {
                id = ReadID(file);
                DataRow r = dt.NewRow();
                r["FileName"] = file;
                r["ExtractedID"] = id;
                r["Match"] = System.IO.Path.GetFileNameWithoutExtension(file).Split('-')[0] == id ? "True" : "False";
                dt.Rows.Add(r);
            }

            TextFileRW.writeTableToTxtFile(dt, @"S:\DATABASES\AAA MISCELLANEOUS\MAIL MAKERS\JOB 207024 ECU RESULTS MAILING\PDF Extraction\FileNameVerification.txt", "\t");
        }

    }
}
