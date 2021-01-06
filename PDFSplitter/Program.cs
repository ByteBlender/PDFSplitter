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
        static string jobDir = @"S:\DATABASES\AAA MISCELLANEOUS\MAIL MAKERS\JOB 207472 ECU RESULTS MAILING";
        static string sourcePDF = "ecur1070.2139644.pdf";
        static string dataFile = "JOB207472DR - Course Completed Semester 2 2020.txt";
        static void Main(string[] args)
        {
            // MoveFiles();
            VerifyOutput();


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
            PdfReader reader = new PdfReader($@"{jobDir}\PDF Extraction\temp\{sourcePDF}");
            FileStream fs = new FileStream($@"{jobDir}\PDF Extraction\temp\extractedIDs.txt", FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            PdfReaderContentParser parser = new PdfReaderContentParser(reader);
            ITextExtractionStrategy strategy;
            TextMarginFinder finder;
            string previousVal = "";
            string currentVal = "";
            int count = 0;
            string pages = "";

            sw.WriteLine("Index\tID\tPageCounter\tPageNumber\tFileName");

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

                    Rectangle area2 = new Rectangle(465, 565, 555, 635);
                    RenderFilter filter2 = new RegionTextRenderFilter(area2);
                    strategy = new FilteredTextRenderListener(new LocationTextExtractionStrategy(), filter2);
                    pages = PdfTextExtractor.GetTextFromPage(reader, i, strategy);


                    if (previousVal != currentVal)
                    {
                        count = 0;
                    }
                    count++;
                    previousVal = currentVal;
                    sw.WriteLine($"{i}\t{currentVal}\t{pages.Split('\n')[0]}\t{count}\t{currentVal}-{count}");
                }
                catch (Exception)
                {
                    sw.WriteLine($"{i}\tfailed");
                }      
            }
            sw.Flush();
            sw.Close();
        }

        public static void VerifyOutput()
        {
            PdfReader reader = new PdfReader($@"{jobDir}\207472 ECU Transcripts.pdf");
            FileStream fs = new FileStream($@"{jobDir}\PDF Extraction\temp\extractedIDs_test.txt", FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            PdfReaderContentParser parser = new PdfReaderContentParser(reader);
            ITextExtractionStrategy strategy;
            TextMarginFinder finder;
            string previousVal = "";
            string currentVal = "";
            int count = 0;
            string pages = "";

            sw.WriteLine("Index\tID\tPageCounter\tPageNumber\tFileName");

            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                try
                {
                    // finder = parser.ProcessContent(i, new TextMarginFinder());
                    //Rectangle area = new Rectangle(finder.GetLlx(), finder.GetLly(), finder.GetWidth() / 2, finder.GetHeight() / 2);
                    Rectangle area = new Rectangle(61, 600, 265, 680);
                    RenderFilter filter = new RegionTextRenderFilter(area);
                    strategy = new FilteredTextRenderListener(new LocationTextExtractionStrategy(), filter);
                    currentVal = PdfTextExtractor.GetTextFromPage(reader, i, strategy);

                    //Rectangle area2 = new Rectangle(465, 565, 555, 635);
                    //RenderFilter filter2 = new RegionTextRenderFilter(area2);
                    //strategy = new FilteredTextRenderListener(new LocationTextExtractionStrategy(), filter2);
                    //pages = PdfTextExtractor.GetTextFromPage(reader, i, strategy);


                    if (previousVal != currentVal)
                    {
                        count = 0;
                    }
                    count += 2;
                    previousVal = currentVal;
                    sw.WriteLine($"{i}\t{currentVal}\t{pages.Split('\n')[0]}\t{count}\t{currentVal}-{count}");
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
            DataTable source = TextFileRW.readTextFileToTable($@"{jobDir}\{dataFile}", "\t");
            source.Columns.Add("PageCount");
            source.Columns.Add("Index");
            DataTable sample = TextFileRW.readTextFileToTable($@"{jobDir}\PDF Extraction\PageCounts.txt", "\t");
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

            TextFileRW.writeTableToTxtFile(source, $@"{jobDir}\{dataFile.Replace(".txt","")}_merged.txt", "\t");
            TextFileRW.writeTableToTxtFile(sample, $@"{jobDir}\Transcripts List.txt", "\t");

        }

        public static void PdfRename()
        {


            string dir = $@"{jobDir}\PDF Extraction\Transcripts\";

            string[] files = Directory.GetFiles(dir);
            DataTable dt = TextFileRW.readTextFileToTable($@"{jobDir}\PDF Extraction\extractedIDs.txt", "\t");
            Dictionary<int, string> names = new Dictionary<int, string>();
            foreach (DataRow r in dt.Rows)
            {
                names.Add(int.Parse(r["Index"].ToString()), r["FileName"].ToString());
            }

            foreach (var file in files)
            {

                string f = System.IO.Path.GetFileName(file);
                int index = int.Parse(f.Replace($"{sourcePDF.Replace(".pdf","")}_Part", "").Replace(".pdf", ""));

                Directory.Move(file, $@"{dir}\{names[index]}.pdf");
            }

        }

        public static void MoveFiles()
        {


            string dir = $@"{jobDir}\PDF Extraction\Transcripts\";

            string[] files = Directory.GetFiles(dir);
            DataTable dt = TextFileRW.readTextFileToTable($@"{jobDir}\PDF Extraction\Bad Files.txt", "\t");
            Dictionary<int, string> names = new Dictionary<int, string>();
           

            foreach (var file in files)
            {

                string f = System.IO.Path.GetFileName(file);


                foreach (DataRow r in dt.Rows)
                {
                    if(file== r[0].ToString())
                    {
                        Directory.Move(file, $@"{dir}\bad files\{f}");
                    }
                }
        
            }

        }

        public static string ReadID(string fileName)
        {
            try
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
            catch (Exception)
            {

                return "unreadable";
            }
          
        }

        public static void VerifyFileNames()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("FileName");
            dt.Columns.Add("ExtractedID");
            dt.Columns.Add("Match");

            string dir = $@"{jobDir}\PDF Extraction\Transcripts\Fixed\New folder";
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

            TextFileRW.writeTableToTxtFile(dt, $@"{jobDir}\PDF Extraction\FileNameVerification.txt", "\t");
        }

    }
}
