using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ClearChannel
{
    class PdfUtility
    {
        /// <summary>
        /// Create the overlay to write out overtop of the files actual contents.
        /// </summary>
        /// <param name="ocrText">Text returned from the LEADTools OCR engine.</param>
        /// <param name="stamper">itxtSharp object that contains file contents and allows overlays.</param>
        /// <returns>The canvas for overlay.</returns>
        private static PdfContentByte AddOverlayToCanvas(string ocrText, PdfStamper stamper)
        {
            BaseFont font = BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1250, false);
            PdfContentByte canvas = stamper.GetOverContent(1);
            PdfGState gs1 = new PdfGState();

            gs1.FillOpacity = 0.0f;
            gs1.StrokeOpacity = 0.0f;

            canvas.SetGState(gs1);
            canvas.SetFontAndSize(font, 1);
            canvas.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ocrText, 25, 3, 0);
            canvas.ShowTextAligned(PdfContentByte.ALIGN_LEFT, @"PAGE-1", 3, 3, 0);

            return canvas;
        }

        /// <summary>
        /// Create a list of files ordered by their page counts.
        /// </summary>
        /// <param name="filePageCounts">Files and their associated page counts.</param>
        /// <returns>An ordered List of file paths.</returns>
        private static List<string> CreateOrderedList(Dictionary<string, int> filePageCounts)
        {
            List<string> orderedList = new List<string>();
            int numberOfPages = 1;

            while (numberOfPages <= filePageCounts.Values.Max())
            {
                var holder = from pair in filePageCounts
                             where pair.Value == numberOfPages
                             select pair.Key;

                foreach (string file in holder.ToList())
                    orderedList.Add(file);

                numberOfPages++;
            }
            return orderedList;
        }

        /// <summary>
        /// Merges an ordered list of file by page count order and moves the file
        /// to the correct output location.
        /// </summary>
        /// <param name="currentFolder">Current active input directory.</param>
        /// <param name="filePageCounts">Dictionary of file paths and corresponding page counts.</param>
        internal static void MergeAscendingPageCount(KeyValuePair<string, string> currentFolder, Dictionary<string, int> filePageCounts)
        {
            string outputFolder = Constants.outputFolders[currentFolder.Key.ToString()];
            string moveLocation = Path.Combine(Constants.outputDirectory, outputFolder + @"\" + outputFolder + ".pdf");

            List<string> orderedList = CreateOrderedList(filePageCounts);

            string fileToMove = MergeOrderedList(orderedList);

            File.Move(fileToMove, moveLocation);

            orderedList.Clear();
        }

        /// <summary>
        /// Takes an ordered list of files and merges them into one output file.
        /// </summary>
        /// <param name="mergeList">An ordered list of file paths.</param>
        /// <returns>The path to merged file created in this method.</returns>
        private static string MergeOrderedList(List<string> mergeList)
        {
            int counter = 0;
            string filename = Guid.NewGuid() + ".pdf";
            string outfileFullPath = Path.Combine(Constants.inputDirectory.FullName, filename);

            using (var stream = new FileStream(outfileFullPath, FileMode.Create))
            using (var doc = new Document())
            using (var pdf = new PdfCopy(doc, stream))
            {
                PdfReader reader = null;
                PdfImportedPage page = null;

                doc.Open();

                foreach (string file in mergeList)
                {
                    reader = new PdfReader(file);
                    for (int i = 0; i < reader.NumberOfPages; i++)
                    {
                        page = pdf.GetImportedPage(reader, i + 1);
                        pdf.AddPage(page);
                    }
                    counter++;
                }

                pdf.FreeReader(reader);
                reader.Close();
            }

            return outfileFullPath;
        }

        /// <summary>
        /// Reads in file contents, creates overlay of OCR text and marks first page before 
        /// writing back overtop of the oringinal file.  Rotates contents based on whether files
        /// are expected to be portrait or landscape.
        /// </summary>
        /// <param name="ocrText">The text return from the LEADTools OCR engine.</param>
        /// <param name="filePath">The current file path to create the overlay for.</param>
        /// <param name="currentFolder">Used to determine whether files should be portrait or landscape.</param>
        internal static void OverlayOcrText(string ocrText, string filePath, KeyValuePair<string, string> currentFolder)
        {
            byte[] pdf = null;
            var reader = new PdfReader(filePath);
            var streamPDF = new MemoryStream();

            PdfReader.unethicalreading = true;

            if (currentFolder.Key.ToString() != "PremierLandscape")
                reader = RotateLandscapePages(reader);
            else
                reader = RotatePortraitPages(reader);


            using (var stamper = new PdfStamper(reader, streamPDF, '\0', true))
            {
                stamper.RotateContents = false;
                PdfContentByte canvas = AddOverlayToCanvas(ocrText, stamper);
            }

            pdf = streamPDF.ToArray();

            streamPDF.Dispose();
            reader.Dispose();

            File.WriteAllBytes(filePath, pdf);
        }

        /// <summary>
        /// Rotates pages that should be in portrait orientation.
        /// </summary>
        /// <param name="reader">The object holding the current pdf.</param>
        /// <returns>The holding object after it may or may not have rotated pages.</returns>
        private static PdfReader RotateLandscapePages(PdfReader reader)
        {
            int n = reader.NumberOfPages;

            for (int page = 0; page < n;)
            {
                ++page;

                float width = reader.GetPageSize(page).Width;
                float height = reader.GetPageSize(page).Height;

                if (width > height)
                {
                    PdfDictionary pageDict = reader.GetPageN(page);
                    pageDict.Put(PdfName.ROTATE, new PdfNumber(90));
                }
            }
            return reader;
        }

        /// <summary>
        /// Rotates pages that should be in landscape orientation.
        /// </summary>
        /// <param name="reader">The object holding the current pdf.</param>
        /// <returns>The holding object after it may or may not have rotated pages.</returns>
        private static PdfReader RotatePortraitPages(PdfReader reader)
        {
            int n = reader.NumberOfPages;

            for (int page = 0; page < n;)
            {
                ++page;

                float width = reader.GetPageSize(page).Width;
                float height = reader.GetPageSize(page).Height;

                if (height > width)
                {
                    PdfDictionary pageDict = reader.GetPageN(page);
                    pageDict.Put(PdfName.ROTATE, new PdfNumber(90));
                }
            }
            return reader;
        }
    } 
}
