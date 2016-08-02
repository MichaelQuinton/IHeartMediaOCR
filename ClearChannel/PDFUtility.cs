namespace ClearChannel
{
    using iTextSharp.text;
    using iTextSharp.text.pdf;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    internal static class PdfUtility
    {
        /// <summary>
        /// Create the overlay to write out overtop of the files actual contents.
        /// </summary>
        /// <param name="ocrText">Text returned from the LEADTools OCR engine.</param>
        /// <param name="stamper">itxtSharp object that contains file contents and allows overlays.</param>
        /// <returns>The canvas for overlay.</returns>
        private static void AddOverlayToCanvas(string ocrText, PdfStamper stamper)
        {
            var font = BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1250, false);
            var canvas = stamper.GetOverContent(1);
            var gs1 = new PdfGState
            {
                FillOpacity = 0.0f,
                StrokeOpacity = 0.0f
            };

            canvas.SetGState(gs1);
            canvas.SetFontAndSize(font, 1);
            canvas.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ocrText, 25, 3, 0);
            canvas.ShowTextAligned(PdfContentByte.ALIGN_LEFT, @"PAGE-1", 3, 3, 0);
        }

        /// <summary>
        /// Create a list of files ordered by their page counts.
        /// </summary>
        /// <param name="filePageCounts">Files and their associated page counts.</param>
        /// <returns>An ordered List of file paths.</returns>
        private static List<string> CreateOrderedList(Dictionary<string, int> filePageCounts)
        {
            var orderedList = new List<string>();
            var numberOfPages = 1;

            while (numberOfPages <= filePageCounts.Values.Max())
            {
                var pages = numberOfPages;
                var holder = from pair in filePageCounts
                             where pair.Value == pages
                             select pair.Key;

                orderedList.AddRange(holder.ToList());

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
            var outputFolder = Constants.OutputFolders[currentFolder.Key];
            var moveLocation = Path.Combine(Constants.OutputDirectory, outputFolder + @"\" + outputFolder + ".pdf");

            var orderedList = CreateOrderedList(filePageCounts);

            var fileToMove = MergeOrderedList(orderedList);

            File.Move(fileToMove, moveLocation);

            orderedList.Clear();
        }

        /// <summary>
        /// Takes an ordered list of files and merges them into one output file.
        /// </summary>
        /// <param name="mergeList">An ordered list of file paths.</param>
        /// <returns>The path to merged file created in this method.</returns>
        private static string MergeOrderedList(IEnumerable<string> mergeList)
        {
            var filename = Guid.NewGuid() + ".pdf";
            var outfileFullPath = Path.Combine(Constants.InputDirectory.FullName, filename);

            using (var stream = new FileStream(outfileFullPath, FileMode.Create))
            using (var doc = new Document())
            using (var pdf = new PdfCopy(doc, stream))
            {
                PdfReader reader = null;

                doc.Open();

                foreach (var file in mergeList)
                {
                    reader = new PdfReader(file);
                    for (var i = 0; i < reader.NumberOfPages; i++)
                    {
                        var page = pdf.GetImportedPage(reader, i + 1);
                        pdf.AddPage(page);
                    }
                }

                pdf.FreeReader(reader);
                reader?.Close();
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
            var reader = new PdfReader(filePath);
            var streamPdf = new MemoryStream();

            PdfReader.unethicalreading = true;

            reader = currentFolder.Key != "PremierLandscape" ? RotateLandscapePages(reader) : RotatePortraitPages(reader);

            using (var stamper = new PdfStamper(reader, streamPdf, '\0', true))
            {
                stamper.RotateContents = false;
                AddOverlayToCanvas(ocrText, stamper);
            }

            var pdf = streamPdf.ToArray();

            streamPdf.Dispose();
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
            var n = reader.NumberOfPages;

            for (var page = 0; page < n;)
            {
                ++page;

                var width = reader.GetPageSize(page).Width;
                var height = reader.GetPageSize(page).Height;

                if (!(width > height)) continue;
                var pageDict = reader.GetPageN(page);
                pageDict.Put(PdfName.ROTATE, new PdfNumber(90));
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
            var pageCount = reader.NumberOfPages;
            
            for (var p = 1; p <= pageCount; p++)
            {
                if (reader.GetPageSize(p).Width > reader.GetPageSize(p).Height) continue;
                var page = reader.GetPageN(p);
                
                var rotate = page.GetAsNumber(PdfName.ROTATE);
                page.Put(PdfName.ROTATE, rotate == null ? new PdfNumber(90) : new PdfNumber((rotate.IntValue + 90)%360));
            }
            return reader;
        }
    }
}