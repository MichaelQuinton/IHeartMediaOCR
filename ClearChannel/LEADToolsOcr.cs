using Leadtools;
using Leadtools.Codecs;
using Leadtools.Forms;
using Leadtools.Forms.Ocr;
using Leadtools.Pdf;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace ClearChannel
{
    class LEADToolsOcr
    {
        private static string assemblyDir = Path.GetDirectoryName(Assembly.GetAssembly(typeof(Program)).FullName);
        private static string ocrAdvantageRuntimeDir = Path.Combine(assemblyDir, "OcrRuntime");
        private static string ocrWorkingDir = Path.Combine(assemblyDir, "Temp");

        /// <summary>
        /// Uses an instance of the LEADTools OCR Engine to "read" the text
        /// in the pre-defined OCR region representing the mailing address
        /// area.  Is thread safe.
        /// </summary>
        /// <param name="ocrEngine"></param>
        /// <param name="document"></param>
        /// <returns></returns>
        private static string GetAddressBlockText(IOcrEngine ocrEngine, PDFDocument document)
        {
            string returnedText = null;

            using (var codecs = new RasterCodecs())
            using (var image = new RasterImage(document.GetPageImage(codecs, 1)))
            using (IOcrDocument ocrDocument = ocrEngine.DocumentManager.CreateDocument())
            using (IOcrPage ocrPage = ocrDocument.Pages.AddPage(image, null))
            {
                var myZone = new OcrZone();
                myZone.Bounds = new LogicalRectangle(0, 2, 4, 1.4, LogicalUnit.Inch);
                ocrPage.Zones.Add(myZone);
                ocrPage.Recognize(null);
                returnedText = ocrPage.GetText(0).ToUpper();
            }
            return returnedText;
        }

        /// <summary>
        /// Uses parallel processing to perform OCR on the mailing address region of the pdf files.
        /// Calls other methods to update the pdf files with that OCR information and builds a
        /// dictionary for file paths with associated page counts.
        /// </summary>
        /// <param name="inputFiles">List of files to be processed.</param>
        /// <param name="currentFolder">Active input directory.</param>
        /// <returns>Dictionary of file paths and associated page counts.</returns>
        internal static Dictionary<string, int> Process(List<string> inputFiles, KeyValuePair<string, string> currentFolder)
        {
            var filePageCounts = new ConcurrentDictionary<string, int>();

            SetupOcrWorkingDirectory();

            Parallel.ForEach(inputFiles, file =>
            {
                string returnedText = null;
                using (var document = new PDFDocument(file))
                {
                    filePageCounts.TryAdd(file.ToString(), document.Pages.Count);

                    using (IOcrEngine ocrEngine = OcrEngineManager.CreateEngine(OcrEngineType.Advantage, false))
                    {
                        ocrEngine.Startup(null, null, ocrWorkingDir, ocrAdvantageRuntimeDir);
                        ocrEngine.SpellCheckManager.SpellCheckEngine = OcrSpellCheckEngine.None;
                        returnedText = GetAddressBlockText(ocrEngine, document);
                        ocrEngine.Shutdown();
                    }
                }
                PdfUtility.OverlayOcrText(returnedText, file, currentFolder);
            }
                );
            var returnDictionary = filePageCounts.ToDictionary(kvp => kvp.Key,
                                                               kvp => kvp.Value);
            return returnDictionary;
        }

        /// <summary>
        /// Creates or Cleans a working directory for the OCR engine.
        /// </summary>
        private static void SetupOcrWorkingDirectory()
        {
            if (!Directory.Exists(ocrWorkingDir))
                Directory.CreateDirectory(ocrWorkingDir);
            else
            {
                DirectoryInfo tempDirInfo = new DirectoryInfo(ocrWorkingDir);
                tempDirInfo.DeleteAllContents();
            }
        }
    } 
}