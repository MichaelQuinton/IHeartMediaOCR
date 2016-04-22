using ICSharpCode.SharpZipLib.Core;
using Leadtools;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Net.Mail;
using System.Threading.Tasks;
using SharpZip = ICSharpCode.SharpZipLib.Zip;

namespace ClearChannel
{
    class Program
    {
        private static List<string> inputFiles = new List<string>();
        private static Dictionary<string, int> filePageCounts = new Dictionary<string, int>();

        /// <summary>
        /// IHeart OCR and Merge Utility created by: Michael Quinton
        /// 
        /// Takes the groups of input PDF files and performs OCR on the
        /// mailing address.  OCR output is overlaid with a first page marking
        /// onto the files, then all files are merged in page count order.  The 
        /// resulting file is then moved to the appropriate drop folder.
        /// </summary>
        static void Main()
        {
            Console.WriteLine("IHeart Media OCR and Merge Utility");
            if (!CheckForInputFiles())
                return;

            SetLeadtoolsLicense();

            foreach (KeyValuePair<string, string> currentFolder in Constants.inputFolders)
            {
                Constants.inputDirectory.DeleteAllContents();
                GetInputFiles(currentFolder);

                if (Constants.inputDirectory.IsEmpty())
                    continue;

                LoadInputFileList();

                filePageCounts = LEADToolsOcr.Process(inputFiles, currentFolder);

                PdfUtility.MergeAscendingPageCount(currentFolder, filePageCounts);

                inputFiles.Clear();
                filePageCounts.Clear();

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            Constants.inputDirectory.DeleteAllContents();
            ArchiveInputFiles();
            DeleteOldErrorLogs();
            
            return;
        }

        /// <summary>
        /// Create a compressed folder with todays date.
        /// Copy zip files used for input into compressed folder.
        /// </summary>
        private static void ArchiveInputFiles()
        {
            string fileName = "Clear Channel " + DateTime.Now.ToString("MMM_d_yyyy") + ".zip";
            string zipArchivePath = Path.Combine(Constants.archiveDirectory, fileName);

            using (var stream = new FileStream(zipArchivePath, FileMode.Create))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Create))
                foreach (KeyValuePair<string, string> currentFolder in Constants.inputFolders)
                {
                    string filePath = Path.Combine(Constants.networkFTPDir, currentFolder.Value.ToString() + "/");
                    string[] files = Directory.GetFiles(filePath);

                    foreach (string file in files)
                    {
                        string zipPath = currentFolder.Key.ToString() + "/" + Path.GetFileName(file);
                        zip.CreateEntryFromFile(file, zipPath, CompressionLevel.Optimal);
                    }
                }
        }

        /// <summary>
        /// Checks network input locations for files.
        /// </summary>
        /// <returns>True if input files are present.</returns>
        private static bool CheckForInputFiles()
        {
            foreach (KeyValuePair<string, string> currentFolder in Constants.inputFolders)
            {
                var directory = new DirectoryInfo(Path.Combine(Constants.networkFTPDir, currentFolder.Value.ToString()));

                if (!directory.IsEmpty())
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Cleans up error logs created over a month ago.
        /// </summary>
        private static void DeleteOldErrorLogs()
        {
            foreach (var file in Constants.errorLogDirectory.GetFiles())
                if (file.CreationTime < DateTime.Now.AddMonths(-1))
                    file.Delete();
        }

        /// <summary>
        /// Uses the SharpZipLib to unzip a file
        /// </summary>
        /// <param name="archiveFilenameIn">Location of the file to be unzipped</param>
        /// <param name="outFolder">Where the unzipped files should be placed</param>
        /// <param name="password">Optional parameter to allow for the handling of AES encrypted files.</param>
        private static void ExtractZipFile(string archiveFilenameIn, string outFolder, string password = null)
        {
            try
            {
                using (FileStream fs = File.OpenRead(archiveFilenameIn))
                using (var zf = new SharpZip.ZipFile(fs))
                {
                    zf.IsStreamOwner = true;

                    if (!String.IsNullOrEmpty(password))
                        zf.Password = password;  //AES encrypted entries are handled automatically

                    foreach (SharpZip.ZipEntry zipEntry in zf)
                    {
                        if (!zipEntry.IsFile)
                            continue;   //Ignore Directories

                        String entryFileName = zipEntry.Name;

                        if (!entryFileName.EndsWith(".pdf", StringComparison.InvariantCultureIgnoreCase))
                            continue;

                        byte[] buffer = new byte[4096];
                        Stream zipStream = zf.GetInputStream(zipEntry);

                        String fullZipToPath = Path.Combine(outFolder, entryFileName);
                        string directoryName = Path.GetDirectoryName(fullZipToPath);

                        using (FileStream streamWriter = File.Create(fullZipToPath))
                            StreamUtils.Copy(zipStream, streamWriter, buffer);
                    }
                }
            }
            catch (Exception)
            {
                string errorMessage = archiveFilenameIn + " was unable to be opened. File is likely corrupted.";
                try
                {
                    SendCorruptFileEmail(errorMessage);
                }
                catch (Exception)
                {
                    LogErrorToFile(archiveFilenameIn, errorMessage);
                }


            }
        }
        
        /// <summary>
        /// Unzips PDF files into the local input folder.  If an error is encountered
        /// unzipping files from a zip file an email with the corrupted zip file info
        /// will be sent to the Laser Team.
        /// </summary>
        /// <param name="currentFolder">Current input folder.</param>
        private static void GetInputFiles(KeyValuePair<string, string> currentFolder)
        {
            string zipPath = Path.Combine(Constants.networkFTPDir, currentFolder.Value.ToString());
            string[] files = Directory.GetFiles(zipPath);

            Parallel.ForEach(files, file =>
            {
                ExtractZipFile(file, Constants.inputDirectory.FullName);
            });
        }

        /// <summary>
        /// Populate the inputFiles List with the file paths of files in the input folder.
        /// </summary>
        private static void LoadInputFileList()
        {
            FileInfo[] filePaths = Constants.inputDirectory.GetFiles("*.pdf");

            foreach (FileInfo path in filePaths)
                if (path.Length > 0)
                    inputFiles.Add(path.FullName);
        }

        /// <summary>
        /// Logs the error message to a file as a fallback in case the email process fails.
        /// </summary>
        /// <param name="archiveFilenameIn"></param>
        /// <param name="errorMessage"></param>
        private static void LogErrorToFile(string archiveFilenameIn, string errorMessage)
        {
            if (!Constants.errorLogDirectory.Exists)
                Constants.errorLogDirectory.Create();

            File.WriteAllText(Path.Combine(Constants.errorLogDirectory.FullName,
                Path.GetFileName(archiveFilenameIn) + " " + DateTime.Now.ToString("MMddyyyy") + ".txt"), errorMessage);
        }

        /// <summary>
        /// Uses our smtp server to generate and send an email to Laser Team
        /// that contains the corrupt zip file information.
        /// </summary>
        /// <param name="messageInsert">The error to be added to the message</param>
        private static void SendCorruptFileEmail(string messageInsert)
        {
            string subject = "ERROR: Corrupted Files Encountered";
            string message = "This is an automatically generated message, please DO NOT respond. \r\n\r\n" + messageInsert;

            var client = new SmtpClient("CERBERUS.bgms.local")
            {
                Credentials = new NetworkCredential("michael.quinton@wearebluegrass.com", "bgms16"),
                EnableSsl = true
            };
            client.Send("ihearterrors@wearebluegrass.com", "michael.quinton@wearebluegrass.com", subject, message);
        }

        /// <summary>
        /// Loads the LEADTools license to allow LEADTools functionality
        /// to be used.
        /// </summary>
        private static void SetLeadtoolsLicense()
        {
            RasterSupport.SetLicense(Constants.LEAD_LICENSE, Constants.LEAD_KEY);
        }
    } 
}

