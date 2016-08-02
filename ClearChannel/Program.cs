namespace ClearChannel
{
    using Extention_Methods;
    using ICSharpCode.SharpZipLib.Core;
    using Leadtools;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Compression;
    using System.Linq;
    using System.Net;
    using System.Net.Mail;
    using System.Threading.Tasks;
    using SharpZip = ICSharpCode.SharpZipLib.Zip;

    internal class Program
    {
        private static readonly List<string> InputFiles = new List<string>();
        private static Dictionary<string, int> _filePageCounts = new Dictionary<string, int>();

        /// <summary>
        /// IHeart OCR and Merge Utility created by: Michael Quinton
        ///
        /// Takes the groups of input PDF files and performs OCR on the
        /// mailing address.  OCR output is overlaid with a first page marking
        /// onto the files, then all files are merged in page count order.  The
        /// resulting file is then moved to the appropriate drop folder.
        /// </summary>
        private static void Main()
        {
            Console.WriteLine("IHeart Media OCR and Merge Utility");
            if (!CheckForInputFiles())
                return;

            SetLeadtoolsLicense();

            foreach (var currentFolder in Constants.InputFolders)
            {
                Constants.InputDirectory.DeleteAllContents();
                GetInputFiles(currentFolder);

                if (Constants.InputDirectory.IsEmpty())
                    continue;

                LoadInputFileList();

                _filePageCounts = LeadToolsOcr.Process(InputFiles, currentFolder);

                PdfUtility.MergeAscendingPageCount(currentFolder, _filePageCounts);

                InputFiles.Clear();
                _filePageCounts.Clear();

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            Constants.InputDirectory.DeleteAllContents();
            ArchiveInputFiles();
            DeleteOldErrorLogs();
        }

        /// <summary>
        /// Create a compressed folder with todays date.
        /// Copy zip files used for input into compressed folder.
        /// </summary>
        private static void ArchiveInputFiles()
        {
            var fileName = "Clear Channel " + DateTime.Now.ToString("MMM_d_yyyy") + ".zip";
            var zipArchivePath = Path.Combine(Constants.ArchiveDirectory, fileName);

            using (var stream = new FileStream(zipArchivePath, FileMode.Create))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Create))
                foreach (var currentFolder in Constants.InputFolders)
                {
                    var filePath = Path.Combine(Constants.NetworkFtpDir, currentFolder.Value + "/");
                    var files = Directory.GetFiles(filePath);

                    foreach (var file in files)
                    {
                        var zipPath = currentFolder.Key + "/" + Path.GetFileName(file);
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
            return Constants.InputFolders.Select(currentFolder =>
                new DirectoryInfo(Path.Combine(Constants.NetworkFtpDir, currentFolder.Value.ToString()))).Any(directory => !directory.IsEmpty());
        }

        /// <summary>
        /// Cleans up error logs created over a month ago.
        /// </summary>
        private static void DeleteOldErrorLogs()
        {
            foreach (var file in Constants.ErrorLogDirectory.GetFiles())
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
                using (var fs = File.OpenRead(archiveFilenameIn))
                using (var zf = new SharpZip.ZipFile(fs))
                {
                    zf.IsStreamOwner = true;

                    if (!string.IsNullOrEmpty(password))
                        zf.Password = password;  //AES encrypted entries are handled automatically

                    foreach (SharpZip.ZipEntry zipEntry in zf)
                    {
                        if (!zipEntry.IsFile)
                            continue;   //Ignore Directories

                        var entryFileName = zipEntry.Name;

                        if (!entryFileName.EndsWith(".pdf", StringComparison.InvariantCultureIgnoreCase))
                            continue;

                        var buffer = new byte[4096];
                        var zipStream = zf.GetInputStream(zipEntry);

                        var fullZipToPath = Path.Combine(outFolder, entryFileName);

                        using (var streamWriter = File.Create(fullZipToPath))
                            StreamUtils.Copy(zipStream, streamWriter, buffer);
                    }
                }
            }
            catch (Exception)
            {
                var errorMessage = archiveFilenameIn + " was unable to be opened. File is likely corrupted.";
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
            var zipPath = Path.Combine(Constants.NetworkFtpDir, currentFolder.Value);
            var files = Directory.GetFiles(zipPath);

            Parallel.ForEach(files, file =>
            {
                ExtractZipFile(file, Constants.InputDirectory.FullName);
            });
        }

        /// <summary>
        /// Populate the inputFiles List with the file paths of files in the input folder.
        /// </summary>
        private static void LoadInputFileList()
        {
            var filePaths = Constants.InputDirectory.GetFiles("*.pdf");

            foreach (var path in filePaths)
                if (path.Length > 0)
                    InputFiles.Add(path.FullName);
        }

        /// <summary>
        /// Logs the error message to a file as a fallback in case the email process fails.
        /// </summary>
        /// <param name="archiveFilenameIn"></param>
        /// <param name="errorMessage"></param>
        private static void LogErrorToFile(string archiveFilenameIn, string errorMessage)
        {
            if (!Constants.ErrorLogDirectory.Exists)
                Constants.ErrorLogDirectory.Create();

            File.WriteAllText(Path.Combine(Constants.ErrorLogDirectory.FullName,
                Path.GetFileName(archiveFilenameIn) + " " + DateTime.Now.ToString("MMddyyyy") + ".txt"), errorMessage);
        }

        /// <summary>
        /// Uses our smtp server to generate and send an email to Laser Team
        /// that contains the corrupt zip file information.
        /// </summary>
        /// <param name="messageInsert">The error to be added to the message</param>
        private static void SendCorruptFileEmail(string messageInsert)
        {
            const string subject = "ERROR: Corrupted Files Encountered";
            var message = "This is an automatically generated message, please DO NOT respond. \r\n\r\n" + messageInsert;
            const string user = @"iheartreport@wearebluegrass.com";
            const string password = @"Uer?5E.k";

            var client = new SmtpClient("CERBERUS.bgms.local")
            {
                Credentials = new NetworkCredential(user, password),
                EnableSsl = true
            };
            client.Send("ihearterrors@wearebluegrass.com", "LaserTeam@wearebluegrass.com", subject, message);
        }

        /// <summary>
        /// Loads the LEADTools license to allow LEADTools functionality
        /// to be used.
        /// </summary>
        private static void SetLeadtoolsLicense()
        {
            RasterSupport.SetLicense(Constants.LeadLicense, Constants.LeadKey);
        }
    }
}