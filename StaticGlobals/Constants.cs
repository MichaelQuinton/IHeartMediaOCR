using System.Collections.Generic;
using System.IO;
using System.Reflection;

class Constants
{
    internal const string archiveDirectory = @"\\bgtera01\TempJobs\Laser";
    internal const string networkFTPDir = @"\\kratos\clearchannel\";
    internal const string outputDirectory = @"\\poseidon\submit\FTP\Clear Channel\";

    internal static string assemblyDir { get { return Path.GetDirectoryName(Assembly.GetAssembly(typeof(Constants)).FullName); } }
    internal static DirectoryInfo inputDirectory { get { return new DirectoryInfo(Path.Combine(assemblyDir, "Input")); } }
    internal static string LEAD_KEY { get { return File.ReadAllText(@"C:\LEADTOOLS 19\Leadtools V19 License\Leadtools v19.key"); } }
    internal static string LEAD_LICENSE { get { return @"C:\LEADTOOLS 19\Leadtools V19 License\Leadtools v19.lic"; } }

    internal static Dictionary<string, string> inputFolders = new Dictionary<string, string>()
            {
                {"Aloha","Aloha Trust"},
                {"Premier", "CCSAPB"},
                {"PremierLandscape", "CCSAPBL"},
                {"SpecialBilling", "CCSASB"},
                {"TotalTraffic", "CCSATT"},
                {"Radio", "Clear Channel"}

            };
    internal static Dictionary<string, string> outputFolders = new Dictionary<string, string>()
            {
                {"Aloha", "aloha"},
                {"Premier", "CCSAPB"},
                {"PremierLandscape", "CCSAPBL"},
                {"SpecialBilling", "CCSASB"},
                {"TotalTraffic", "CCSATT"},
                {"Radio", "iheart"},
            };
}

