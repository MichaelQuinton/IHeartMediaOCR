namespace ClearChannel
{
    using System.Collections.Generic;
    using System.IO;
    using System.Reflection;

    internal static class Constants
    {
        internal const string ArchiveDirectory = @"\\bgtera01\TempJobs\Laser";
        internal const string NetworkFtpDir = @"\\kratos\clearchannel\";
        internal const string OutputDirectory = @"\\poseidon\submit\FTP\Clear Channel\";

        private static string AssemblyDir => Path.GetDirectoryName(Assembly.GetAssembly(typeof(Program)).FullName);
        internal static DirectoryInfo ErrorLogDirectory => new DirectoryInfo(Path.Combine(AssemblyDir, "ErrorLogs"));
        internal static DirectoryInfo InputDirectory => new DirectoryInfo(Path.Combine(AssemblyDir, "Input"));
        internal static string LeadKey => File.ReadAllText(@"C:\LEADTOOLS 19\Leadtools V19 License\Leadtools v19.key");
        internal static string LeadLicense => @"C:\LEADTOOLS 19\Leadtools V19 License\Leadtools v19.lic";

        internal static readonly Dictionary<string, string> InputFolders = new Dictionary<string, string>()
            {
                //{"Aloha","Aloha Trust"},
                //{"Premier", "CCSAPB"},
                {"PremierLandscape", "CCSAPBL"}
                //{"SpecialBilling", "CCSASB"},
                //{"TotalTraffic", "CCSATT"},
                //{"Radio", "Clear Channel"},
                //{"LockboxInsert", "LockboxInsert"}
            };

        internal static readonly Dictionary<string, string> OutputFolders = new Dictionary<string, string>()
            {
                {"Aloha", "aloha"},
                {"Premier", "CCSAPB"},
                {"PremierLandscape", "CCSAPBL"},
                {"SpecialBilling", "CCSASB"},
                {"TotalTraffic", "CCSATT"},
                {"Radio", "iheart"},
                {"LockboxInsert", "iheart_insert"}
            };
    }
}