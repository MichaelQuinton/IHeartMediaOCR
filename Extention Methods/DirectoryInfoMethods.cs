namespace Extention_Methods
{
    using System.IO;
    using System.Linq;

    /// <summary>
    /// A collection of methods to extend functionality
    /// </summary>
    public static class DirectoryInfoMethods
    {
        /// <summary>
        /// Add method to delete all files and folders from the directory.
        /// </summary>
        public static void DeleteAllContents(this DirectoryInfo directory)
        {
            foreach (var file in directory.GetFiles())
                file.Delete();

            foreach (var subDirectory in directory.GetDirectories())
                subDirectory.Delete(true);
        }

        /// <summary>
        /// Returns true if the directory contains no files.
        /// </summary>
        public static bool IsEmpty(this DirectoryInfo directory)
        {
            return !Directory.EnumerateFileSystemEntries(directory.FullName).Any();
        }
    }
}