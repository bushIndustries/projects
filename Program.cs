using System;
using System.IO;

namespace InstructionSheetProcessor
{
    class Program
    {
        static void Main()
        {
            FileInfo errorText = null;
            try
            {
                // Delete if exists
                errorText = new FileInfo(@"Error.txt");
                if (errorText.Exists)
                    errorText.Delete();

                // new error file to record what happens here
                using (StreamWriter sw = errorText.CreateText())
                    sw.WriteLine("Started Program");

                // do it
                ProcessFiles(errorText);

                // all done
                using (StreamWriter sw1 = errorText.AppendText())
                    sw1.WriteLine("All files processed.");

            }
            catch (Exception ex)
            {
                try
                {
                    using (StreamWriter sw2 = errorText.AppendText())
                        sw2.WriteLine("Error: " + ex.Message);
                }
                finally{ }
            }
        }

        static public void ProcessFiles(FileInfo errorText)
        {
            string archiveDirectory = @"\\netapp-01\instruction\AI Processed\";
            string archiveFilePath = null;
            string sourceDirectory = @"\\netapp-01\instruction\AI Released\";
            string sourceFileName = null;
            string sourceFilePath = null;
            string destinationDirectory = @"\\S060BD3R\WWW\bushps\htdocs\instructionsheets\submit\";
            string destinationFilePath = null;

            using (StreamWriter sw = errorText.AppendText())
            {
                sw.WriteLine("Gonna read files ...");

                try
                {
                    var sourceFiles = Directory.EnumerateFiles(sourceDirectory);
                    sw.WriteLine("Read them. gonna walk the list ...");
                    foreach (string currentFile in sourceFiles)
                    {
                        sourceFileName = currentFile.Substring(sourceDirectory.Length);
                        sourceFilePath = Path.Combine(sourceDirectory, sourceFileName);
                        destinationFilePath = Path.Combine(destinationDirectory, sourceFileName);
                        archiveFilePath = Path.Combine(archiveDirectory, sourceFileName);

                        // remove any weird perms and push file to iSeries
                        sw.WriteLine("  - Going to copy file: " + sourceFileName);
                        File.SetAttributes(sourceFilePath, FileAttributes.Normal);
                        File.Copy(sourceFilePath, destinationFilePath, true);

                        // move file to the arhive folder
                        sw.WriteLine("  - Going to archive file: " + sourceFileName);
                        File.Move(sourceFilePath, archiveFilePath, true);
                    }
                }
                catch (Exception ex)
                {
                    sw.WriteLine("Error: " + ex.Message);
                }
            }

        }
    }
}
