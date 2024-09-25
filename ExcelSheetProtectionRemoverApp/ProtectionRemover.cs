using System;
using System.IO.Compression;
using System.IO;
using System.Text.RegularExpressions;

namespace ProtectionRemoverLib
{

    public class ProtectionRemover
    {
        public static void RemoveProtection(string filePath)
        {
            var file = new FileInfo(filePath);
            var zipFile = ZipFile.Open(file.FullName, mode: ZipArchiveMode.Read);

            //UnZip
            var tempPath = Path.Combine(Path.GetTempPath(), file.Name.Replace(file.Extension, ""));
            if (Directory.Exists(tempPath))
            {
                Directory.Delete(tempPath, true);
            }
            Directory.CreateDirectory(tempPath);
            zipFile.ExtractToDirectory(tempPath);

            //Get all xl\worksheets xml file
            var xmls = Directory.GetFiles(Path.Combine(tempPath, @"xl\worksheets"));
            foreach (var xml in xmls)
            {
                var context = File.ReadAllText(xml);
                var remove = RemoveXmlProtection(context);
                File.WriteAllText(xml, remove);
            }

            //Create Zip
            ZipFile.CreateFromDirectory(tempPath, file.FullName + "(No Protect)" + file.Extension);
            Directory.Delete(tempPath, true);
        }
        protected static string RemoveXmlProtection(string str)
        {
            var pattern = "(<sheetProtection[^>]*\\s*/>)";
            var replacePattern = "";
            return Regex.Replace(str, pattern, replacePattern, RegexOptions.IgnoreCase);
        }
    }
}
