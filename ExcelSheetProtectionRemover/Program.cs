using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;


namespace ExcelSheetProtectionRemover
{
    internal class Program
    {
        static void Main(string[] args)
        {
            foreach (string arg in args)
            {
                var file = new FileInfo(arg);
                var zipFile = ZipFile.Open(file.FullName, mode: ZipArchiveMode.Update);

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
                    var remove = RemoveProtection(context);
                    File.WriteAllText(xml, remove);
                }

                //Create Zip
                ZipFile.CreateFromDirectory(tempPath, file.FullName + "(No Protect)" + file.Extension);
                Directory.Delete(tempPath, true);
            }

        }

        public static string RemoveProtection(string str)
        {
            var pattern = "(<sheetProtection[^>]*\\s*/>)";
            var replacePattern = "";
            return Regex.Replace(str, pattern, replacePattern, RegexOptions.IgnoreCase);
        }


    }
}
