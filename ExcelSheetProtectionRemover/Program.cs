using ProtectionRemoverLib;
using System;

namespace ExcelSheetProtectionRemover
{
    internal class Program
    {
        static void Main(string[] args)
        {
            foreach (string arg in args)
            {
                ProtectionRemover.RemoveProtection(arg);
            }
            Console.WriteLine("OK");
        }
    }
}
