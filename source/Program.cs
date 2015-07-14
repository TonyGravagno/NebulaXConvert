using System;
using System.Linq;

namespace NebulaRnD.Utils.NebulaXConvert
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            if (args.Count() < 2)
            {
                Console.WriteLine("not enough arguments");
            }
            else
            {
                using (ExcelConvertor c = new ExcelConvertor(args[0], args[1], (args.Length > 2 ? args[2] : "")))
                {
                }
            }
        }
    }
}