using System;
using System.IO;
using System.Security.Permissions;
using Xml2Excel;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            MemoryStream memoryStream = null;
            try
            {
                Xml2ExcelCore xml2ExcelCore = new Xml2ExcelCore();
                string xml = File.ReadAllText("test-excel.xml");
                using (memoryStream = xml2ExcelCore.Generate(xml))
                {
                    File.WriteAllBytes("test1.xlsx", memoryStream.ToArray());
                    Console.WriteLine("Success");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error:");
            }
            finally
            {
                memoryStream.Dispose();
            }
            
        }

    }
}
