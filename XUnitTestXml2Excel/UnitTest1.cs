using System.IO;
using Xml2Excel;
using Xunit;

namespace XUnitTestXml2Excel
{
    public class UnitTest1
    {
        [Fact]
        public void GenerateExcel()
        {
            Xml2ExcelCore xml2ExcelCore = new Xml2ExcelCore();
            string xml = File.ReadAllText("test-excel.xml");
            bool result=xml2ExcelCore.Generate(xml, "test1.xlsx");
            Assert.True(result, "Se genero correctamente el excel");
        }
    }
}
