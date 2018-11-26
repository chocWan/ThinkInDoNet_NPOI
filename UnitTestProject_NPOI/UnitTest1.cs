using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace UnitTestProject_NPOI
{
    [TestClass]
    public class UnitTest1
    {

        public string excelFileTesDirPath = string.Empty;

        [TestInitialize]
        public void Init()
        {
            var dir = new DirectoryInfo(Environment.CurrentDirectory);
            excelFileTesDirPath = Path.Combine(dir.Parent.Parent.Parent.FullName, "ExcelTestFile");
        }

        [TestMethod]
        public void TestCreateExcel()
        {
            HSSFWorkbook workbook2003 = new HSSFWorkbook(); //新建xls工作簿
            workbook2003.CreateSheet("Sheet1");  //新建3个Sheet工作表
            workbook2003.CreateSheet("Sheet2");
            workbook2003.CreateSheet("Sheet3");
            FileStream file2003 = new FileStream(Path.Combine(excelFileTesDirPath, "Excel2003.xls"), FileMode.Create);
            workbook2003.Write(file2003);
            file2003.Close();  //关闭文件流
            workbook2003.Close();

            XSSFWorkbook workbook2007 = new XSSFWorkbook();  //新建xlsx工作簿
            workbook2007.CreateSheet("Sheet1");
            workbook2007.CreateSheet("Sheet2");
            workbook2007.CreateSheet("Sheet3");
            FileStream file2007 = new FileStream(Path.Combine(excelFileTesDirPath, "Excel2007.xlsx"), FileMode.Create);
            workbook2007.Write(file2007);
            file2007.Close();
            workbook2007.Close();
        }

        [TestMethod]
        public void TestSetExcelCellStyle()
        {




        }

        [TestMethod]
        public void TestCreateMergeRegion()
        {
        }

        [TestMethod]
        public void TestResolveMergeRegion()
        {
        }

        [TestMethod]
        public void TestExportExcelTemplate()
        {
        }

        [TestMethod]
        public void TestExportContentToExcel()
        {
        }

        [TestMethod]
        public void TestImportExcelToContent()
        {
        }



        [TestMethod]
        public void TestMethod1()
        {
        }


    }
}
