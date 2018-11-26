using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
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

            //两个版本DLL一起使用会有问题
            //XSSFWorkbook workbook2007 = new XSSFWorkbook();  //新建xlsx工作簿
            //workbook2007.CreateSheet("Sheet1");
            //workbook2007.CreateSheet("Sheet2");
            //workbook2007.CreateSheet("Sheet3");
            //FileStream file2007 = new FileStream(Path.Combine(excelFileTesDirPath, "Excel2007.xlsx"), FileMode.Create);
            //workbook2007.Write(file2007);
            //file2007.Close();
            //workbook2007.Close();
        }

        [TestMethod]
        public void TestSetExcelCellStyle()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();  
            ISheet sheet = workbook.CreateSheet("Sheet1");

            //背景颜色
            HSSFPalette palette = workbook.GetCustomPalette();
            palette.SetColorAtIndex((short)9, (byte)255, (byte)121, (byte)121);
            HSSFColor hssFColor = palette.FindColor((byte)255, (byte)121, (byte)121);
            ICellStyle bgColorCellStyle = workbook.CreateCellStyle();
            bgColorCellStyle.FillPattern = FillPattern.SolidForeground;
            bgColorCellStyle.FillForegroundColor = hssFColor.Indexed;
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.SetCellValue("背景颜色");
            cell.CellStyle = bgColorCellStyle;

            //字体
            IFont font = workbook.CreateFont();
            font.Boldweight = short.MinValue;//粗体     
            font.FontName = "宋体";
            font.FontHeightInPoints = 20;
            font.Color = HSSFColor.DarkRed.Index;
            font.Underline = FontUnderlineType.Double;
            ICellStyle fontCellStyle = workbook.CreateCellStyle();
            fontCellStyle.SetFont(font);
            row = sheet.CreateRow(1);
            cell = row.CreateCell(0);
            cell.SetCellValue("字体");
            cell.CellStyle = fontCellStyle;

            //保留2位小数
            ICellStyle decimal2CellStyle = workbook.CreateCellStyle();
            decimal2CellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");//保留两位小数
            row = sheet.CreateRow(2);
            cell = row.CreateCell(0);
            cell.SetCellValue(1.222222);
            cell.CellStyle = decimal2CellStyle;

            //日期格式
            IDataFormat datetimeFormat = workbook.CreateDataFormat();
            ICellStyle datetimeCellStyle = workbook.CreateCellStyle();
            datetimeCellStyle.DataFormat = datetimeFormat.GetFormat("yyyy年m月d日");
            row = sheet.CreateRow(3);
            cell = row.CreateCell(0);
            cell.SetCellValue(new DateTime(2018,11,26));
            cell.CellStyle = datetimeCellStyle;

            //货币格式
            IDataFormat currencyFormat = workbook.CreateDataFormat();
            ICellStyle currencyCellStyle = workbook.CreateCellStyle();
            currencyCellStyle.DataFormat = currencyFormat.GetFormat("¥#,##0");
            row = sheet.CreateRow(4);
            cell = row.CreateCell(0);
            cell.SetCellValue(888);
            cell.CellStyle = currencyCellStyle;

            //百分比
            ICellStyle percentCellStyle = workbook.CreateCellStyle();
            percentCellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00%");//百分比
            row = sheet.CreateRow(5);
            cell = row.CreateCell(0);
            cell.SetCellValue(0.88);
            cell.CellStyle = percentCellStyle;

            //中文大写
            IDataFormat chsFormat = workbook.CreateDataFormat();
            ICellStyle chsCellStyle = workbook.CreateCellStyle();
            chsCellStyle.DataFormat = chsFormat.GetFormat("[DbNum2][$-804]0");
            row = sheet.CreateRow(6);
            cell = row.CreateCell(0);
            cell.SetCellValue("万事如意");
            cell.CellStyle = chsCellStyle;

            //边框
            ICellStyle borderCellStyle = workbook.CreateCellStyle();
            borderCellStyle.BorderBottom = BorderStyle.Thin;
            borderCellStyle.BottomBorderColor = HSSFColor.Red.Index;
            borderCellStyle.BorderTop = BorderStyle.Thin;
            borderCellStyle.BorderLeft = BorderStyle.Thin;
            borderCellStyle.BorderRight = BorderStyle.Thin;
            row = sheet.CreateRow(7);
            cell = row.CreateCell(0);
            cell.SetCellValue("边框");
            cell.CellStyle = borderCellStyle;

            //自动换行
            ICellStyle autoGrowCellStyle = workbook.CreateCellStyle();
            autoGrowCellStyle.WrapText = true;
            row = sheet.CreateRow(8);
            cell = row.CreateCell(0);
            cell.SetCellValue("yiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyiyi");
            cell.CellStyle = autoGrowCellStyle;

            //高度宽度
            ICellStyle whCellStyle = workbook.CreateCellStyle();
            whCellStyle.WrapText = true;
            row = sheet.CreateRow(9);
            cell = row.CreateCell(0);
            //设置第1列的宽度，第10行的高度
            sheet.SetColumnWidth(0,10*256);
            row.Height = 10*256;
            cell.SetCellValue("yiyi");
            cell.CellStyle = whCellStyle;

            //对齐
            ICellStyle dqCellStyle = workbook.CreateCellStyle();
            dqCellStyle.Alignment = HorizontalAlignment.Center;
            dqCellStyle.VerticalAlignment = VerticalAlignment.Center;
            row = sheet.CreateRow(10);
            cell = row.CreateCell(0);
            sheet.SetColumnWidth(0, 20 * 256);
            row.Height = 10 * 256;
            cell.SetCellValue("yiyi");
            cell.CellStyle = whCellStyle;

            //公式
            row = sheet.CreateRow(11);
            row.CreateCell(0).SetCellValue(1);
            row.CreateCell(1).SetCellValue(2);
            cell = row.CreateCell(2);
            cell.SetCellFormula("SUM(A12,B12)");
            cell = row.CreateCell(3);
            cell.SetCellFormula("SUM(A12:C12)");

            //下拉列表 第二行 第二列
            CellRangeAddressList regions = new CellRangeAddressList(1, 1, 1, 1);
            DVConstraint constraint = DVConstraint.CreateExplicitListConstraint(new string[] { "itemA", "itemB", "itemC" });
            HSSFDataValidation dataValidate = new HSSFDataValidation(regions, constraint);
            sheet.AddValidationData(dataValidate);

            //冻结行列 前2行，前1列
            sheet.CreateFreezePane(1, 2, 1, 2);

            FileStream fileStream = new FileStream(Path.Combine(excelFileTesDirPath, "TestSetExcelCellStyle.xls"), FileMode.Create);
            workbook.Write(fileStream);
            fileStream.Close();
            workbook.Close();
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
