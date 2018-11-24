using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.IO;

public class ExeclHelper
{
	public DataTable ImportExcelData(Stream s, bool msExcel = false)
	{
		return msExcel ? ImportOldExcelData(s) : ImportNewExcelData(s);
	}

	public DataTable ImportExcelData(string html5Uploadbase64String)
	{
		string[] array = html5Uploadbase64String.Split(new string[1]
		{
			"base64,"
		}, StringSplitOptions.RemoveEmptyEntries);
		bool msExcel = array[0].Contains("application/vnd.ms-excel");
		byte[] array2 = Convert.FromBase64String(array[1]);
		MemoryStream memoryStream = null;
		DataTable result = null;
		try
		{
			memoryStream = new MemoryStream();
			memoryStream.Write(array2, 0, array2.Length);
			memoryStream.Position = 0L;
			result = ImportExcelData(memoryStream, msExcel);
		}
		catch (Exception)
		{
		}
		finally
		{
			memoryStream.Close();
		}
		return result;
	}

	private DataTable ImportOldExcelData(Stream s)
	{
		s.Position = 0L;
		HSSFWorkbook hSSFWorkbook = new HSSFWorkbook(s);
		ISheet sheetAt = hSSFWorkbook.GetSheetAt(0);
		IRow columns = GetColumns(sheetAt);
		DataTable dataTable = CreateTable(columns);
		FillDataTable(sheetAt, dataTable, columns);
		return dataTable;
	}

	private IRow GetColumns(ISheet sheet)
	{
		for (int i = 0; i < 50; i++)
		{
			IRow row = sheet.GetRow(i);
			string text = string.Empty;
			int num = 1;
			if (row != null)
			{
				for (int j = 0; j < row.Cells.Count; j++)
				{
					if (string.IsNullOrEmpty(row.Cells[j].ToString()) || row.Cells[j].ToString() == text)
					{
						row.Cells[j].SetCellValue(text + num);
						num++;
					}
					else
					{
						text = row.Cells[j].ToString();
						num = 1;
					}
				}
				return row;
			}
		}
		return null;
	}

	private DataTable CreateTable(IRow columns)
	{
		DataTable dataTable = new DataTable();
		if (columns != null)
		{
			string text = string.Empty;
			int num = 1;
			foreach (ICell cell in columns.Cells)
			{
				DataColumn dataColumn = new DataColumn();
				if (string.IsNullOrEmpty(cell.ToString()) || cell.ToString() == text)
				{
					dataColumn.ColumnName = text + num;
					num++;
				}
				else
				{
					dataColumn.ColumnName = cell.ToString();
					text = dataColumn.ColumnName;
					num = 1;
				}
				dataColumn.DataType = typeof(string);
				dataTable.Columns.Add(dataColumn);
			}
			return dataTable;
		}
		return dataTable;
	}

	private void FillDataTable(ISheet sheet, DataTable dt, IRow columns)
	{
		for (int i = columns.RowNum + 1; sheet.GetRow(i) != null || sheet.GetRow(i + 1) != null; i++)
		{
			IRow row = sheet.GetRow(i);
			if (row != null)
			{
				DataRow dataRow = dt.NewRow();
				foreach (ICell cell2 in columns.Cells)
				{
					ICell cell = FindColumnValue(row, cell2.ColumnIndex);
					dataRow[cell2.ToString()] = ((cell == null) ? string.Empty : cell.ToString());
				}
				dt.Rows.Add(dataRow);
			}
		}
	}

	private ICell FindColumnValue(IRow columns, int columnsIndex)
	{
		ICell result = null;
		foreach (ICell cell in columns.Cells)
		{
			if (cell.ColumnIndex == columnsIndex)
			{
				result = cell;
				break;
			}
		}
		return result;
	}

	private DataTable ImportNewExcelData(Stream s)
	{
		s.Position = 0L;
		IWorkbook workbook = new XSSFWorkbook(s);
		ISheet sheetAt = workbook.GetSheetAt(0);
		IRow columns = GetColumns(sheetAt);
		DataTable dataTable = CreateTable(columns);
		FillDataTable(sheetAt, dataTable, columns);
		return dataTable;
	}
}
