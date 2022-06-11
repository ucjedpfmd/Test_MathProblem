using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Style;
using UnityEngine;

public class ExcelHelper
{

    public static Excel LoadExcel(string path)
    {
        Debug.Log("LoadExcel" + path);
        FileInfo file = new FileInfo(path);
        ExcelPackage ep = new ExcelPackage(file);
        Excel xls = new Excel(ep.Workbook);
        return xls;
    }

	public static Excel CreateExcel(string path) {
		ExcelPackage ep = new ExcelPackage ();
		ep.Workbook.Worksheets.Add ("sheet");
		Excel xls = new Excel(ep.Workbook);
		SaveExcel (xls, path);
		return xls;
	}

    public static void SaveExcel(Excel xls, string path)
    {
        FileInfo output = new FileInfo(path);
        ExcelPackage ep = new ExcelPackage();
        for (int i = 0; i < xls.Tables.Count; i++)
        {
            ExcelTable table = xls.Tables[i];
            ExcelWorksheet sheet = ep.Workbook.Worksheets.Add(table.TableName);
            sheet.PrinterSettings.RightMargin = 1M / 2.54M;
            sheet.PrinterSettings.LeftMargin = 1M / 2.54M;
            var widthArr = new[]
                { 5, 3, 5, 3, 5, 3, 5, 3, 5, 3, 5, 3, 5, 3, 5, 3, 5, 3, 5, 3, 5, 3, 5, 3, };
            var heightArr = new[]
                { 25, 25, 25, 25, 25, 25, 25, 25, 25, 25, 25, 25,25, 25, 25, 25, 25, 25,25, 25, 25, 25, 25, 25,25,25, 25,25};
           
            sheet.InsertColumn(1, widthArr.Length);
            sheet.InsertRow(1, heightArr.Length);
            for (int j = 1; j <= widthArr.Length; j++) {
                sheet.Column(j).Width = widthArr[j - 1];
            }
            
            for (int j = 1; j <= heightArr.Length; j++) {
                sheet.Row(j + 2).Height = heightArr[j - 1];
            }
            sheet.Cells[1, 1, 1, 24].Merge = true;
            sheet.Cells[2, 1, 2, 24].Merge = true;
            for (int row = 1; row <= table.NumberOfRows; row++) {
                for (int column = 1; column <= table.NumberOfColumns; column++) {
                    var cell = sheet.Cells[row, column];
                    cell.Style.Font.Name = "宋体";
                    cell.Style.Font.Size = 14;
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[row, column].Value = table.GetValue(row, column);
                }
            }
        }
        ep.SaveAs(output);
    }
}
