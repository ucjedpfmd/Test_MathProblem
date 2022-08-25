using UnityEngine;
using UnityEditor;
using System.IO;
using System.Text;

public class Excel4Unity : Editor {
    [MenuItem("阿尔法/自动出题", false, 11)]
    private static void SetCalcExcel() {
        string filePath = $"{Application.dataPath}/abc.xlsx";
        if (File.Exists(filePath)) {
            Excel xls = ExcelHelper.LoadExcel(filePath);
            var len = 100;
            var firstTable = xls.Tables[0];
            xls.Tables.Clear();
            for (int j = 0; j < len; j++) {
                var table = new ExcelTable();
                table.TableName = "加减法" + j;
                var starRow = 3; //从第三行开始
                var maxNum = 26;
                
                for (int i = 0; i < starRow; i++) {
                    for (int k = 0; k < maxNum; k++) {
                        table.SetValue(i, k, firstTable.GetValue(i,k).ToString());
                    }
                }
                
                for (int i = starRow; i < starRow + 25; i++) {
                    var calcArr = GetCaleArr();
                    table.SetValue(i, 1, calcArr[0]);
                    table.SetValue(i, 2, calcArr[3]);
                    table.SetValue(i, 3, calcArr[1]);
                    table.SetValue(i, 4, "=");
                    table.SetValue(i, 5, calcArr[2]);

                    calcArr = GetCaleArr();
                    table.SetValue(i, 7, calcArr[0]);
                    table.SetValue(i, 8, calcArr[3]);
                    table.SetValue(i, 9, calcArr[1]);
                    table.SetValue(i, 10, "=");
                    table.SetValue(i, 11, calcArr[2]);

                    calcArr = GetCaleArr();
                    table.SetValue(i, 13, calcArr[0]);
                    table.SetValue(i, 14, calcArr[3]);
                    table.SetValue(i, 15, calcArr[1]);
                    table.SetValue(i, 16, "=");
                    table.SetValue(i, 17, calcArr[2]);

                    calcArr = GetCaleArr();
                    table.SetValue(i, 19, calcArr[0]);
                    table.SetValue(i, 20, calcArr[3]);
                    table.SetValue(i, 21, calcArr[1]);
                    table.SetValue(i, 22, "=");
                    table.SetValue(i, 23, calcArr[2]);
                }
                xls.Tables.Add(table);
            }
            

            Debug.Log("自动出题完成");
            ExcelHelper.SaveExcel(xls, filePath);
        } else {
            Debug.Log("SetCalcExcel 不存在" + filePath);
        }
    }

    private static string[] GetCaleArr() {
        var emptyNum = Random.Range(0, 3);
        var symbolVal = Random.Range(0, 2) == 0 ? "+" : "-";
        var valArr = new[] { "", "", "", symbolVal };
        valArr[emptyNum] = "(  )";
        var maxValue = 10;
        var maxEmpty2Value = 1;
        if (emptyNum == 0) {
            if (symbolVal == "+") {
                valArr[1] = Random.Range(1, maxValue).ToString();
                valArr[2] = Random.Range(int.Parse(valArr[1]), maxValue).ToString();
            } else {
                valArr[1] = Random.Range(1, 10).ToString();
                valArr[2] = Random.Range(0, maxValue - int.Parse(valArr[1])).ToString();
            }
        } else if (emptyNum == 1) {
            if (symbolVal == "+") {
                valArr[0] = Random.Range(1, maxValue).ToString();
                valArr[2] = (Random.Range(int.Parse(valArr[0]), maxValue)).ToString();
            } else {
                valArr[0] = Random.Range(maxEmpty2Value, maxValue).ToString();
                valArr[2] = Random.Range(maxEmpty2Value, int.Parse(valArr[0]) + 1).ToString();
            }
        } else if (emptyNum == 2) {
            if (symbolVal == "+") {
                valArr[0] = Random.Range(5, maxValue).ToString();
                valArr[1] = Random.Range(1, maxValue - int.Parse(valArr[0])).ToString();
            } else {
                valArr[0] = Random.Range(maxEmpty2Value, maxValue).ToString();
                valArr[1] = (Random.Range(0, 10)).ToString();
            }
        }

        return valArr;
    }
}
