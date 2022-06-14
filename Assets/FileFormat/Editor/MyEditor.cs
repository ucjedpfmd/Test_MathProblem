using System.IO;
using UnityEngine;
using UnityEditor;

public class Excel4Unity : Editor {
    [MenuItem("阿尔法/自动出题", false, 11)]
    private static void SetCalcExcel() {
        string filePath = @$"{Application.dataPath}\abc.xlsx";
        if (File.Exists(filePath)) {
            Excel xls = ExcelHelper.LoadExcel(filePath);
            for (int j = 0; j < xls.Tables.Count; j++) {
                var table = xls.Tables[j];
                if (table.TableName.Contains("加减法")) {
                    var starRow = 3; //从第三行开始
                    var maxNum = 20;
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
                }
            }

            Debug.Log("自动出题完成");
            ExcelHelper.SaveExcel(xls, filePath);
        } else {
            Debug.Log("SetCalcExcel 不存在" + filePath);
        }
    }

    public int[] rangeList = new[] { 0, 1, 2 };
    private static string[] GetCaleArr() {
        var emptyNum = 0;
        var r = Random.Range(0, 10000);
        if (r < 500) {
            emptyNum = 0;
        }else if (r < 1000) {
            emptyNum = 1;
        } else {
            emptyNum = 2;
        }
        
       
        var symbolVal = Random.Range(0, 2) == 0 ? "+" : "-";
        var valArr = new[] { "", "", "", symbolVal };
        valArr[emptyNum] = "(  )";
        if (emptyNum == 0) {
            valArr[1] = Random.Range(0, 10).ToString();
            if (symbolVal == "+") {
                valArr[2] = Random.Range(int.Parse(valArr[1]), 21).ToString();
            } else {
                valArr[2] = Random.Range(0, 20 - int.Parse(valArr[1])).ToString();
            }
        } else if (emptyNum == 1) {
            if (symbolVal == "+") {
                valArr[0] = Random.Range(10, 20).ToString();
                valArr[2] = (Random.Range(int.Parse(valArr[0]), 21)).ToString();
            } else {
                valArr[0] = Random.Range(10, 20).ToString();
                valArr[2] = Random.Range(0, int.Parse(valArr[0]) + 1).ToString();
            }
        } else if (emptyNum == 2) {
            if (symbolVal == "+") {
                valArr[0] = Random.Range(5, 10).ToString();
                valArr[1] = Random.Range(5, 10).ToString();
            } else {
                valArr[0] = Random.Range(10, 20).ToString();
                valArr[1] = (Random.Range(0, 10)).ToString();
            }
        }

        return valArr;
    }
}
