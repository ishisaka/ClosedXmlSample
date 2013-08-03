using System;
using ClosedXML.Excel;

namespace ClosedXMLSample
{
    /// <summary>
    /// ClosedXMLを使用してExcelのファイルを作成する。
    /// </summary>
    class ClosedXmlSample1
    {
        static void Main(string[] args) {
            //WorkBookオブジェクトを追加
            using (var wb = new XLWorkbook()) {
                //WorkSheetをWorkBookオブジェクトに追加
                var ws = wb.Worksheets.Add("Sample");
                //セルの値を変更
                ws.Cell("A1").Value = "Hello World";
                //作業内容を保存
                wb.SaveAs("HelloWorld.xlsx");
            }

        }
    }
}
