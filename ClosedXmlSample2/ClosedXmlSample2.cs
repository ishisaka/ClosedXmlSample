using System;
using System.Linq;
using ClosedXML.Excel;


namespace ClosedXmlSample
{
    /// <summary>
    /// ClosedXMLを使用し、Excelからセルの値を取得する
    /// </summary>
    class ClosedXmlSample2
    {
        static void Main(string[] args) {
            string filePath = args[0];      //ファイルパス
            string sheetName = args[1];     //データを読み込むセルがあるワークシートの名前
            string addressName = args[2];   //ワークシートのセル位置

            using (var wb = new XLWorkbook(filePath)) {
                var ws = wb.Worksheets.First(s => s.Name == sheetName);
                string format;
                //セルの型がわかるので処理がしやすい
                switch (ws.Cell(addressName).DataType) {
                    case XLCellValues.Text:
                        Console.WriteLine("Excel Data Type is Text.");
                        Console.WriteLine(ws.Cell(addressName).Value);
                        break;
                    case XLCellValues.Number:
                        Console.WriteLine("Excel Data Type is Number.");
                        format = ws.Cell(addressName).Style.NumberFormat.Format;
                        Console.WriteLine("Format: " + format);
                        Console.WriteLine(ws.Cell(addressName).Value);
                        break;
                    case XLCellValues.Boolean:
                        Console.WriteLine("Excel Data Type is Boolean.");
                        Console.WriteLine(ws.Cell(addressName).Value);
                        break;
                    case XLCellValues.DateTime:
                        Console.WriteLine("Excel Data Type is DateTime.");
                        format = ws.Cell(addressName).Style.NumberFormat.Format;
                        Console.WriteLine("Format: " + format);
                        //セルのバリューはObject型なので、必要に応じてキャストしてやる
                        Console.WriteLine(((DateTime)ws.Cell(addressName).Value).ToString());
                        break;
                    case XLCellValues.TimeSpan:
                        Console.WriteLine("Excel Data Type is TimeSpan");
                        format = ws.Cell(addressName).Style.NumberFormat.Format;
                        Console.WriteLine("Format: " + format);
                        Console.WriteLine(((TimeSpan)ws.Cell(addressName).Value).ToString());
                        break;
                    default:
                        break;
                }
            }
            Console.Read();
        }
    }
}
