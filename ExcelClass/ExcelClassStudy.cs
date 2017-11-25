using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelClass
{
    public static class ExcelClassStudy
    {
        public static void CleanPassword(this Excel.Worksheet sheet)
        {
            sheet.Protect
            (
                DrawingObjects: Office.MsoTriState.msoTrue,
                Contents: Office.MsoTriState.msoTrue,
                Scenarios: Office.MsoTriState.msoTrue,
                AllowFiltering: Office.MsoTriState.msoTrue,
                AllowUsingPivotTables: Office.MsoTriState.msoTrue
            );

            sheet.Protect
            (
                DrawingObjects: Office.MsoTriState.msoTrue,
                Contents: Office.MsoTriState.msoTrue,
                Scenarios: Office.MsoTriState.msoTrue,
                AllowFiltering: Office.MsoTriState.msoTrue,
                AllowUsingPivotTables: Office.MsoTriState.msoTrue
            );

            sheet.Protect
            (
                DrawingObjects: Office.MsoTriState.msoTrue,
                Contents: Office.MsoTriState.msoTrue,
                Scenarios: Office.MsoTriState.msoTrue,
                AllowFiltering: Office.MsoTriState.msoTrue,
                AllowUsingPivotTables: Office.MsoTriState.msoTrue
            );

            sheet.Protect
            (
                DrawingObjects: Office.MsoTriState.msoTrue,
                Contents: Office.MsoTriState.msoTrue,
                Scenarios: Office.MsoTriState.msoTrue,
                AllowFiltering: Office.MsoTriState.msoTrue,
                AllowUsingPivotTables: Office.MsoTriState.msoTrue
            );

            sheet.Unprotect();
        }

        public static string GetContent1(string url)
        {
            //第一句，抓
            XElement xml = XElement.Load(url);

            //第二句，取
            string txt = "------------ 数学大冒险 -------------" + "\r\n";
            var list = xml.Element("channel").Elements("item")
                .Select((m, index1) => txt += index1.ToString() + ":" + m.Element("title").Value + "\r\n")
                .Where((n, index2) => index2 < 5)
                .ToList();

            //
            return txt;
        }

        public static string GetContent(string url)
        {
            //第一句，抓
            XElement xml = XElement.Load(url);

            //第二句，取
            string txt = "------------ 数学大冒险 -------------" + "\r\n";
            var list = xml.Element("channel").Elements("item")
                .Select((m, index) => index + 1 + ":" + m.Element("title").Value + "\r\n")
                //.Where((n, index) => index < 5 && n.StartsWith("1"))
                .ToList();

            foreach (var item in list)
            {
                txt += item;
            }
            return txt;
        }

        /// <summary>
        /// 插入一列数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowID">行号</param>
        /// <param name="columnID">列号</param>
        /// <param name="txt">内容以\r\n分段</param>
        public static void InsertColumnText(this Excel.Worksheet sheet, int rowID, int columnID, string txt)
        {

            string[] txtArray = txt.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < txtArray.Length; i++)
            {
                Excel.Range rng = (Excel.Range)sheet.Cells[rowID + i, columnID];
                rng.Value2 = txtArray[i];
            }

        }
    }
}
