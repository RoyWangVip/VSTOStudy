using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using ExcelClass;
using System.Text.RegularExpressions;

namespace ExcelStudyHome
{
    public partial class ExcelStudyHomeRibbon
    {
        private void ExcelStudyHomeRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;

            sheet.CleanPassword();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            string url = @"http://www.matrix67.com/blog/feed";
            string txt = ExcelClassStudy.GetContent(url);
            System.Windows.Forms.MessageBox.Show(txt);
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
            string url = @"http://www.matrix67.com/blog/feed";
            string txt = ExcelClassStudy.GetContent(url);

            sheet.InsertColumnText(1, 1, txt);
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            //测试
            Excel.Application app = Globals.ThisAddIn.Application;

            string str = System.Windows.Forms.Clipboard.GetText();

            //[王二] 年龄 [22] 岁，小孩[2]个
            //ab
            //[张三] 年龄 [33] 岁，小孩[3]个
            //oth
            //[李四] 年龄 [44] 岁，小孩[4]个
            //[王五] 年龄 [55] 岁，小孩[5]个
            //[赵六] 年龄 [66] 岁，小孩[6]个


            int n = 0;
            object[,] arr = new object[100, 11];
            var with_1 = new Regex("\\|.*?\\|");
            MatchCollection col1 = with_1.Matches(str);
            foreach (Match mm1 in col1)
            {
                arr[n, 0] = mm1.Value;
                n = n + 1;
            }
            app.Range["a1"].get_Resize(n, 1).Value2 = arr;
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;

            string str = System.Windows.Forms.Clipboard.GetText();

            List<List<object>> array = new List<List<object>>();

            string[] items = str.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var item in items)
            {
                var with_1 = new Regex("\\[.*?\\]");
                MatchCollection col1 = with_1.Matches(item);
                List<object> list = new List<object>();
                foreach (Match mm1 in col1)
                {
                    list.Add(mm1.ToString().Substring(mm1.ToString().LastIndexOf('[') + 1, mm1.ToString().LastIndexOf(']') - 1 - mm1.ToString().LastIndexOf('[')));
                }
                if (col1.Count > 0)
                {
                    array.Add(list);
                }
            }
            if (array.Count > 0)
            {
                object[,] value = new object[array.Count, array[0].Count];

                for (int i = 0; i < array.Count; i++)
                {
                    for (int j = 0; j < array[0].Count; j++)
                    {
                        value[i, j] = array[i][j];
                    }
                }
                //object[,] arr = { { "张三", 22, 2 }, { "李四", 33, 3 }, { "赵五", 44, 4 } };
                //app.Range["a1"].get_Resize(1, 1).Value2 = 1;
                app.Range["a1"].get_Resize(array.Count, 3).Value2 = value;
                //app.Range["b1"].get_Resize(3, 2).Select();
                //app.Range["a10"].get_Resize(3, 3).Value2 = app.WorksheetFunction.Transpose(array.ToArray());
            }
        }
    }
}
