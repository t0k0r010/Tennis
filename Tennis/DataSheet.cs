using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Microsoft;
using Microsoft.Office;
using Microsoft.Office.Interop;
namespace Tennis
{
    using Excel = Microsoft.Office.Interop.Excel;
    abstract class DataSheet
    {
        //表のラベルの行数
        public int LabelRowHeight { get; private set; }
        //エクセルのワークシート
        public Excel.Worksheet Sheet { get; private set; }

        public DataSheet(Excel.Worksheet sheet, int labelRowHeight)
        {
            Sheet = sheet;
            LabelRowHeight = labelRowHeight;
        }

        protected void SetPosition(Excel.Range range, string time, PointF p)
        {
            range.get_Range("A1").Value2 = p.X;
            range.get_Range("B1").Value2 = p.Y;
        }

        public abstract void SetPosition(string time, PointF point);

        // エクセルの列番号(A, AA)を整数値に変換する
        public static int ColToInt(string col)
        {
            // A = 1と数える.
            int i = 0;
            foreach (var c in col)
            {
                i *= 26;
                i += (int)(c - 'A') + 1;
            }

            return i;
        }

        // 1以上の整数値を A ~ Z, AA, AB となるエクセルの列番号に変換する
        public static string IntToCol(int x)
        {
            if (x < 0)
            {
                //例外を投げる.
            }

            string res = "";
            x--;    //1オリジン
            while (x >= 0)
            {
                char a = (char)((x % 26) + 'A');
                res = a.ToString() + res;

                // 10進数(26) = 26進数(11) = AA とならなければならないので
                // 単純に 0 = A と対応するのではなく, 1 = A とする必要がある 
                // その為,桁を一つ下げた後に -1　している
                x = (int)(x / 26) - 1;
            }
            return res;
        }

        //各ラベル(列)
        public class DataLabel
        {
            public string LeftCol { get; private set; } //列(定数)
            public string RightCol { get; private set; }
            public int Row {get; set;}              //行
            public DataLabel(string left, string right, int row)
            {
                LeftCol = left;
                RightCol = right;
                Row = row;
            }
        }
    };
}
