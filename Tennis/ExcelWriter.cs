using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Drawing;
using Microsoft;
using Microsoft.Office;
using Microsoft.Office.Interop;

//エクセルファイルへの書き出しを行うクラス
//どんなクラスからもアクセスできるが,このクラスは誰も保持しない.
namespace Tennis
{
    using Excel = Microsoft.Office.Interop.Excel;
    class ExcelWriter
    {
        static ExcelWriter instance = null;
        public static ExcelWriter Instance
        {
            get;
            private set;
        }

        Excel.Application app;
        Excel.Workbook wb;

        public class RallyDataSheet : DataSheet
        {
            public bool IsHitter { get; private set; }
            public RallyDataSheet(Excel.Worksheet sheet, int labelRowHeight)
                : base(sheet, labelRowHeight)
            {
                IsHitter = false;
            }

            public override void SetPosition(string time, PointF point)
            {
            }
        };

        public class ShotDataSheet : DataSheet
        {
            //調査対象
            public enum Surveyed
            {
                BoundPos,   //バウンド位置
                PlayerPos   //プレイヤー位置
            };
            public Surveyed Surveying { get; set; }
            public bool IsHitter { get; private set; }

            DataLabel BoundPos, PlayerPos, AttackAngle;

            int rallyNum = 0;   //ラリー回数

            public ShotDataSheet(Excel.Worksheet sheet, int labelRowHeight)
                : base(sheet, labelRowHeight)
            {
                IsHitter = true;
                PlayerPos   = new DataLabel("AD", "AG", LabelRowHeight+1);
                BoundPos    = new DataLabel("AB", "AC", LabelRowHeight+1);
                AttackAngle = new DataLabel("AL", "AL", LabelRowHeight + 1);
            }

            public override void SetPosition(string time, PointF point)
            {
                if (Surveying == Surveyed.BoundPos)
                    SetBoundPosition(time, point);
                else
                    SetPlayerPosition(time, point);
            }

            //ラリーの終わりの表す線を書き込む
            public void EndRally()
            {
                int row = Surveying == Surveyed.BoundPos? BoundPos.Row : PlayerPos.Row;
                Excel.Range range = Sheet.get_Range("A" + row, "CN" + row);
                range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium;

                rallyNum = 0;
            }

            void SetBoundPosition(string time, PointF p)
            {
                SetPosition(Sheet.get_Range(BoundPos.LeftCol + BoundPos.Row, BoundPos.RightCol + BoundPos.Row), time, p);
                BoundPos.Row++;
            }

            void SetPlayerPosition(string time, PointF p)
            {
                Excel.Range r = Sheet.get_Range(PlayerPos.LeftCol + PlayerPos.Row, PlayerPos.RightCol + PlayerPos.Row);

                Excel.Range range = IsHitter ? r.get_Range("A1" , "B1") : r.get_Range("C1" , "D1");
                SetPosition(range, time, p);

                IsHitter = !IsHitter;

                if (!IsHitter){
                    // 2ショット以上続けば角度が出せる
                    if (rallyNum >= 2)
                    {
                        SetAttackAngle();
                    }
                }

                //被打選手の設定が終われば次の行へ
                if (IsHitter)
                {
                    PlayerPos.Row++;                    
                    rallyNum++;
                }
            }

            void SetAttackAngle()
            {
                Excel.Range range = Sheet.get_Range( AttackAngle.LeftCol + (PlayerPos.Row - 1) );

                string l = PlayerPos.LeftCol;
                string r = IntToCol(ColToInt(PlayerPos.LeftCol) + 1);

                string ax = l + (PlayerPos.Row-2);
                string ay = r + (PlayerPos.Row - 2);
                string bx = l + (PlayerPos.Row - 1);
                string by = r + (PlayerPos.Row - 1);
                string cx = l + (PlayerPos.Row - 0);
                string cy = r + (PlayerPos.Row - 0);
                string Ax = "(" + ax + "-" + bx + ")";  //1本目のベクトル
                string Ay = "(" + ay + "-" + by + ")";
                string Bx = "(" + cx + "-" + bx + ")";  //2本目のベクトル
                string By = "(" + cy + "-" + by + ")";

                string absBA = "SQRT( " + Ax+ "^2+" + Ay + "^2)";
                string absCB = "SQRT( " + Bx+ "^2+" + By + "^2)";
                string cos = "(" + Ax + "*" + Bx + "+" + Ay + "*" + By + ") / " + absBA + " / " + absCB;

                range.Value2 = "=" + cos;
            }
        };

        public ShotDataSheet shotSheet { get; private set; }

        public RallyDataSheet rallySheet { get; private set; }

        //ファイルを開く
        public static void Open()
        {
            //すでにファイルを開いているときは,新しく開きなおすか確認する.
            if ( Available() )
            {
                if (MessageBox.Show("今あるファイルを閉じて別のファイルを開きますか?", "確認", 
                    MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                    return;

                Instance.Close();
            }

            //ファイルを開く
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "エクセルファイルを開く";

            //開くボタンを押したとき
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string fileName = dialog.FileName;

                try
                {
                    Instance = new ExcelWriter(fileName);
                }
                catch(Exception ex)
                {
                    MessageBox.Show( ex.Message, "開けませんでした");
                    Instance = null;
                }
            }
        }

        //エクセルに書き込みが可能か調べる
        public static bool Available()
        {
            if (Instance == null || Instance.app == null)
                return false;

            return Instance.app.Visible;
        }

        ExcelWriter(string fileName)
        {
            //エクセルアプリを開く
            app = new Excel.Application();
            app.Visible = true;

            //新しいファイルを開く
            wb = app.Workbooks.Open(Filename: fileName);

            try
            {
                CopyTemplate(app, wb);  //新しく作る場合

                rallySheet = new RallyDataSheet(wb.Sheets[2], 3);
                shotSheet = new ShotDataSheet(wb.Sheets[1], 4);
            }
            catch (Exception ex)
            {
                wb.Close(false);
                app.Quit();
                throw new Exception();
            }
        }

        //新しく作成した場合はテンプレートからコピーする
        void CopyTemplate(Excel.Application ap, Excel.Workbook wb)
        {
            try
            {
                List<string> sheets = new List<string>();
                //元のブックにあるシートの名前を保存
                foreach (Excel.Worksheet ws in wb.Sheets)                
                    sheets.Add(ws.Name);                   
                
                string fileName = System.IO.Directory.GetCurrentDirectory() + "/template.xlsx";
                var fromwb = ap.Workbooks.Open(Filename: fileName, ReadOnly: true);
  
                foreach (Excel.Worksheet ws in wb.Sheets)
                    if (ws.Name == ((Excel.Worksheet)(fromwb.Sheets[1])).Name)
                        return;

                for(int i=1; i <= fromwb.Sheets.Count; i++)                
                    fromwb.Sheets[i].Copy(wb.Sheets[i]);                

                //もともとあったシートは削除
                foreach (var s in sheets){
                    Console.WriteLine(s);
                    var sheet = (Excel.Worksheet)wb.Sheets[s.ToString()];
                    sheet.Delete();
                }
                fromwb.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //エクセルを閉じたときに呼ばれる
        void ClosedEvent(object sender, EventArgs e)
        {
            Instance = null;
        }

        //エクセルを閉じる
        void Close()
        {
            if (app != null)
            {
                app.Quit();
            }
        }

    }

}
