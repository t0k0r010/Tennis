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

            DataLabel AttackAngleMe;    //自分に対する攻撃角度
            DataLabel AttackAngleOther; //相手に対する攻撃角度

            public RallyDataSheet(Excel.Worksheet sheet, int labelRowHeight)
                : base(sheet, labelRowHeight)
            {
                IsHitter = false;

                AttackAngleOther = new DataLabel("AS", "BB", LabelRowHeight+1);
                AttackAngleMe = new DataLabel("BC", "BL", +LabelRowHeight+1);
            }

            public override void SetPosition(string time, PointF point)
            {

            }

            public void EndRally()
            {
                Court court = Court.Instance;
                int rallyNum = court.RallyNum;

                var shots = court.ShotDirectionsInRally;
                var shotAngleForOther = court.ShotAngleForOther;
                var shotAngleForMe = court.ShotAngleForMe;
                var hitterPoint = court.Positions_p[Court.Surveyed.HitterPos];

                //バウンドを見ている場合はやらなくていい
                if (hitterPoint.Count == 0)
                    return;

                // 0 : サーバー, 1 : レシーバー
                double[] SumAng = {0,0};
                double[] MaxAng = { 0, 0};
                double[] MinAng = {1000, 1000};

                //プレイヤーAはサーバーかレシーバか
                int Aindex = 0;

                //サーバーがコートの下側で, 上側のプレイヤーがAのとき
                //もしくは,サーバーがコートの上側で,上側のプレイヤーがBの時Aはレシーバー
                if ((hitterPoint[0].Y < 0 && court.UpperPlayer == Court.Players.PlayerA) ||
                    (hitterPoint[0].Y > 0 && court.UpperPlayer == Court.Players.PlayerB))
                    Aindex = 1;

                int Bindex = (Aindex + 1) % 2;  //Aと逆

                //相手に対する角度を出す
                for(int i=0; i<court.ShotAngleForOther.Count; i++)
                {
                    //偶数番目がレシーバー
                    int k = 1 - i % 2;
                    var deg = court.ShotAngleForOther[i];
                    SumAng[k] += deg;

                    if (MaxAng[k] < deg)
                        MaxAng[k] = deg;

                    if (MinAng[k] > deg)
                        MinAng[k] = deg;
                }
                //書き出し
                var cnt = court.ShotAngleForOther.Count;
                int serverCnt   =  cnt % 2 == 0 ? cnt/2 : (cnt+1)/2;
                int recieverCnt = cnt % 2 == 0 ? cnt / 2 : (cnt - 1) / 2;
                int[] counts = { serverCnt, recieverCnt };
                Excel.Range r = AttackAngleOther.GetRange(Sheet);
                if( counts[Aindex] > 0)
                {
                    r.get_Range("A1").Value2 = SumAng[Aindex];   //合計
                    r.get_Range("C1").Value2 = SumAng[Aindex] / counts[Aindex];
                    r.get_Range("E1").Value2 = MaxAng[Aindex];
                    r.get_Range("G1").Value2 = MinAng[Aindex];
                    r.get_Range("I1").Value2 = MaxAng[Aindex] - MinAng[Aindex];
                }
                if( counts[Bindex] > 0)
                {
                    r.get_Range("B1").Value2 = SumAng[Bindex];
                    r.get_Range("D1").Value2 = SumAng[Bindex] / counts[Bindex];
                    r.get_Range("F1").Value2 = MaxAng[Bindex];
                    r.get_Range("H1").Value2 = MinAng[Bindex];
                    r.get_Range("J1").Value2 = MaxAng[Bindex] - MinAng[Bindex];

                }

                AttackAngleOther.Row++;

                //自分に対する角度を出す
                for (int i = 0; i < court.ShotAngleForMe.Count; i++ )
                {
                    //偶数番目がサーバー
                    int k = i % 2;

                    var deg = court.ShotAngleForMe[i];
                    SumAng[k] += deg;

                    if (MaxAng[k] < deg)
                        MaxAng[k] = deg;

                    if (MinAng[k] > deg)
                        MinAng[k] = deg;
                }

                cnt = court.ShotAngleForMe.Count;
                serverCnt = cnt % 2 == 0 ? cnt / 2 : (cnt + 1) / 2;
                recieverCnt = cnt % 2 == 0 ? cnt / 2 : (cnt - 1) / 2;
                counts = new int[]{serverCnt, recieverCnt};
                r = AttackAngleMe.GetRange(Sheet);

                if( counts[Aindex] > 0)
                {
                    r.get_Range("A1").Value2 = SumAng[Aindex];   //合計
                    r.get_Range("C1").Value2 = SumAng[Aindex] / counts[Aindex];
                    r.get_Range("E1").Value2 = MaxAng[Aindex];
                    r.get_Range("G1").Value2 = MinAng[Aindex];
                    r.get_Range("I1").Value2 = MaxAng[Aindex] - MinAng[Aindex];
                }

                if( counts[Bindex] > 0)
                {
                    r.get_Range("B1").Value2 = SumAng[Bindex];
                    r.get_Range("D1").Value2 = SumAng[Bindex] / counts[Bindex];
                    r.get_Range("F1").Value2 = MaxAng[Bindex];
                    r.get_Range("H1").Value2 = MinAng[Bindex];
                    r.get_Range("J1").Value2 = MaxAng[Bindex] - MinAng[Bindex];
                }

                AttackAngleMe.Row++;
                Excel.Range range = Sheet.get_Range("A" + AttackAngleMe.Row, "CN" + AttackAngleMe.Row);
                range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium;
            }

            public void Update()
            {
                Court court  = Court.Instance;
                int rallyNum = court.RallyNum;

                var shots = court.ShotDirectionsInRally;

                //ラリーが2回未満なら角度が出せない
                if (shots.Count < 2)
                    return;

                //2つのベクトルの角度を求める
                var angleOther = System.Windows.Vector.AngleBetween(shots[shots.Count - 1], shots[shots.Count - 2]);
                var rangeOther = AttackAngleOther.GetRange(Sheet);
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
                PlayerPos   = new DataLabel("AA", "AD", LabelRowHeight+1);
                BoundPos    = new DataLabel( "Q",  "R", LabelRowHeight+1);
                AttackAngle = new DataLabel("AL", "AL", LabelRowHeight + 1);
            }

            public override void SetPosition(string time, PointF point)
            {
                var court = Court.Instance;
                switch(court.LastSurveyed)
                {
                    case Court.Surveyed.BoundPos:
                        base.SetPosition(BoundPos.GetRange(Sheet), time, point);
                        BoundPos.Row++;
                        break;
                    case Court.Surveyed.HitterPos:
                        base.SetPosition(PlayerPos.GetRange(Sheet).get_Range("A1", "B1"), time, point);
                        break;
                    case Court.Surveyed.RecieverPos:
                        base.SetPosition(PlayerPos.GetRange(Sheet).get_Range("C1", "D1"), time, point);
                        PlayerPos.Row++;
                        break;
                }
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
                base.SetPosition(Sheet.get_Range(BoundPos.LeftCol + BoundPos.Row, BoundPos.RightCol + BoundPos.Row), time, p);
                BoundPos.Row++;
            }

            void SetPlayerPosition(string time, PointF p)
            {
                Excel.Range r = Sheet.get_Range(PlayerPos.LeftCol + PlayerPos.Row, PlayerPos.RightCol + PlayerPos.Row);

                Excel.Range range = IsHitter ? r.get_Range("A1" , "B1") : r.get_Range("C1" , "D1");
                base.SetPosition(range, time, p);
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

                shotSheet  = new ShotDataSheet(wb.Sheets[1], 4);
                rallySheet = new RallyDataSheet(wb.Sheets[3], 3);
            }
            catch (Exception ex)
            {
                wb.Close(false);
                app.Quit();
                throw ex;
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
