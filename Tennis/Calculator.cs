using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft;
using Microsoft.Office;
using Microsoft.Office.Interop;
using System.Windows.Forms;

namespace Tennis
{
    using Excel = Microsoft.Office.Interop.Excel;
    using RallyInfo = RallyDataMaker.RallyInfo;

    class Calculator
    {
        public static OpenFileDialog ExcelOpenDialog()
        {
            //ファイルを開く
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "エクセルファイルを開く";
            dialog.CheckFileExists = true;
            //開くボタンを押したとき
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string fileName = dialog.FileName;
                try
                {
                    if (System.IO.Path.GetExtension(fileName) != ".xlsx")
                    {
                        MessageBox.Show(".xlsx のファイルを選択してください");
                        return null;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ファイルを開くことができませんでした.");
                    return null;
                }

                return dialog;
            }
            return null;
        }

       public static void CalcAngleFromShotData()
        {
           var dialog = ExcelOpenDialog();

           if (dialog == null)
               return;
           
           //エクセルアプリを開く
           var app = new Excel.Application();
           app.Visible = true;

           //新しいファイルを開く
           var workbook = app.Workbooks.Open(Filename: dialog.FileName);

           try
           {
               var shotSheet = workbook.Sheets[2];
               Calculate(shotSheet);
           }
           catch (Exception ex)
           {
               MessageBox.Show("変換できませんでした");
               workbook.Close(false);
               app.Quit();
               return;
           }
        }


       static void Calculate(Excel.Worksheet shotSheet)
       {
           const string InputLeftCulumn  = "Q";
           const string InputRightCulumn = "AD";
           const int InputTopRow = 5;
           const string Service = "K";

           //進捗バーの作成
           var pgDiag = new ProgressDialog();
           pgDiag.Show();
           pgDiag.Pg.Minimum = 0;
           pgDiag.Pg.Maximum = shotSheet.UsedRange.Rows.Count;
           pgDiag.Pg.Value   = 0;

           //ポイント始めとなる行番号を取得
           List<int> PointStartRows = new List<int>();
           for (int row = InputTopRow; row <= shotSheet.UsedRange.Rows.Count; row++)
           {
               Excel.Range range = shotSheet.get_Range(InputLeftCulumn + row.ToString());
               var topLine = (Excel.XlLineStyle)range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle;
               var botLine = (Excel.XlLineStyle)range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle;

               //ラインのひかれ方で判定
               if (topLine != Excel.XlLineStyle.xlLineStyleNone && !PointStartRows.Contains(row))
                   PointStartRows.Add(row);

               if (botLine != Excel.XlLineStyle.xlLineStyleNone && !PointStartRows.Contains(row + 1))
                   PointStartRows.Add(row + 1);

               pgDiag.Pg.Value = row;
           }

           pgDiag.Text = "計算中";
           pgDiag.Pg.Value = 0;

           const int StartRow = 4;

           for (int i = 0; i < PointStartRows.Count - 1; i++)
           {
               Excel.Range inputRange  = shotSheet.get_Range(InputLeftCulumn + PointStartRows[i], InputRightCulumn + (PointStartRows[i + 1] - 1));
               Excel.Range outputRange = shotSheet.get_Range("BQ" + PointStartRows[i], "DS" + (PointStartRows[i + 1] - 1));

               int row    = StartRow - PointStartRows[0] + PointStartRows[i];

               bool serverIsA = shotSheet.get_Range(Service + PointStartRows[i]).Value2 != null;

               try
               {
                   Console.WriteLine("row = " + (PointStartRows[i + 1] - PointStartRows[i]));
                   Calc(inputRange, outputRange);
               }
               catch
               {
                   MessageBox.Show("Error at " + PointStartRows[i] + " to " + PointStartRows[i + 1]);
                   pgDiag.Close();
                   throw;
               }
               pgDiag.Pg.Value = PointStartRows[i];
           }

           pgDiag.Close();
           MessageBox.Show("変換完了");
       }



       // hitter -> reciever
       // hitter -> bound   のベクトルの角度を返す
       static double ShotAngle(RallyInfo.Vec2 hitterPos, RallyInfo.Vec2 recieverPos, RallyInfo.Vec2 boundPos)
       {
           var v1 = recieverPos.v - hitterPos.v;
           var v2 = boundPos.v - hitterPos.v;

           return Math.Abs(System.Windows.Vector.AngleBetween(v1, v2));
       }


       static void SetInfo(Excel.Range range, RallyInfo preInfo, RallyInfo curInfo, RallyInfo nexInfo, bool isRecieve = false)
       {
           RallyInfo.Vec2 nextBoundPos = nexInfo.BoundPos != null ? nexInfo.BoundPos : nexInfo.WinnerPos;

           {
               var p = nexInfo.HitterPos != null ? nexInfo.HitterPos : nextBoundPos;

               if( p != null)
               {
                   var v2 = preInfo.HitterPos.v - curInfo.HitterPos.v;
                   //角度(打)
                   range.get_Range("A1").Value2 = Math.Abs(System.Windows.Vector.AngleBetween(p.v - curInfo.HitterPos.v, v2));

                   var v3 = curInfo.RecieverPos.v - curInfo.HitterPos.v;
                   //角度(被打)
                   range.get_Range("B1").Value2 = Math.Abs(System.Windows.Vector.AngleBetween(p.v - curInfo.HitterPos.v, v3));
               }
           }

           //大きい三角形面積(打, 打, 打)
           if(nexInfo.HitterPos != null)
               range.get_Range("C1").Value2 = RallyDataMaker.GetArea(preInfo.HitterPos.v, curInfo.HitterPos.v, nexInfo.HitterPos.v);


           //小さい三角形面積(打,打,打). ボレーの場合は無い
           if (nextBoundPos != null)
               range.get_Range("D1").Value2 = RallyDataMaker.GetSmallArea(nextBoundPos.v, curInfo.HitterPos.v, preInfo.HitterPos.v);

           //大きい攻撃面積(被打,打,打). ボレーの場合はない
           if (nexInfo.HitterPos != null)
               range.get_Range("E1").Value2 = RallyDataMaker.GetArea(curInfo.RecieverPos.v, curInfo.HitterPos.v, nexInfo.HitterPos.v);

           //小さい攻撃面積(被打,打,打)
           if (nextBoundPos != null)
               range.get_Range("F1").Value2 = RallyDataMaker.GetArea(nextBoundPos.v, curInfo.HitterPos.v, curInfo.RecieverPos.v);

           //動かされ距離
           {
               var move = preInfo.RecieverPos.v - curInfo.HitterPos.v; 
               range.get_Range("G1").Value2 = Math.Abs(move.X);
               range.get_Range("H1").Value2 = Math.Abs(move.Y);
               range.get_Range("I1").Value2 = move.Length;
           }

           //深さ. レシーブは次のバウンド,ラリーは現在のバウンドを用いる
           if(isRecieve == false){
               range.get_Range("J1").Value2 = Math.Abs(curInfo.BoundPos.Y);
           }
           else if(nextBoundPos != null)
           {
               range.get_Range("J1").Value2 = Math.Abs(nextBoundPos.Y);
           }
          
       }


       static void Calc(Excel.Range inputRange, Excel.Range outputRange)
       {
           //そのポイントの情報を取得
           List<RallyInfo> rallys = new List<RallyInfo>();
           //Console.WriteLine(inputRange.Rows.Count + "," + rowNum);

           for (int i = 1; i <= inputRange.Rows.Count; i++)
           {
               RallyInfo info = new RallyInfo(inputRange.get_Range("A" + i, "N" + i));
               rallys.Add(info);
           }

           //2以下だとエラー
           if (rallys.Count < 2)
               return;

           //ダブルフォルトの場合は何もない
           if (rallys[1].BoundPos == null)
               return;

           {
               //サーブ角
               outputRange.get_Range("A1").Value2 = ShotAngle(rallys[0].HitterPos, rallys[0].RecieverPos, rallys[1].BoundPos);

               if (rallys[1].HitterPos != null)
               {
                   //サーブ動かし距離
                   var move = rallys[1].HitterPos.v - rallys[0].RecieverPos.v;
                   outputRange.get_Range("B1").Value2 = Math.Abs(move.X);
                   outputRange.get_Range("C1").Value2 = Math.Abs(move.Y);
                   outputRange.get_Range("D1").Value2 = move.Length;
               }
               //サーブの速度


               //サーブの深さ
               outputRange.get_Range("F1").Value2 = Math.Abs(rallys[1].BoundPos.Y);
           }

           //サービスエースの場合,レシーブ以降の情報は無い
           if (rallys.Count < 3)
               return;


           //リターンの情報を設定
           {
               SetInfo(outputRange.get_Range("H2", "R2"), rallys[0], rallys[1], rallys[2], true);
           }

           //ラリーの情報を設定
           for (int i = 2; i < rallys.Count - 1; i++)
           {
               SetInfo(outputRange.get_Range("S" + (i + 1), "AG" + (i + 1)), rallys[i - 1], rallys[i], rallys[i + 1]);
           }
       }

    }
}
