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

    class RallyDataMaker
    {
        static Excel.Application app;
        static Excel.Workbook wb;
        public static void MakeRally()
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
                    if( System.IO.Path.GetExtension(fileName) != ".xlsx" )
                    {
                        MessageBox.Show(".xlsx のファイルを選択してください");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ファイルを開くことができませんでした.");
                    return;
                }

                //エクセルアプリを開く
                app = new Excel.Application();
                app.Visible = true;

                //新しいファイルを開く
                wb = app.Workbooks.Open(Filename: fileName);

                try
                {
                    string tmpFileName = System.IO.Directory.GetCurrentDirectory() + "/template.xlsx";
                    var fromwb = app.Workbooks.Open(Filename: tmpFileName, ReadOnly: true);

                    var shotSheet = wb.Sheets[1];
                    fromwb.Sheets[3].Copy(After:shotSheet); //shotSheetの後にコピーする
                    ShotToRally(shotSheet, wb.Sheets[2]);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("変換できませんでした");
                    wb.Close(false);
                    app.Quit();
                    return;
                }
            }
        }

        static List<int> PointStarts = new List<int>();
        static Excel.Worksheet ShotToRally(Excel.Worksheet shotSheet, Excel.Worksheet rallySheet)
        {
            const string Left = "Q", Right = "AD";
            const string Service = "K";

            var pgDiag = new ProgressDialog();
            pgDiag.Show();
            pgDiag.Pg.Minimum = 0;
            pgDiag.Pg.Maximum = shotSheet.UsedRange.Rows.Count;
            pgDiag.Pg.Value = 0;
            PointStarts.Clear();

            for (int row = 1; row <= shotSheet.UsedRange.Rows.Count; row++ )
            {
                Excel.Range range = shotSheet.get_Range(Left + row.ToString());
                var topLine = (Excel.XlLineStyle)range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle;
                var botLine = (Excel.XlLineStyle)range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle;

                if (topLine != Excel.XlLineStyle.xlLineStyleNone && !PointStarts.Contains(row))
                {
                    PointStarts.Add(row);
                }

                if (botLine != Excel.XlLineStyle.xlLineStyleNone && !PointStarts.Contains(row + 1))
                {
                    PointStarts.Add(row + 1);
                }

                pgDiag.Pg.Value = row;
            }

            pgDiag.Text = "計算中";
            pgDiag.Pg.Value = 0;

            const int StartRow = 4;
            for(int i=0; i<=20; i+=2)
            {
                int left = DataSheet.ColToInt("AS") + i;
                string col = DataSheet.IntToCol(left) ;
                var range = rallySheet.get_Range(col + StartRow, col + (PointStarts.Count + 1));
                range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThin;
            }

            for (int i = 0; i < PointStarts.Count-1; i++ )
            {
                Excel.Range shotRange = shotSheet.get_Range(Left + PointStarts[i], Right + (PointStarts[i + 1]-1));

                int row = StartRow + i;
                Excel.Range rallyRange = rallySheet.get_Range("AS" + row, "BL" + row);
                
                bool serverIsA = shotSheet.get_Range(Service + PointStarts[i]).Value2 != null;
                MakeLine(shotRange, rallyRange, serverIsA, PointStarts[i + 1] - PointStarts[i]);

                Excel.Range range = rallySheet.get_Range("A" + row, "BM" + row);
                range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThin;

                pgDiag.Pg.Value = PointStarts[i];
            }

            pgDiag.Close();
            MessageBox.Show("変換完了");

            return null;
        }

         static void MakeLine(Excel.Range shotRange, Excel.Range rallyRange, bool serverIsA, int rowNum)
        {
            List<RallyInfo> rallys = new List<RallyInfo>();
            for(int i=1; i<=rowNum; i++)
            {
                RallyInfo info = new RallyInfo( shotRange.get_Range("A"+i, "N"+i) );
                rallys.Add(info);
            }

            if (rallys.Count < 3)
                return;

            List<System.Windows.Vector> shotVector = new List<System.Windows.Vector>();
            for(int i=1; i < rallys.Count; i++)
            {
                if (rallys[i].HitterPos == null)
                    break;

                // 打者がnullになる可能性があるのは最後だけなので[i-1]のnullチェックは必要ない
                System.Windows.Vector vec = rallys[i].HitterPos.v - rallys[i-1].HitterPos.v;
                shotVector.Add(vec);
            }

            // 0: サーバー, 1:レシーバー
            AngleInfo[] OtherAng = new AngleInfo[2];
            AngleInfo[] MeAng = new AngleInfo[2];

             //個々のインスタンスを作成
            for (int i = 0; i < 2; i++ )
            {
                OtherAng[i] = new AngleInfo();
                MeAng[i] = new AngleInfo();
            }

            for (int i = 1; i < shotVector.Count; i++)
            {
                var vec1 = shotVector[i];
                var vec2 = -shotVector[i - 1];

                int index = i % 2; //サーバーかレシーバか

                double angForOther = Math.Abs(System.Windows.Vector.AngleBetween(vec1, vec2));
                OtherAng[index].Update(angForOther);

                if (i == 1)
                    continue;

                var vec3 = shotVector[i - 2];
                double angForMe = Math.Abs(System.Windows.Vector.AngleBetween(vec1, vec3));

                MeAng[index].Update(angForMe);
            }

            Excel.Range AngOtherRangeA = rallyRange.get_Range("A1", "I1");
            Excel.Range AngOtherRangeB = rallyRange.get_Range("B1", "J1");  //一つずれる

            Excel.Range AngMeRangeA = rallyRange.get_Range("K1", "S1");
            Excel.Range AngMeRangeB = rallyRange.get_Range("L1", "T1");  //一つずれる

            if(serverIsA)
            {
                OtherAng[0].Write(AngOtherRangeA);
                OtherAng[1].Write(AngOtherRangeB);

                MeAng[0].Write(AngMeRangeA);
                MeAng[1].Write(AngMeRangeB);
            }
            else
            {
                OtherAng[1].Write(AngOtherRangeA);
                OtherAng[0].Write(AngOtherRangeB);

                MeAng[1].Write(AngMeRangeA);
                MeAng[0].Write(AngMeRangeB);
            }
        }

         class RallyInfo
         {
             public Vec2 BoundPos;
             public Vec2 WinnerPos;
             public Vec2 MissPos;
             public Vec2 HitterPos;
             public Vec2 RecieverPos;

             public RallyInfo(Excel.Range range)
             {
                 BoundPos    = GetVector(range.get_Range("A1", "B1"));
                 WinnerPos   = GetVector(range.get_Range("C1", "D1"));
                 MissPos     = GetVector(range.get_Range("E1", "F1"));
                 HitterPos   = GetVector(range.get_Range("K1", "L1"));
                 RecieverPos = GetVector(range.get_Range("M1", "N1"));
             }

             //nullを許容するためにラップしている
             public class Vec2
             {
                 public System.Windows.Vector v;
             }

             static RallyInfo.Vec2 GetVector(Excel.Range towCols)
             {
                 var Xcell = towCols.get_Range("A1").Value2;
                 var Ycell = towCols.get_Range("B1").Value2;
                 if ( Xcell == null || Ycell == null)
                     return null;

                 try
                 {
                     double X = (double)(Xcell);
                     double Y = (double)(Ycell);
                     RallyInfo.Vec2 res = new RallyInfo.Vec2();
                     res.v = new System.Windows.Vector(X, Y);
                     return res;
                 }catch(System.NullReferenceException)
                 {
                     throw;
                 }
                 catch
                 {
                     Console.WriteLine("Catch NULL Exp");
                     return null;
                 }
             }
         }

        class AngleInfo
        {
            public double sum { get; private set; }
            public int cnt { get; private set; }
            public double max { get; private set; }
            public double min { get; private set; }

            public AngleInfo()
            {
                sum = 0;
                cnt = 0;
                max = -1000;
                min = 1000;
            }

            public void Update(double deg)
            {
                sum += deg;
                cnt++;

                if (max < deg)
                    max = deg;

                if (min > deg)
                    min = deg;
            }

            public void Write(Excel.Range range)
            {
                if (cnt == 0)
                    return;
                
                range.get_Range("A1").Value2 = sum;
                range.get_Range("C1").Value2 = sum / (double)cnt;
                range.get_Range("E1").Value2 = max;
                range.get_Range("G1").Value2 = min;
                range.get_Range("I1").Value2 = max - min;
            }
        }
    }
}
