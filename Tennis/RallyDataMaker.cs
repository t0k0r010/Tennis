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
                    if (System.IO.Path.GetExtension(fileName) != ".xlsx")
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
                    fromwb.Sheets[3].Copy(After: shotSheet); //shotSheetの後にコピーする
                    fromwb.Close();
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
            const string ShotLeft = "Q", ShotRight = "AD";
            const string Service = "K";

            var pgDiag = new ProgressDialog();
            pgDiag.Show();
            pgDiag.Pg.Minimum = 0;
            pgDiag.Pg.Maximum = shotSheet.UsedRange.Rows.Count;
            pgDiag.Pg.Value = 0;
            PointStarts.Clear();

            for (int row = 1; row <= shotSheet.UsedRange.Rows.Count; row++)
            {
                Excel.Range range = shotSheet.get_Range(ShotLeft + row.ToString());
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

            int rallyLastColumn  = rallySheet.UsedRange.Columns.Count;
            int rallyBeginColumn = DataSheet.ColToInt("AS");
            for (int i = rallyBeginColumn; i <= rallyLastColumn; i += 2)
            {
                string col = DataSheet.IntToCol(i);
                var range = rallySheet.get_Range(col + StartRow, col + (PointStarts.Count + 1));

                if( (i-rallyBeginColumn) % (DataInfo.NeedCols+1) == 0)
                {
                    range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThick;
                }
                else
                {
                    range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThin;
                }
            }

            int debugLine = 0;
            bool lastServiceIsA = false;
            for (int i = 0; i < PointStarts.Count - 1; i++)
            {
                debugLine += PointStarts[i];
                Excel.Range shotRange = shotSheet.get_Range(ShotLeft + PointStarts[i], ShotRight + (PointStarts[i + 1] - 1));

                int row = StartRow + i;
                Excel.Range rallyRange = rallySheet.get_Range("AS" + row, "BP" + row);

                bool serverIsA = shotSheet.get_Range(Service + PointStarts[i]).Value2 != null;

                try
                {
                    MakeLine(shotRange, rallyRange, serverIsA, PointStarts[i + 1] - PointStarts[i]);
                }
                catch
                {
                    MessageBox.Show("Error at " + debugLine + " to " + (debugLine + PointStarts[i + 1]));
                    pgDiag.Close();
                    throw;
                }

                if (i > 0 && lastServiceIsA != serverIsA)
                {
                    Excel.Range range = rallySheet.get_Range("A" + row, "BM" + row);
                    range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThin;
                }

                lastServiceIsA = serverIsA;
                pgDiag.Pg.Value = PointStarts[i];
            }

            pgDiag.Close();
            MessageBox.Show("変換完了");


            return null;
        }

        static double GetArea(System.Windows.Vector p1, System.Windows.Vector p2, System.Windows.Vector p3)
        {
            var v1 = p1 - p2;
            var v2 = p3 - p2;
            double cross = System.Windows.Vector.CrossProduct(v1, v2);
            return Math.Abs(cross) / 2.0;
        }

        //バウンド座標を用いた小さい面積を取得
        // boundPos : バウンド座標, hitterPos : その球の打点, otherPos : 情報によって変わるもう一つの座標
        // hitterPos, boundPosと 
        // y = boundPos.Y の直線と, hitterPosとotherPosを結ぶ直線の交点による面積を返す
        static double GetSmallArea(System.Windows.Vector boundPos, System.Windows.Vector hitterPos, System.Windows.Vector otherPos)
        {
            var delta = (boundPos.Y - otherPos.Y) / (hitterPos.Y - otherPos.Y);
            var vec = hitterPos - otherPos;
            return GetArea(boundPos, hitterPos, otherPos + vec * delta);
        }

        delegate Excel.Range del(int leftInt);
        static void MakeLine(Excel.Range shotRange, Excel.Range rallyRange, bool serverIsA, int rowNum)
        {
            List<RallyInfo> rallys = new List<RallyInfo>();
            for (int i = 1; i <= rowNum; i++)
            {
                RallyInfo info = new RallyInfo(shotRange.get_Range("A" + i, "N" + i));
                rallys.Add(info);
            }

            if (rallys.Count < 3)
                return;

            // 0: サーバー, 1:レシーバー
            DataInfo[] OtherHitterAng = new DataInfo[2];    //相手のショットに対する打点座標を用いる角度
            DataInfo[] OtherReciveAng = new DataInfo[2];    //相手の被打点座標を用いる角度
            DataInfo[] MeHitterAng = new DataInfo[2];    //自分のショットに対する打点座標を用いる角度
            DataInfo[] OtherHitBigArea = new DataInfo[2];   //相手打点座標による大きい攻撃面積
            DataInfo[] OtherRecBigArea = new DataInfo[2];   //相手被打点座標による大きい攻撃面積
            DataInfo[] OtherHitSmallArea = new DataInfo[2]; //相手打点座標による小さい攻撃面積
            DataInfo[] OtherRecSmallArea = new DataInfo[2]; //相手被打点座標による小さい攻撃面積

            del lambda = leftInt => rallyRange.get_Range(DataSheet.IntToCol(leftInt)                    + 1,
                                                         DataSheet.IntToCol(leftInt + DataInfo.NeedCols) + 1);
            //個々のインスタンスを作成
            for (int i = 0; i < 2; i++)
            {
                int offset = serverIsA ? i : (i + 1) % 2;
                int leftInt = DataSheet.ColToInt("A") + offset;
                OtherHitterAng[i] = new DataInfo( lambda( leftInt ) );

                leftInt += DataInfo.NeedCols + 1;
                OtherReciveAng[i] = new DataInfo( lambda( leftInt ) );

                leftInt += DataInfo.NeedCols + 1;
                MeHitterAng[i] = new DataInfo(lambda(leftInt));

                leftInt += DataInfo.NeedCols + 1;
                OtherHitBigArea[i] = new DataInfo(lambda(leftInt));

                leftInt += DataInfo.NeedCols + 1;
                OtherHitSmallArea[i] = new DataInfo(lambda(leftInt));

                leftInt += DataInfo.NeedCols + 1;
                OtherRecBigArea[i] = new DataInfo(lambda(leftInt));

                leftInt += DataInfo.NeedCols + 1;
                OtherRecSmallArea[i] = new DataInfo(lambda(leftInt));
            }


            for (int i = 1; i < rallys.Count; i++)
            {
                if (rallys[i].HitterPos == null)
                {
                    if (i != rallys.Count - 1)
                        Console.WriteLine((rallys.Count - 1 - i) + " from Last is NULL");
                    break;
                }

                //サーバーの角度(面積)かレシーバのかを決めるインデックス
                int index_odd_0 = (i + 1) % 2;   //奇数番が0になるインデックス
                int index_even_0 = i % 2;        //偶数番が0になるインデックス

                var current = rallys[i];
                var prev1 = rallys[i - 1];
                var vecHtoH_c = current.HitterPos.v - prev1.HitterPos.v;   //前の打者 -> 今の打者  へのベクトル
                var vecHtoR_p1 = prev1.RecieverPos.v - prev1.HitterPos.v;   //前の打者 -> 前の被打者へのベクトル

                //相手の被打点座標を用いる角度
                double angForOtherH = Math.Abs(System.Windows.Vector.AngleBetween(vecHtoH_c, vecHtoR_p1));
                OtherReciveAng[index_odd_0].Update(angForOtherH);

                //相手被打点座標による大きい攻撃面積
                OtherRecBigArea[index_odd_0].Update(GetArea(current.HitterPos.v, prev1.HitterPos.v, prev1.RecieverPos.v));

                if (current.BoundPos != null)
                {
                    //相手被打点座標による小さい攻撃面積
                    OtherRecSmallArea[index_odd_0].Update(GetSmallArea(current.BoundPos.v, prev1.HitterPos.v, prev1.RecieverPos.v));
                }

                if (i < 2)
                    continue;

                //相手のショットに対する,打点座標を用いる角度
                var prev2 = rallys[i - 2];
                var vecHtoH_p1 = prev1.HitterPos.v - prev2.HitterPos.v;
                double angForOtherR = Math.Abs(System.Windows.Vector.AngleBetween(vecHtoH_c, -vecHtoH_p1));
                OtherHitterAng[index_odd_0].Update(angForOtherR);

                // 相手打点座標による大きい攻撃面積
                OtherHitBigArea[index_odd_0].Update(GetArea(current.HitterPos.v, prev1.HitterPos.v, prev2.HitterPos.v));

                if(current.BoundPos != null)
                {
                    //相手打点座標による小さい攻撃面積
                    OtherHitSmallArea[index_odd_0].Update(GetSmallArea(current.BoundPos.v, prev1.HitterPos.v, prev2.HitterPos.v));
                }

                if (i < 3)
                    continue;

                var prev3 = rallys[i - 3];
                var vecHtoH_p2 = prev2.HitterPos.v - prev3.HitterPos.v;
                double angForMeH = Math.Abs(System.Windows.Vector.AngleBetween(vecHtoH_p2, vecHtoR_p1));
                MeHitterAng[index_odd_0].Update(angForMeH);
            }

            for(int i=0; i<2; i++)
            {
                OtherHitterAng[i].Write();
                OtherReciveAng[i].Write();
                MeHitterAng[i].Write();
                OtherHitBigArea[i].Write();
                OtherHitSmallArea[i].Write();
                OtherRecBigArea[i].Write();
                OtherRecSmallArea[i].Write();
            }
            /*
            // A プレイヤーA要
            Excel.Range AngOtherRangeA = rallyRange.get_Range("A1", "K1");
            Excel.Range AngOtherRangeB = rallyRange.get_Range("B1", "L1");  //一つずれる

            Excel.Range AngMeRangeA = rallyRange.get_Range("M1", "W1");
            Excel.Range AngMeRangeB = rallyRange.get_Range("N1", "X1");  //一つずれる

            Excel.Range[] AngOtherRange = new Excel.Range[2];
            Excel.Range[] AngMeRange = new Excel.Range[2];
            if (serverIsA)
            {
                OtherHitterAng[0].Write(AngOtherRangeA);
                OtherHitterAng[1].Write(AngOtherRangeB);

                OtherReciveAng[0].Write(AngMeRangeA);
                OtherReciveAng[1].Write(AngMeRangeB);
            }
            else
            {
                OtherHitterAng[1].Write(AngOtherRangeA);
                OtherHitterAng[0].Write(AngOtherRangeB);

                OtherReciveAng[1].Write(AngMeRangeA);
                OtherReciveAng[0].Write(AngMeRangeB);
            }*/
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
                BoundPos = GetVector(range.get_Range("A1", "B1"));
                WinnerPos = GetVector(range.get_Range("C1", "D1"));
                MissPos = GetVector(range.get_Range("E1", "F1"));
                HitterPos = GetVector(range.get_Range("K1", "L1"));
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
                if (Xcell == null || Ycell == null)
                    return null;

                try
                {
                    double X = (double)(Xcell);
                    double Y = (double)(Ycell);
                    RallyInfo.Vec2 res = new RallyInfo.Vec2();
                    res.v = new System.Windows.Vector(X, Y);
                    return res;
                }
                catch (System.NullReferenceException)
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

        //合計,平均,最大最少,標準偏差を格納するためのクラス
        class DataInfo
        {
            public double sum { get; private set; }
            public double max { get; private set; }
            public double min { get; private set; }
            List<double> datas = new List<double>();

            public const int NeedCols = 11;

            Excel.Range range;
            public DataInfo(Excel.Range r)
            {
                range = r;
                sum = 0;
                max = -1000;
                min = 1000;
            }

            public void Update(double data)
            {
                sum += data;

                if (max < data)
                    max = data;

                if (min > data)
                    min = data;

                datas.Add(data);
            }

            public void Write(/*Excel.Range range*/)
            {
                if (datas.Count == 0)
                    return;

                double mean = sum / (double)datas.Count;
                range.get_Range("A1").Value2 = sum;
                range.get_Range("C1").Value2 = mean;
                range.get_Range("E1").Value2 = max;
                range.get_Range("G1").Value2 = min;
                range.get_Range("I1").Value2 = max - min;


                //分散を計算
                double vari = 0;
                foreach (var deg in datas)
                    vari += Math.Pow(mean - deg, 2);

                range.get_Range("K1").Value2 = Math.Sqrt(vari / (double)datas.Count);
            }
        }
    }
}
