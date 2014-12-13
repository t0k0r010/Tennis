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

        static Excel.Worksheet ShotToRally(Excel.Worksheet shotSheet, Excel.Worksheet rallySheet)
        {
            const string ShotLeft = "Q", ShotRight = "AD";
            const string Service = "K";

            Console.WriteLine(Enum.GetValues(typeof(Kinds)).Length);
            var pgDiag = new ProgressDialog();
            pgDiag.Show();
            pgDiag.Pg.Minimum = 0;
            pgDiag.Pg.Maximum = shotSheet.UsedRange.Rows.Count;
            pgDiag.Pg.Value = 0;
            List<int> PointStarts = new List<int>();
            Console.WriteLine(shotSheet.UsedRange.Rows.Count);
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

            //縦線を引く
            int rallyLastColumn  = rallySheet.UsedRange.Columns.Count;
            int rallyBeginColumn = DataSheet.ColToInt("AS");

            //データを計算する
            bool lastServiceIsA = false;
            for (int i = 0; i < PointStarts.Count - 1; i++)
            {
                Excel.Range shotRange = shotSheet.get_Range(ShotLeft + PointStarts[i], ShotRight + (PointStarts[i + 1] - 1));

                int row = StartRow + i;
                var left  = "AS" + row;
                var right = DataSheet.IntToCol(rallyLastColumn) + row;
                Excel.Range rallyRange = rallySheet.get_Range(left, right);

                bool serverIsA = shotSheet.get_Range(Service + PointStarts[i]).Value2 != null;

                try
                {
                    MakeLine(shotRange, rallyRange, serverIsA, PointStarts[i + 1] - PointStarts[i]);
                }
                catch(Exception e)
                {
                    MessageBox.Show("Error at " + PointStarts[i]+ " to " + (PointStarts[i + 1]-1) + "\n" + e.Message);
                    pgDiag.Close();
                    throw;
                }

                if (i > 0 && lastServiceIsA != serverIsA)
                {
                    Excel.Range range = rallySheet.get_Range("A" + row, DataSheet.IntToCol(rallyLastColumn) + row);
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


        enum Kinds
        {
            OtherHitAng,        //相手のショットに対する打点座標を用いる角度
            OtherRecAng,        //相手の被打点座標を用いる角度
            MeHitterAng,        //自分のショットに対する打点座標を用いる角度
            OtherHitBigArea,    //相手打点座標による大きい攻撃面積
            OtherHitSmallArea,  //相手被打点座標による大きい攻撃面積
            OtherRecBigArea,    //相手打点座標による小さい攻撃面積
            OtherRecSmallArea,  //相手被打点座標による小さい攻撃面積
            SurpriseHitAng,     //相手の打点座標を用いる角度のうち,逆を突いた分
            SurpriseRecAng,     //相手の被打点座標を用いる角度のうち,逆を突いた分
            SurpriseMeAng,      //自分のショットに対する打点座標を用いる角度のうち,逆を突いた分
            RunningDistanceHitToHit,    //前回の打点から今回の打点までの走った距離
            RunningDistanceHitToRec     //被打点から打点までの距離
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

            //1行の時(ダブルフォルト)の時は何も書かずに飛ばす
            if (rallys.Count <= 1)
                return;

            Dictionary<Kinds, DataInfo[]> Datas = new Dictionary<Kinds, DataInfo[]>();

            int left = DataSheet.ColToInt("A");
            foreach(Kinds key in Enum.GetValues(typeof(Kinds)))
            {
                // 0: サーバー, 1:レシーバー
                Datas.Add(key, new DataInfo[2]);

                for(int i=0; i<2; i++)
                {
                    int offset = serverIsA ? i : (i + 1) % 2;

                    int len = key == Kinds.SurpriseHitAng || key == Kinds.SurpriseMeAng || key == Kinds.SurpriseRecAng ? 
                        DataInfo.NeedCols-1+2 : DataInfo.NeedCols-1;

                    var range = rallyRange.get_Range(DataSheet.IntToCol(left + offset) + 1,
                                                     DataSheet.IntToCol(left + offset + len) + 1);
                    Datas[key][i] = new DataInfo(range);
                }

                left += Datas[key][0].Cols;
            }

            for (int i = 1; i < rallys.Count; i++)
            {
                //サーバーの角度(面積)かレシーバのかを決めるインデックス
                int index_odd_0 = (i + 1) % 2;   //奇数番が0になるインデックス
                int index_even_0 = i % 2;       //偶数番が0になるインデックス

                var current = rallys[i];
                var prev1   = rallys[i - 1];

                //最後の打者位置がnullの場合は, バウンド位置を使うものだけ判断する
                if( current.HitterPos == null)
                {
                    if( i != rallys.Count-1)
                        throw new Exception("入力されていない打選手,被打選手の座標があります");

                    
                    if (current.BoundPos != null)
                    {
                        //相手被打点座標による小さい攻撃面積
                        Datas[Kinds.OtherRecSmallArea][index_odd_0].Update(GetSmallArea(current.BoundPos.v, prev1.HitterPos.v, prev1.RecieverPos.v));

                        //相手打点座標による小さい攻撃面積
                        if (i > 1)
                            Datas[Kinds.OtherHitSmallArea][index_odd_0].Update(GetSmallArea(current.BoundPos.v, prev1.HitterPos.v, rallys[i - 2].HitterPos.v));
                    }
                    break;
                }

                //前の打者 -> 今の打者  へのベクトル
                var vecHtoH_c  = current.HitterPos.v - prev1.HitterPos.v;

                //前の打者 -> 前の被打者へのベクトル
                var vecHtoR_p1 = prev1.RecieverPos.v - prev1.HitterPos.v;

                //相手の被打点座標を用いる角度
                double angForOtherR = Math.Abs(System.Windows.Vector.AngleBetween(vecHtoH_c, prev1.HitterToReciever));
                Datas[Kinds.OtherRecAng][index_odd_0].Update(angForOtherR);

                //相手被打点座標による大きい攻撃面積
                Datas[Kinds.OtherRecBigArea][index_odd_0].Update(GetArea(current.HitterPos.v, prev1.HitterPos.v, prev1.RecieverPos.v));

                if (current.BoundPos != null){
                    //相手被打点座標による小さい攻撃面積
                    Datas[Kinds.OtherRecSmallArea][index_odd_0].Update(GetSmallArea(current.BoundPos.v, prev1.HitterPos.v, prev1.RecieverPos.v));
                }

                //被打点 -> 打点の距離
                Datas[Kinds.RunningDistanceHitToRec][index_even_0].Update((current.HitterPos.v - prev1.RecieverPos.v).Length);

                //以下2ラリー前の座標を用いる情報
                if (i < 2)
                    continue;

                //相手のショットに対する,打点座標を用いる角度
                var prev2 = rallys[i - 2];
                var vecHtoH_p1 = prev1.HitterPos.v - prev2.HitterPos.v;
                double angForOtherH = Math.Abs(System.Windows.Vector.AngleBetween(vecHtoH_c, -vecHtoH_p1));
                Datas[Kinds.OtherHitAng][index_odd_0].Update(angForOtherH);

                //逆を突いたかどうか
                bool takeSurprise = (prev1.RecieverPos.v.X - prev2.HitterPos.v.X) * (current.HitterPos.v.X - prev1.RecieverPos.v.X) < 0;

                //逆を突く形になった場合
                //別の項目にも保存する.
                if (takeSurprise)
                {
                    Datas[Kinds.SurpriseHitAng][index_odd_0].Update(angForOtherH);
                    Datas[Kinds.SurpriseRecAng][index_odd_0].Update(angForOtherR);
                }

                // 相手打点座標による大きい攻撃面積
                Datas[Kinds.OtherHitBigArea][index_odd_0].Update(GetArea(current.HitterPos.v, prev1.HitterPos.v, prev2.HitterPos.v));

                if(current.BoundPos != null){
                    //相手打点座標による小さい攻撃面積
                    Datas[Kinds.OtherHitSmallArea][index_odd_0].Update(GetSmallArea(current.BoundPos.v, prev1.HitterPos.v, prev2.HitterPos.v));
                }

                //打点 -> 打点の距離
                Datas[Kinds.RunningDistanceHitToHit][index_even_0].Update((current.HitterPos.v - prev2.HitterPos.v).Length);

                //以下3ラリー前の座標を用いる場合
                if (i < 3)
                    continue;

                var prev3 = rallys[i - 3];
                var vecHtoH_p2 = prev2.HitterPos.v - prev3.HitterPos.v;
                double angForMeH = Math.Abs(System.Windows.Vector.AngleBetween(vecHtoH_p2, vecHtoR_p1));
                Datas[Kinds.MeHitterAng][index_odd_0].Update(angForMeH);

                //逆を突く形になった場合
                //別の項目にも保存する.
                if (takeSurprise)
                {
                    Datas[Kinds.SurpriseMeAng][index_odd_0].Update(angForMeH);
                }
            }

            //ダブルフォルトじゃない場合
            //サーブの角度を追加
            if( rallys.Count > 1 )
            {
                if(rallys[1].BoundPos == null)
                {
                    throw new Exception("サーブのバウンド位置が入力されていません");
                }
                Console.WriteLine(rallys[1].BoundPos.v);
                rallyRange.get_Range(DataSheet.IntToCol(left) + 1).Value2 = Math.Abs(System.Windows.Vector.AngleBetween(
                    rallys[1].BoundPos.v - rallys[0].HitterPos.v, rallys[0].HitterToReciever));
            }

            //レシーバの角度
            if( Datas[Kinds.OtherHitAng][1].datas.Count > 0)
                rallyRange.get_Range(DataSheet.IntToCol(left+1) + 1).Value2 = Datas[Kinds.OtherHitAng][1].datas[0];

            if(Datas[Kinds.OtherRecAng][1].datas.Count > 0)
                rallyRange.get_Range(DataSheet.IntToCol(left + 2) + 1).Value2 = Datas[Kinds.OtherRecAng][1].datas[0];

            foreach( var data in Datas.Values) {
                for(int i=0; i<2; i++) {
                    data[i].Write();
                }
            }
        }

        class RallyInfo
        {
            public Vec2 BoundPos;       //前の選手が打ったショットのバウンド位置
            public Vec2 WinnerPos;
            public Vec2 MissPos;
            public Vec2 HitterPos;      //打選手の位置
            public Vec2 RecieverPos;    //被打選手の位置

            public System.Windows.Vector HitterToReciever
            {
                get
                {
                    return RecieverPos.v - HitterPos.v;
                }
            }

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
            public List<double> datas { get; private set; }

            public const int NeedCols = 12;

            public int Cols { get; private set; }
            public Excel.Range range { get; private set; }
            public DataInfo(Excel.Range r, bool cntCols = false)
            {
                datas = new List<double>();

                Cols = r.Columns.Count;
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

            public void Write()
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
                
                if (Cols > NeedCols)
                    range.get_Range("M1").Value2 = datas.Count;
            }
        }
    }
}
