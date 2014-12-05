using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Linq;

namespace Tennis
{
    class Court
    {
        public static readonly Color BoundColor = Color.Green;
        public static readonly Color HitterColor = Color.Red;
        public static readonly Color RecieverColor = Color.Blue;
        public const float Width_m = 10.97f;    //現実のコートのサイズ[m]
        public const float Height_m = 23.77f;
        public const float SingleCourtWidth_m = 8.23f;   //シングルコートの幅
        public const float ServiceLineHeight_m = 6.4f;  //ネットからサービスラインまでの長さ

        public static Court Instance = null;

        public enum Players
        {
            PlayerA,
            PlayerB
        }

        //コートの上側にいるプレイヤー
        public Players UpperPlayer{ get; set;}
        public Players LowerPlayer { get {
            return UpperPlayer == Players.PlayerA ? Players.PlayerB : Players.PlayerA; 
        } 
        }
        //次にクリックする位置はどの座標のことを指すかの候補
        public enum Surveyed
        {
            BoundPos,       //バウンド位置
            HitterPos,      //打者の位置
            RecieverPos,    //被打者の位置
            None            //何も見ていない
        }

        //バウンド位置を見るか、プレイヤー位置を見るか
        bool CheckingBoundPos = true;

        //最後にクリックしたものが何かどうか
        public Surveyed LastSurveyed
        {
            get
            {
                if (CheckingBoundPos)
                    return Surveyed.BoundPos;

                if (Positions_p[Surveyed.HitterPos].Count == Positions_p[Surveyed.RecieverPos].Count)
                    return Positions_p[Surveyed.HitterPos].Count == 0 ? Surveyed.None : Surveyed.RecieverPos;
                else
                    return Surveyed.HitterPos;

            }
        }

        //次のクリック対象
        public Surveyed NextSurveying
        {
            get
            {
                if (CheckingBoundPos)
                    return Surveyed.BoundPos;

                if (LastSurveyed == Surveyed.RecieverPos || LastSurveyed == Surveyed.None)
                    return Surveyed.HitterPos;
                else
                    return Surveyed.RecieverPos;
            }
        }

        //ラリー回数
        public int RallyNum { get; private set; }

        int Width_p { get; set; }       //画面上のコートのサイズ [pixel]
        int Height_p { get; set; }
        Point Center_p { get; set; }    //コートの中心位置 [pixel]

        public Dictionary<Surveyed, List<Point>> Positions_p { get; private set; }
        public Dictionary<Surveyed, Color> MarkColors;

        //1ポイントにおけるショットの方向を表す.
        //偶数番目 : サーバのショット
        //奇数番目 : レシーバのショット
        public List<System.Windows.Vector> ShotDirectionsInRally{get; private set;}

        //1ポイントにおける相手のショットに対する角度を表す
        //偶数番目 : レシーバからみた角度
        //奇数番目 : サーバからみた角度
        public List<double> ShotAngleForOther { get; private set; }

        //1ポイントにおける相手のショットに対する角度を表す
        //偶数番目 : サーバからみた角度
        //奇数番目 : レシーバからみた角度
        public List<double> ShotAngleForMe { get; private set; }

        Panel Panel;

        Label infomation;

        public Court(Panel panel)
        {
            Positions_p = new Dictionary<Surveyed, List<Point>>();
            Positions_p.Add(Surveyed.BoundPos, new List<Point>());
            Positions_p.Add(Surveyed.HitterPos, new List<Point>());
            Positions_p.Add(Surveyed.RecieverPos, new List<Point>());

            MarkColors = new Dictionary<Surveyed, Color>();
            MarkColors.Add(Surveyed.BoundPos, BoundColor);
            MarkColors.Add(Surveyed.HitterPos, HitterColor);
            MarkColors.Add(Surveyed.RecieverPos, RecieverColor);

            ShotDirectionsInRally = new List<System.Windows.Vector>();
            ShotAngleForMe = new List<double>();
            ShotAngleForOther = new List<double>();

            this.Panel = panel;
            panel.Paint += new System.Windows.Forms.PaintEventHandler(Draw);
            SetCourtSize(panel);

            infomation = new Label();
            infomation.Parent = Panel;

            Instance = this;
        }

        public ShotInfo GetLastShotInfo()
        {
            //ラリーが始まっていない場合は,情報がとれない
            if (!(Positions_p.ContainsKey(Surveyed.HitterPos) && Positions_p.ContainsKey(Surveyed.RecieverPos)))
                return null;

            if( Positions_p[Surveyed.HitterPos].Count == 0 || Positions_p[Surveyed.RecieverPos].Count == 0)
                return null;

            ShotInfo s = new ShotInfo();
            var hitter   = Positions_p[Surveyed.HitterPos];
            var reciever = Positions_p[Surveyed.RecieverPos];
            s.Hitter   = ToRealUnit( hitter[hitter.Count - 1] );
            s.Reciever = ToRealUnit(reciever[reciever.Count - 1]);

            return s;
        }

        //新しくクリックした位置を追加
        //ここではSurveyは変化しない => この後データシートから見た時にずれる為
        //リストの最新の値とSurveyは同期するようにする.
        public void AddPosition(Point p)
        {
            var s = NextSurveying;
            if( !Positions_p.ContainsKey(s))
                Positions_p.Add(s, new List<Point>());

            Positions_p[s].Add(p);

            //打点の場合
            if (s == Surveyed.HitterPos && Positions_p[Surveyed.HitterPos].Count > 1)
            {
                var hitter = Positions_p[Surveyed.HitterPos];
                System.Windows.Vector direction = ToVector(hitter[hitter.Count - 2]) - ToVector(hitter[hitter.Count - 1]);

                ShotDirectionsInRally.Add(direction);

                var num = ShotDirectionsInRally.Count;
                //方向ベクトルが2つ以上あると, 相手のショットに対する角度が出せる
                if(num > 1)
                {
                    // 最新2つの角度を求める(コピー渡しされる)
                    var vec1 = -ShotDirectionsInRally[num - 2];
                    var vec2 = ShotDirectionsInRally[num - 1];

                    //ショットの角度を計算
                    var deg = Math.Abs(ToRoundDown(System.Windows.Vector.AngleBetween(vec1, vec2), 2));
                    ShotAngleForOther.Add(deg);
                }

                //3本以上あると, 自分のショットに対する角度が出せる
                if(num > 2)
                {
                    var vec1 = ShotDirectionsInRally[num - 3];
                    var vec2 = ShotDirectionsInRally[num - 1];

                    //ショットの角度を計算
                    var deg = Math.Abs(ToRoundDown(System.Windows.Vector.AngleBetween(vec1, vec2), 2));
                    ShotAngleForMe.Add(deg);
                }
            }

            //被打者orバウンド位置を見ていた場合 => ラリーが増える
            if( s == Surveyed.RecieverPos || s == Surveyed.BoundPos)
                RallyNum++;

            Panel.Invalidate();
        }

        public static System.Windows.Vector ToVector(Point p)
        {
            return new System.Windows.Vector(p.X, p.Y);
        }

        public static System.Windows.Vector ToVector(PointF p)
        {
            return new System.Windows.Vector(p.X, p.Y);
        } 

        public void ChangeSurveyed(Surveyed s)
        {
            bool before = CheckingBoundPos;
            CheckingBoundPos = (s == Surveyed.BoundPos);

            if( before != CheckingBoundPos)
                ClearPosition();
        }

        //クリックした位置の削除
        public void ClearPosition()
        {
            foreach(var points in Positions_p.Values)
                points.Clear();

            ShotDirectionsInRally.Clear();
            ShotAngleForOther.Clear();
            ShotAngleForMe.Clear();
            RallyNum = 0;
            Panel.Invalidate();
        }

        //コートのサイズを再計算
        void SetCourtSize(Panel panel)
        {
            //パネルの半分の大きさにする
            if( (panel.Width / Width_m) < (panel.Height / Height_m) )
            {
                Width_p = (int) (panel.Width * 0.7f);
                Height_p = (int)Math.Round(Width_p / Width_m * Height_m);
            } else
            {
                Height_p = (int)(panel.Height*0.7f);
                Width_p = (int)Math.Round(Height_p / Height_m * Width_m);
            }

            //中心を原点とする.
            Center_p = new Point(panel.Width / 2, panel.Height / 2);
        }

        // [m] -> [pixel] 変換
        public int MeterToPixel(float meter)
        {
            //整数にするので四捨五入
            return (int)Math.Round(meter * Width_p / Width_m);
        }

        // [pixel] -> [m] 変換
        public float PixelToMeter(float pixel)
        {
            return pixel * Width_m / Width_p;
        }

        //コートの中心を原点とした,実世界の[m]単位に変換
        public PointF ToRealUnit(Point p)
        {
            float x = (float)Math.Round( PixelToMeter(p.X - Center_p.X), 2, MidpointRounding.AwayFromZero);
            float y = (float)Math.Round( PixelToMeter(p.Y - Center_p.Y), 2, MidpointRounding.AwayFromZero);

            return new PointF(x, y);
        }

        //コートの中心を原点とした,実世界の[m]単位に変換
        public PointF ToRealUnit(int p_x, int p_y)
        {
            float x = (float)Math.Round(PixelToMeter(p_x - Center_p.X), 2, MidpointRounding.AwayFromZero);
            float y = (float)Math.Round(PixelToMeter(p_y - Center_p.Y), 2, MidpointRounding.AwayFromZero);

            return new PointF(x, y);
        }

        //描画
        void Draw(Object panel, PaintEventArgs e)
        {
            SetCourtSize((Panel)panel);    //サイズ変わった時の為に毎回セットする.
            DrawCourt(e.Graphics);
            DrawMarks(e.Graphics);
        }

        //コートの描画
        void DrawCourt(Graphics g)
        {
            //線の色やサイズを決める.
            Pen pen = new Pen(Color.Black);

            //コートの左上座標を計算
            Point upperLeft = new Point(Center_p.X - Width_p / 2, Center_p.Y - Height_p / 2);

            //外枠を描画
            g.DrawRectangle(pen, upperLeft.X, upperLeft.Y, Width_p, Height_p);

            //シングルラインの描画
            int singleCourtWidth_p = MeterToPixel(SingleCourtWidth_m);
            g.DrawLine(pen,                                                 //左のライン
                Center_p.X - singleCourtWidth_p / 2, upperLeft.Y,
                Center_p.X - singleCourtWidth_p / 2, upperLeft.Y + Height_p);
            g.DrawLine(pen,                                                 //右のライン
                Center_p.X + singleCourtWidth_p / 2, upperLeft.Y,
                Center_p.X + singleCourtWidth_p / 2, upperLeft.Y + Height_p);

            //サービスラインの描画
            int serviceLineHeight_p = MeterToPixel(ServiceLineHeight_m);
            g.DrawLine(pen,
                Center_p.X - singleCourtWidth_p / 2, Center_p.Y - serviceLineHeight_p,    //上のライン
                Center_p.X + singleCourtWidth_p / 2, Center_p.Y - serviceLineHeight_p);
            g.DrawLine(pen,
                Center_p.X - singleCourtWidth_p / 2, Center_p.Y + serviceLineHeight_p,    //下のライン
                Center_p.X + singleCourtWidth_p / 2, Center_p.Y + serviceLineHeight_p);
            

            //センターマークの描画

            //サーブ補助線の描画
            int SupportLineHeight_p = MeterToPixel(1); //1mをpixelに変換

            
            g.DrawLine(pen,
                Center_p.X, Center_p.Y - Height_p / 2,
                Center_p.X, Center_p.Y - Height_p / 2 + SupportLineHeight_p / 5);
             
            
            g.DrawLine(pen,
                Center_p.X, Center_p.Y + Height_p / 2,
                Center_p.X, Center_p.Y + Height_p / 2 - SupportLineHeight_p / 5);

            //本来であればセンターマークは10㎝だが，見やすさ重視で20㎝に変換する
            

            //中心のラインの描画
            g.DrawLine(pen,
               Center_p.X, Center_p.Y - serviceLineHeight_p,
               Center_p.X, Center_p.Y + serviceLineHeight_p);

            //ネットの描画
            int SupportLineWidth_p = MeterToPixel(0.91f);
            g.DrawLine(pen,
                Center_p.X - Width_p / 2 - SupportLineWidth_p, Center_p.Y,
                Center_p.X + Width_p / 2 + SupportLineWidth_p, Center_p.Y);
            
            //破線に変更
            pen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;

            //サーブ補助線の描画
           // int SupportLineHeight_p = MeterToPixel(1); //1mをpixelに変換
            
            //コート下±１ｍの補助線上下
            g.DrawLine(pen,
                Center_p.X - Width_p / 2 - SupportLineWidth_p, Center_p.Y + Height_p / 2 - SupportLineHeight_p,
                 Center_p.X + Width_p / 2 + SupportLineWidth_p, Center_p.Y + Height_p / 2 - SupportLineHeight_p);

            g.DrawLine(pen,
                Center_p.X - Width_p / 2 - SupportLineWidth_p, Center_p.Y + Height_p / 2 + SupportLineHeight_p,
                Center_p.X + Width_p / 2 + SupportLineWidth_p, Center_p.Y + Height_p / 2 + SupportLineHeight_p);

            //コート下2mの補助線
            g.DrawLine(pen,
                Center_p.X - Width_p / 2 - SupportLineWidth_p, Center_p.Y + Height_p / 2 + SupportLineHeight_p*2,
                 Center_p.X + Width_p / 2 + SupportLineWidth_p, Center_p.Y + Height_p / 2 + SupportLineHeight_p*2);

            //コート上±１ｍの補助線
            g.DrawLine(pen,
                Center_p.X - Width_p / 2 - SupportLineWidth_p, Center_p.Y - Height_p / 2 + SupportLineHeight_p,
                Center_p.X + Width_p / 2 + SupportLineWidth_p, Center_p.Y - Height_p / 2 + SupportLineHeight_p);

            g.DrawLine(pen,
                Center_p.X - Width_p / 2 - SupportLineWidth_p, Center_p.Y - Height_p / 2 - SupportLineHeight_p,
                Center_p.X + Width_p / 2 + SupportLineWidth_p, Center_p.Y - Height_p / 2 - SupportLineHeight_p);

            //コート上2mの補助線
            g.DrawLine(pen,
               Center_p.X - Width_p / 2 - SupportLineWidth_p, Center_p.Y - Height_p / 2 - SupportLineHeight_p*2,
               Center_p.X + Width_p / 2 + SupportLineWidth_p, Center_p.Y - Height_p / 2 - SupportLineHeight_p*2);

            //コート4分の１補助線下左右
            g.DrawLine(pen,
                Center_p.X - Width_p / 4, Center_p.Y+ Height_p / 2 - SupportLineHeight_p / 2,
                Center_p.X - Width_p / 4, Center_p.Y+ Height_p / 2 + SupportLineHeight_p / 2);

            g.DrawLine(pen,
                Center_p.X + Width_p / 4, Center_p.Y + Height_p / 2 - SupportLineHeight_p / 2,
                Center_p.X + Width_p / 4, Center_p.Y + Height_p / 2 + SupportLineHeight_p / 2);

            //コート4分の１補助線上左右
            g.DrawLine(pen,
                Center_p.X - Width_p / 4, Center_p.Y - Height_p / 2 - SupportLineHeight_p / 2,
                Center_p.X - Width_p / 4, Center_p.Y - Height_p / 2 + SupportLineHeight_p / 2);

            g.DrawLine(pen,
                Center_p.X + Width_p / 4, Center_p.Y - Height_p / 2 - SupportLineHeight_p / 2,
                Center_p.X + Width_p / 4, Center_p.Y - Height_p / 2 + SupportLineHeight_p / 2);            
        }

        //バウンド跡を描画
        void DrawMarks(Graphics g)
        {
            int markerSize = Width_p / 30;


            foreach( var p in Positions_p)
            {
                Pen pen = new Pen( MarkColors[p.Key] );
                SolidBrush brush = new SolidBrush(MarkColors[p.Key]);
                Font f = new Font("Arial", 12);
                foreach (var point in p.Value.Select((v, i) => new { v, i }))
                {
                    g.FillEllipse(brush, point.v.X - markerSize / 2, point.v.Y - markerSize / 2, markerSize, markerSize);
                    g.DrawString(point.i.ToString(), f, brush, new PointF(point.v.X + markerSize/2, point.v.Y - markerSize/2));
                }
            }

            var hitMark = Positions_p[Surveyed.HitterPos];

            if (hitMark.Count > 1)
                g.DrawLines(new Pen(MarkColors[Surveyed.HitterPos]), hitMark.ToArray());


            if (hitMark.Count > 2)
            {
                Font f = new Font("Arial", 12);
                SolidBrush brush = new SolidBrush(Color.Black);

                //StringFormatを作成
                StringFormat sf = new StringFormat();
                //文字を真ん中に表示
                sf.Alignment = StringAlignment.Center;
                sf.LineAlignment = StringAlignment.Center;

                //角度を表示
                for (int i = 1; i < hitMark.Count - 1; i++)
                {
                    var vec1 = -ShotDirectionsInRally[i - 1];   //コピー渡しされる
                    var vec2 =  ShotDirectionsInRally[i];
                    //文字を入れる位置
                    vec1.Normalize();
                    vec2.Normalize();
                    var v = -2 * markerSize * (vec1 + vec2);
                    g.DrawString(ShotAngleForOther[i - 1].ToString() + "°", f, brush, new PointF(hitMark[i].X + (float)v.X, hitMark[i].Y + (float)v.Y), sf);
                }

                //横に面積を表示
                for (int i = 2; i < hitMark.Count; i++)
                {
                    var a = ToVector(hitMark[i-2]);
                    var b = ToVector(hitMark[i-1]);
                    var c = ToVector(hitMark[i]);

                    var p = (a + b + c) / 3;

                    var area = Math.Abs(System.Windows.Vector.CrossProduct(a - b, c - b) / 2);
                    area = ToRoundDown( PixelToMeter( PixelToMeter((float)area)), 2);
                    g.DrawString(area.ToString(), f, brush, new PointF(0, i * 15));
                }
            }
        }

        //一回のショットの情報
        public class ShotInfo
        {
            public const float BigNum = -10000;
            //打者,被打者,バウンド位置
            public PointF Hitter = new PointF(BigNum, BigNum);
            public PointF Reciever = new PointF(BigNum, BigNum);
            public PointF Bound = new PointF(BigNum, BigNum);

            public bool Available()
            {
                return (Hitter.X > BigNum + 1) && (Reciever.X > BigNum + 1);
            }
        }

        public static double ToRoundDown(double dValue, int iDigits)
        {
            double dCoef = System.Math.Pow(10, iDigits);

            return dValue > 0 ? System.Math.Floor(dValue * dCoef) / dCoef :
                                System.Math.Ceiling(dValue * dCoef) / dCoef;
        }

    }
}