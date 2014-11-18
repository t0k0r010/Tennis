using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Windows;

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

        int Width_p { get; set; }       //画面上のコートのサイズ [pixel]
        int Height_p { get; set; }
        Point Center_p { get; set; }    //コートの中心位置 [pixel]

        public Dictionary<Color, List<Point>> Positions_p{get; private set;}
        Panel Panel;

        Label infomation;

        public Court(Panel panel)
        {
            Positions_p = new Dictionary<Color, List<Point>>();
            Positions_p.Add(BoundColor, new List<Point>());
            Positions_p.Add(HitterColor, new List<Point>());
            Positions_p.Add(RecieverColor, new List<Point>());

            this.Panel = panel;
            panel.Paint += new System.Windows.Forms.PaintEventHandler(Draw);
            SetCourtSize(panel);

            infomation = new Label();
            infomation.Parent = Panel;
        }

        
        public ShotInfo GetLastShotInfo()
        {
            //ラリーが始まっていない場合は,情報がとれない
            if (!(Positions_p.ContainsKey(HitterColor) && Positions_p.ContainsKey(RecieverColor)))
                return null;

            if( Positions_p[HitterColor].Count == 0 || Positions_p[RecieverColor].Count == 0)
                return null;

            ShotInfo s = new ShotInfo();
            var hitter   = Positions_p[HitterColor];
            var reciever = Positions_p[RecieverColor];
            s.Hitter   = ToRealUnit( hitter[hitter.Count - 1] );
            s.Reciever = ToRealUnit(reciever[reciever.Count - 1]);

            return s;
        }

        //新しくクリックした位置を追加
        public void AddPosition(Point _p, Color c)
        {
            if( !Positions_p.ContainsKey(c))
                Positions_p.Add(c, new List<Point>());

            Positions_p[c].Add(_p);
            
            Panel.Invalidate();
        }

        //クリックした位置の削除
        public void ClearPosition()
        {
            foreach(var points in Positions_p.Values)
                points.Clear();

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
        public float PixelToMeter(int pixel)
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
                Pen pen = new Pen(p.Key);
                SolidBrush brush = new SolidBrush(p.Key);
                foreach(var point in p.Value)
                    g.FillEllipse(brush, point.X - markerSize / 2, point.Y - markerSize / 2, markerSize, markerSize);
            }

            var hitMark = Positions_p[HitterColor];

            if (hitMark.Count > 1)            
                g.DrawLines(new Pen(HitterColor), hitMark.ToArray());

            if (hitMark.Count > 2)
            {
                Font f = new Font("Arial", 12);
                SolidBrush brush = new SolidBrush(Color.Black);
                for (int i = 1; i < hitMark.Count - 1; i++)
                {
                    var a = hitMark[i - 1];
                    var b = hitMark[i];
                    var c = hitMark[i + 1];

                    float deg = GetDegree(new PointF(a.X - b.X, a.Y - b.Y), new PointF(c.X - b.X, c.Y - b.Y));

                    g.DrawString(deg.ToString() + "°", f, brush, b);
                }
            }
            
        }

        float GetDegree(PointF vec1, PointF vec2)
        {
            double dot = vec1.X * vec2.X + vec1.Y * vec2.Y;
            double absA = Math.Sqrt(vec1.X * vec1.X + vec1.Y * vec1.Y);
            double absB = Math.Sqrt(vec2.X * vec2.X + vec2.Y * vec2.Y);

            double rad = Math.Acos(dot / absA / absB);

            return (float)(rad * 180 / Math.PI);
            
        }
        void DrawMarks(Graphics g, List<Point> marks, SolidBrush brush)
        {
            int markerSize = Width_p / 30;
            foreach (var point in marks)
                g.FillEllipse(brush, point.X - markerSize / 2, point.Y - markerSize / 2, markerSize, markerSize);
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

        public class ServeInfo
        {
            public PointF HitterPos, RecieverPos;
            public PointF Direction;
        }

        
    }
}