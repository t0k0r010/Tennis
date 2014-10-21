using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
//using System.Math;
namespace Tennis
{
    class Court
    {
        const float Width_m = 10.97f;    //現実のコートのサイズ[m]
        const float Height_m = 23.77f;
        const float SingleCourtWidth_m = 8.23f;   //シングルコートの幅
        const float ServiceLineHeight_m = 6.4f;  //ネットからサービスラインまでの長さ

        int Width_p { get; set; }       //画面上のコートのサイズ [pixel]
        int Height_p { get; set; }
        Point Center_p { get; set; }    //コートの中心位置 [pixel]

        List<Point> BoundPositions_p = new List<Point>();

        Panel Panel;

        Label infomation;

        public bool IsPlayerA { get;  private set;  }

        public bool CheckBoundPosition { get; private set; }

        public void SetCheckBoundPosition(bool bound)
        {
            if (CheckBoundPosition == bound)
                return;

            CheckBoundPosition = bound;
            ExcelWriter.Instance.MoveToFistLine();
        }

        public Court(Panel panel)
        {
            this.Panel = panel;
            panel.Paint += new System.Windows.Forms.PaintEventHandler(Draw);
            panel.MouseClick += MouseClick;
            SetCourtSize(panel);

            infomation = new Label();
            infomation.Parent = Panel;
        }

        //コートをクリック => バウンド位置を設定
        void MouseClick(object sender, MouseEventArgs e)
        {
            if (!Form1.IsStarted)
                return;

            if (CheckBoundPosition)
                ClickedBoundPosition(sender, e);
            else
                ClickedPlayerPosition(sender, e);
            Panel.Invalidate(); //再描画命令
        }

        void ChangeIndicator()
        {
            infomation.Text = IsPlayerA ? "PlayerA" : "PlayerB";
        }

        //バウンド地点を書き出す
        void ClickedBoundPosition(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                BoundPositions_p.Add(new Point(e.X, e.Y));
                PointF point_m = ToRealUnit(e.X, e.Y);

                ExcelWriter.Instance.SetBoundPosition("no time", point_m.X, point_m.Y);
                ExcelWriter.Instance.MoveToNextLine();  //次のラインへ
            }
            else
            {
                ExcelWriter.Instance.WriteLine();
                BoundPositions_p.Clear();
            }
        }

        void ClickedPlayerPosition(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                BoundPositions_p.Add(new Point(e.X, e.Y));
                PointF point_m = ToRealUnit(e.X, e.Y);

                if(IsPlayerA)
                {
                    ExcelWriter.Instance.SetRecieverPosition("no time", point_m.X, point_m.Y);
                    ExcelWriter.Instance.MoveToNextLine();  //次のラインへ
                }
                else
                {
                    ExcelWriter.Instance.SetHitterPosition("no time", point_m.X, point_m.Y);
                }

                IsPlayerA = !IsPlayerA;
            }
            else
            {
                ExcelWriter.Instance.WriteLine();
                BoundPositions_p.Clear();
            }
        }

        void Draw(Object panel, PaintEventArgs e)
        {
            SetCourtSize((Panel)panel);    //サイズ変わった時の為に毎回セットする.
            DrawCourt(e.Graphics);
            DrawBoundMarks(e.Graphics);
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
            int SupportLineHeight_p = MeterToPixel(1); //1mをpixelに変換

            g.DrawLine(pen,
                Center_p.X - Width_p / 2 - SupportLineWidth_p, Center_p.Y + Height_p / 2 - SupportLineHeight_p,
                 Center_p.X + Width_p / 2 + SupportLineWidth_p, Center_p.Y + Height_p / 2 - SupportLineHeight_p);

            g.DrawLine(pen,
                Center_p.X - Width_p / 2 - SupportLineWidth_p, Center_p.Y + Height_p / 2 + SupportLineHeight_p,
                Center_p.X + Width_p / 2 + SupportLineWidth_p, Center_p.Y + Height_p / 2 + SupportLineHeight_p);

            g.DrawLine(pen,
                Center_p.X - Width_p / 2 - SupportLineWidth_p, Center_p.Y - Height_p / 2 + SupportLineHeight_p,
                Center_p.X + Width_p / 2 + SupportLineWidth_p, Center_p.Y - Height_p / 2 + SupportLineHeight_p);

            g.DrawLine(pen,
                Center_p.X - Width_p / 2 - SupportLineWidth_p, Center_p.Y - Height_p / 2 - SupportLineHeight_p,
                Center_p.X + Width_p / 2 + SupportLineWidth_p, Center_p.Y - Height_p / 2 - SupportLineHeight_p);


        }

        //バウンド跡を描画
        void DrawBoundMarks(Graphics g)
        {
            Pen pen = new Pen(Color.Red);
            int markerSize = Width_p / 10;
            foreach (Point p in BoundPositions_p)
            {
                g.DrawEllipse(pen, p.X - markerSize / 2, p.Y - markerSize / 2, markerSize, markerSize);
            }
        }
        

        /*
private void ClickCourt(object sender, MouseEventArgs e)
{
    //左クリックで新しくバウンド位置を追加
    if (e.Button == MouseButtons.Left)
    {
        cursor = new Point(e.X, e.Y);
        cursorPosition.Text = "(" + e.X + "," + e.Y + ")";

        cursorText.Text = e.X + "," + e.Y;

        BoundPositions.Add(cursor);
    }
        //右クリックでバウンドの軌跡を消去
    else if (e.Button == MouseButtons.Right)
    {
        Output();
        BoundPositions.Clear();
    }
    CourtPannel.Invalidate(); //再描画しろっていう命令
}*/
    }
}