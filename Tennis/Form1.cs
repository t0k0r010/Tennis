using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Collections;
using Microsoft.Office.Interop;
namespace Tennis
{
    public partial class Form1 : Form
    {
        Point cursor = new Point();
        MediaPlayerForm mediaPlayer = null;
        double mediaPos = 0;

        ArrayList BoundPositions = new ArrayList();

        Microsoft.Office.Interop.Excel.Application exelApp;
        Microsoft.Office.Interop.Excel.Workbook wb;

        class CourtSetting
        {

            public float scale;
            public Point origin;

        }
        CourtSetting courtSet = new CourtSetting();
        int rary_num =1;
        void setCourtsetting( Panel panel)
        {
            courtSet.scale = 10.97f / panel.Width/2.0f; //[m/pixel]
            //クリックした点を円で描画
            courtSet.origin = new Point(panel.Width / 2, panel.Height / 2);
        }
        public Form1()
        {
            //初期設定を書く
            InitializeComponent();

            this.Court.Paint += new System.Windows.Forms.PaintEventHandler(this.DrawCourt);
            this.Court.MouseClick += new System.Windows.Forms.MouseEventHandler(this.ClickCourt);

            //ファイルダイアログを開く
            DialogResult res = openFileDialog1.ShowDialog();

            //開くボタンを押したとき
            if (res == System.Windows.Forms.DialogResult.OK)
            {
                string fileName  = openFileDialog1.FileName;
                exelApp = new Microsoft.Office.Interop.Excel.Application();
                exelApp.Visible = true;
                wb = exelApp.Workbooks.Open(Filename: fileName);
                try
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[1]).Select();
                }
                catch (Exception ex)
                {
                    wb.Close(false);
                    exelApp.Quit();
                    MessageBox.Show("開けませんでした");
                    return;
                }
            }
            else
            {

            }

            //Observeという関数を別に実行する.
            Thread thread = new Thread(Observe);
            thread.Start();
        }

        
        delegate void SetMedia();

        private void Observe()
        {
            while (!this.IsDisposed)
            {
                if (mediaPlayer == null || mediaPlayer.axWindowsMediaPlayer1.IsDisposed)
                    continue;

                mediaPos = mediaPlayer.axWindowsMediaPlayer1.Ctlcontrols.currentPosition;
                this.Invoke( new SetMedia(SetMediaPos));
            }
        }

        private void SetMediaPos()
        {
            moviePos.Text = mediaPos.ToString();
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        void Output()
        {
            Microsoft.Office.Interop.Excel.Range CellRange;

            for(int i=0; i<BoundPositions.Count; i++)
            {
                Point pos = (Point)BoundPositions[i];
                Point p = new Point(pos.X - courtSet.origin.X, pos.Y - courtSet.origin.Y);
                float x = (float)Math.Round(courtSet.scale * p.X, 2, MidpointRounding.AwayFromZero);
                float y = (float)Math.Round(courtSet.scale * p.Y, 2, MidpointRounding.AwayFromZero);

                CellRange = exelApp.Cells[i+1, rary_num] as Microsoft.Office.Interop.Excel.Range;
                CellRange.Value2 = x;

                CellRange = exelApp.Cells[i + 1, rary_num+1] as Microsoft.Office.Interop.Excel.Range;
                CellRange.Value2 = y;
            }
            rary_num += 2;
            MessageBox.Show("出力完了");
        }

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
            Court.Invalidate(); //再描画しろっていう命令
        }

        //コートを描画する
        private void DrawCourt(object sender, PaintEventArgs e)
        {
            Panel panel = (Panel)sender;
            setCourtsetting(panel);
            Graphics g = e.Graphics;
            Pen pen = new Pen(Color.Black);
         //   g.DrawRectangle(pen, 0, 0, panel.Width/2, panel.Height/2);

            int courtWidth = panel.Width /2;
            int courtHeight = (int)(courtWidth / 10.97f * 23.77f);

            //実際の線とプログラム上の線との単位変換
            float scale = 10.97f / courtWidth; //[m/pixel]
            // 100 [pixel] * scale => xxx[m] に変換
            // 100 [m] / scale     => yyy[pixel] に変換

            //コートの外枠を描画
            Point courtUpLeft = new Point( panel.Width / 2 - courtWidth / 2, panel.Height / 2 - courtHeight / 2);
            g.DrawRectangle(pen, courtUpLeft.X, courtUpLeft.Y, courtWidth, courtHeight);

            //ネットの描画
            g.DrawLine(pen, new Point( panel.Width / 2 - courtWidth / 2,
                panel.Height / 2), new Point(panel.Width / 2 + courtWidth / 2, panel.Height / 2));

            //シングルラインの描画
            int singleBaseLineWidth = (int)(courtWidth * 8.23f / 10.97f);
            int singleBaleLineX = courtUpLeft.X + (courtWidth - singleBaseLineWidth) / 2;
            g.DrawLine(pen, new Point(singleBaleLineX, courtUpLeft.Y),
                new Point(singleBaleLineX, courtUpLeft.Y + courtHeight));
            g.DrawLine(pen, new Point(panel.Width / 2 + singleBaseLineWidth / 2, courtUpLeft.Y)
                , new Point(panel.Width / 2 + singleBaseLineWidth / 2, courtUpLeft.Y+courtHeight));

            //サービスラインの描画
            int serviceLineHeight = (int)(courtHeight * 12.8 / 23.77);
            g.DrawLine(pen, new Point(singleBaleLineX, panel.Height / 2 - serviceLineHeight / 2),
                new Point(singleBaleLineX + singleBaseLineWidth, panel.Height / 2 - serviceLineHeight / 2));
            g.DrawLine(pen, new Point(singleBaleLineX, panel.Height / 2 + serviceLineHeight / 2),
                new Point(singleBaleLineX + singleBaseLineWidth, panel.Height / 2 + serviceLineHeight / 2));

            //中心の縦ラインを描画
            g.DrawLine(pen, new Point(panel.Width / 2, panel.Height / 2 - serviceLineHeight / 2), new Point(panel.Width / 2, panel.Height / 2 + serviceLineHeight / 2));

            //補助線を描画

           // g.DrawEllipse(pen, cursor.X-5 , cursor.Y-5, 10, 10);

            //クリックした点を円で描画
            Point center = new Point(panel.Width / 2, panel.Height / 2);
            foreach (Point pos in BoundPositions)
            {
                g.DrawEllipse(pen, pos.X - 5, pos.Y - 5, 10, 10);
                Point p = new Point(pos.X - center.X, pos.Y-center.Y) ;
                float x = (float)Math.Round(scale * p.X, 2, MidpointRounding.AwayFromZero);
                float y = (float)Math.Round(scale * p.Y, 2, MidpointRounding.AwayFromZero);
                string text = x + ", " + y;
                float width = 10 * text.Length;

                RectangleF rect = new RectangleF(pos.X - width / 2, pos.Y + 5, width, 30);

                StringFormat format = new StringFormat();
                format.Alignment = StringAlignment.Center;
                format.LineAlignment = StringAlignment.Center;
                g.DrawString(text, new Font("Arial", 8), Brushes.Black, rect, format);
            }
            /*
            if (cursor.X < courtUpLeft.X)
            {
                cursorPosition.Text = "Out";
            } else {
                cursorPosition.Text = "In";
            }*/
        }

        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //ファイルダイアログを開く
            DialogResult res = openFileDialog1.ShowDialog();

            //開くボタンを押したとき
            if (res == System.Windows.Forms.DialogResult.OK)
            {
                //MediaPlayerのウィンドウを新しく開く
                mediaPlayer = new MediaPlayerForm();
                mediaPlayer.axWindowsMediaPlayer1.URL = openFileDialog1.FileName;
                mediaPlayer.Show();
            }
        }

        private void OpenExelToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
