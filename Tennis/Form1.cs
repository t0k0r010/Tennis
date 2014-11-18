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
        Court court;

        //両方とも開いているか
        public static bool IsStarted { 
            get 
            {
                return ExcelWriter.Available(); 
            } 
        }

        public Form1()
        {
            //初期設定を書く
            InitializeComponent();

            court       = new Court(this.CourtPannel);

            this.CourtPannel.MouseClick += ClickCourt;
            this.KeyDown += Form1_KeyDown;
            this.Resize  += Form1_Resize;
        }

        void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            
        }

        void ClickCourt(object sender, MouseEventArgs e)
        {
            if (!Form1.IsStarted)
                return;

            ExcelWriter writer = ExcelWriter.Instance;
            Point p = new Point(e.X, e.Y);
            if( e.Button == MouseButtons.Left )
            {
                Color c = writer.shotSheet.Surveying == ExcelWriter.ShotDataSheet.Surveyed.BoundPos ? 
                    Court.BoundColor : (writer.shotSheet.IsHitter ? Court.HitterColor : Court.RecieverColor);

                court.AddPosition(p, c);    //コートに新しい位置を追加
                PointF realPos = court.ToRealUnit(p);
                writer.shotSheet.SetPosition("", realPos);
            }
            else if(e.Button == MouseButtons.Right)
            {
                court.ClearPosition();                  //コートに書いている位置を削除
                writer.shotSheet.EndRally(); //エクセルにラインを書き込む
            }
            CourtPannel.Invalidate(); //再描画命令
        }

        void Form1_Resize(object sender, EventArgs e)
        {
            CourtPannel.Invalidate();
        }

        //エクセルを開く
        private void OpenExelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExcelWriter.Open();
        }

        //動画を開く
        private void dougaPlayerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MediaPlayer.Instance.Open();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        //プレイヤー位置をクリックしていく
        private void PlayerPositionMenuItem_Click(object sender, EventArgs e)
        {
            if (!ExcelWriter.Available())
                return;

            ExcelWriter.Instance.shotSheet.Surveying = ExcelWriter.ShotDataSheet.Surveyed.PlayerPos;
            court.ClearPosition();
        }

        //バウンド位置をクリックしていく
        private void BoundPositionMenuItem_Click(object sender, EventArgs e)
        {
            if (!ExcelWriter.Available())
                return;

            ExcelWriter.Instance.shotSheet.Surveying = ExcelWriter.ShotDataSheet.Surveyed.BoundPos;
            court.ClearPosition();
        }

    }
}
