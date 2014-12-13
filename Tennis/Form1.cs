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
            this.CourtPannel.MouseMove += MoveMouse;
            this.KeyDown += Form1_KeyDown;
            this.Resize  += Form1_Resize;

            PosLabel.MouseMove += MoveMouse;
        }

        void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            
        }

        void MoveMouse(object sender, MouseEventArgs e)
        {
            PosLabel.Text = court.ToRealUnit(e.Location).ToString();
        }
        void ClickCourt(object sender, MouseEventArgs e)
        {
            if (!Form1.IsStarted)
                return;

            ExcelWriter writer = ExcelWriter.Instance;
            Point p = new Point(e.X, e.Y);
            if( e.Button == MouseButtons.Left )
            {
                court.AddPosition(p);                                    //コートに新しい位置を追加
                writer.shotSheet.SetPosition("", court.ToRealUnit(p));
            }
            else if(e.Button == MouseButtons.Right)
            {
                if(court.LastSurveyed != Court.Surveyed.BoundPos && 
                    court.Positions_p[Court.Surveyed.HitterPos].Count != court.Positions_p[Court.Surveyed.RecieverPos].Count)
                {
                    MessageBox.Show("被打選手の位置をクリックしてください", "注意");
                    return;
                }
                writer.rallySheet.EndRally();
                writer.shotSheet.EndRally();    //エクセルにラインを書き込む
                court.ClearPosition();          //コートに書いている位置を削除
            }
            CourtPannel.Invalidate(); //再描画命令
        }

        void Form1_Resize(object sender, EventArgs e)
        {
            CourtPannel.Invalidate();
            int padding = 30;
            TopPlayerName.Location = new Point((InputPanel.Width - TopPlayerName.Width) / 2, padding);
            BottomPlayerName.Location = new Point( (InputPanel.Width - BottomPlayerName.Width) / 2, InputPanel.Height - BottomPlayerName.Height - padding);
            ChangeCourtButton.Location = new Point((InputPanel.Width - ChangeCourtButton.Width) / 2, (InputPanel.Height - ChangeCourtButton.Height) / 2);
        }

        //エクセルを開く
        private void OpenExelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExcelWriter.Open();
        }

        //動画を開く
        private void dougaPlayerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //MediaPlayer.Instance.Open();
        }

        //プレイヤー位置をクリックしていく
        private void PlayerPositionMenuItem_Click(object sender, EventArgs e)
        {
            if (!ExcelWriter.Available())
                return;

            ExcelWriter.Instance.shotSheet.Surveying = ExcelWriter.ShotDataSheet.Surveyed.PlayerPos;
            court.ChangeSurveyed(Court.Surveyed.HitterPos);
        }

        //バウンド位置をクリックしていく
        private void BoundPositionMenuItem_Click(object sender, EventArgs e)
        {
            if (!ExcelWriter.Available())
                return;

            ExcelWriter.Instance.shotSheet.Surveying = ExcelWriter.ShotDataSheet.Surveyed.BoundPos;
            court.ChangeSurveyed(Court.Surveyed.BoundPos);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void ChangeCourtButton_Click(object sender, EventArgs e)
        {
            //コートを入れ替え
            if(court.UpperPlayer == Court.Players.PlayerA)
            {
                court.UpperPlayer = Court.Players.PlayerB;
                TopPlayerName.Text = "Player B";
                BottomPlayerName.Text = "Player A";
            }
            else
            {
                court.UpperPlayer = Court.Players.PlayerA;
                TopPlayerName.Text = "Player A";
                BottomPlayerName.Text = "Player B";
            }
        }

        private void MakeRallyToolStrip_Click(object sender, EventArgs e)
        {
            RallyDataMaker.MakeRally();
        }

    }
}
