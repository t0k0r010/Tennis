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
        MediaPlayer mediaPlayer;
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
            mediaPlayer = new MediaPlayer(WMPlayer);

            this.KeyDown += Form1_KeyDown;
            this.Resize += Form1_Resize;
        }

        void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            
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

        private void PlayerPositionMenuItem_Click(object sender, EventArgs e)
        {
            court.SetCheckBoundPosition(false);
        }

        private void BoundPositionMenuItem_Click(object sender, EventArgs e)
        {
            court.SetCheckBoundPosition(true);
        }


    }
}
