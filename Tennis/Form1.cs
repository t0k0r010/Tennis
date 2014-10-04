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
        double mediaPos = 0;

        Court court;

        MediaPlayerForm mediaPlayer = null;

        //両方とも開いているか
        public static bool IsStarted { get { return ExcelWriter.Instance != null && MediaPlayerForm.Instance != null; } }
        public Form1()
        {
            //初期設定を書く
            InitializeComponent();

            court = new Court(this.CourtPannel);
            //Observeという関数を別に実行する.
            Thread thread = new Thread(Observe);
            thread.Start();
        }


        delegate void SetMedia();

        private void Observe()
        {
            while (!this.IsDisposed)
            {
                if (mediaPlayer == null || mediaPlayer.GetMediaPlayer().IsDisposed)
                    continue;


                //moviePos.Text = mediaPlayer.GetMediaPlayer().Ctlcontrols.currentPosition.ToString();
               // mediaPos = mediaPlayer.GetMediaPlayer().Ctlcontrols.currentPosition;
               // this.Invoke( new SetMedia(SetMediaPos));
            }
        }

        private void SetMediaPos()
        {
            moviePos.Text = mediaPos.ToString();
        }

        private void OpenExelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExcelWriter.Open();
        }

        private void dougaPlayerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MediaPlayerForm.Open();
        }
    }
}
