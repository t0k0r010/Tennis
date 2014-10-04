using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Tennis
{
    public partial class MediaPlayerForm : Form
    {
        public static MediaPlayerForm Instance { get; private set; }

        //動画を開く
        public static void Open()
        {
            //すでにファイルを開いているときは,新しく開きなおすか確認する.
            if (Instance != null)
            {
                if (MessageBox.Show("確認", "今あるファイルを閉じて別のファイルを開きますか?", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                    return;
            }
            //ファイルを開く
            OpenFileDialog dialog = new OpenFileDialog();
            //開くボタンを押したとき
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string fileName = dialog.FileName;

                Instance = new MediaPlayerForm(fileName);
                Instance.Show();
            }
        }

        public MediaPlayerForm(string fileName)
        {
            InitializeComponent();
            axWindowsMediaPlayer1.URL = fileName;
        }

        public AxWMPLib.AxWindowsMediaPlayer GetMediaPlayer()
        {
            return axWindowsMediaPlayer1;
        }

        //現在の動画の位置を hh:mm:ss:ff で返す
        public string GetCurrentTimeText()
        {
            double time = MediaPlayerForm.Instance.GetMediaPlayer().Ctlcontrols.currentPosition;

            int hour = (int)Math.Floor(time / 3600);
            int minute = (int)Math.Floor((time - 3600 * hour) / 60);
            int second = (int)Math.Floor(time - 3600 * hour - 60 * minute);
            int milliSec = ((int)Math.Floor(time * 100)) % 100; //ミリ秒を2桁
            TimeSpan d = TimeSpan.FromSeconds(time);

            return d.ToString(@"hh\:mm\:ss");
        }
    }
}
