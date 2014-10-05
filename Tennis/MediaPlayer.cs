using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Tennis
{
    class MediaPlayer
    {
        public static MediaPlayer Instance = null;

        public static bool Available()
        {
            return Instance != null;
        }

        AxWMPLib.AxWindowsMediaPlayer player = null;

        public MediaPlayer(AxWMPLib.AxWindowsMediaPlayer player)
        {
            this.player = player;
            Instance = this;
        }

        public void Open()
        {
            //ファイルを開く
            OpenFileDialog dialog = new OpenFileDialog();
            //開くボタンを押したとき
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                player.URL = dialog.FileName;
                player.Ctlcontrols.stop();
            }
        }

        //現在の動画の位置を hh:mm:ss で返す
        public string GetCurrentTimeText()
        {
            double time = player.Ctlcontrols.currentPosition;

            int hour = (int)Math.Floor(time / 3600);
            int minute = (int)Math.Floor((time - 3600 * hour) / 60);
            int second = (int)Math.Floor(time - 3600 * hour - 60 * minute);
            int milliSec = ((int)Math.Floor(time * 100)) % 100; //ミリ秒を2桁
            TimeSpan d = TimeSpan.FromSeconds(time);

            return d.ToString(@"hh\:mm\:ss");
        }
    }
}
