﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft;
using Microsoft.Office;
using Microsoft.Office.Interop;

//エクセルファイルへの書き出しを行うクラス
//どんなクラスからもアクセスできるが,このクラスは誰も保持しない.
namespace Tennis
{
    using Excel = Microsoft.Office.Interop.Excel;
    class ExcelWriter
    {
        public static ExcelWriter Instance {get; private set;}

       // System.Diagnostics.Process excelProcess = null;
        Excel.Application excelApp;
        Excel.Workbook wb;

        //現在の行番号
        private uint currentRow = 5;    //これは専用のsetterでしかアクセスしないようにする.   
        public uint CurrentRow { get { return currentRow; } }    //取得専用

        //ファイルを開く
        public static void Open()
        {
            //すでにファイルを開いているときは,新しく開きなおすか確認する.
            if ( Available() )
            {
                if (MessageBox.Show("今あるファイルを閉じて別のファイルを開きますか?", "確認", 
                    MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                    return;

                Instance.Close();
            }

            //ファイルを開く
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "エクセルファイルを開く";
            //開くボタンを押したとき
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string fileName = dialog.FileName;

                try
                {
                    Instance = new ExcelWriter(fileName);
                }
                catch(Exception ex)
                {
                    MessageBox.Show( ex.Message, "開けませんでした");
                    Instance = null;
                }
            }
        }
        public static bool Available()
        {
            if (Instance == null || Instance.excelApp == null)
                return false;

            return Instance.excelApp.Visible;
        }

        void ClosedEvent(object sender, EventArgs e)
        {
            Instance = null;
        }

        ExcelWriter(string fileName)
        {
            excelApp = new Excel.Application();
            excelApp.Visible = true;
            wb = excelApp.Workbooks.Open(Filename: fileName);
            try
            {
                ((Excel.Worksheet)wb.Sheets[1]).Select();
                MakeTemplate();
            }
            catch (Exception ex)
            {
                wb.Close(false);
                excelApp.Quit();
                throw new Exception();
            }
        }


        void Close()
        {
            if (excelApp != null)
            {
                excelApp.Quit();
            }
        }

        public void MoveToFistLine()
        {
            currentRow = 5;
        }
        //現在見ている行を次に進める.
        public void MoveToNextLine()
        {
            currentRow++;
        }

        //バウンドした座標を書き込む. 
        public void SetBoundPosition(string time, float x, float y)
        {
            SetPosition(excelApp.get_Range("AB" + CurrentRow, "AC" + CurrentRow), time, x, y);
        }

        public void SetHitterPosition(string time, float x, float y)
        {
            SetPosition(excelApp.get_Range("AD" + CurrentRow, "AE" + CurrentRow), time, x, y);
        }

        public void SetRecieverPosition(string time, float x, float y)
        {
            SetPosition(excelApp.get_Range("AF" + CurrentRow, "AG" + CurrentRow), time, x, y);
        }

        // range(1行2列) にx,y座標を入れて, B列に時間を入れる.
        private void SetPosition(Excel.Range range, string time, float x, float y)
        {
            range.get_Range("A1").Value2 = x;
            range.get_Range("B1").Value2 = y;

            excelApp.get_Range("B" + CurrentRow).Value2 = time;
        }

        //ラリーの終わりの表す線を書き込む
        public void WriteLine()
        {
            Excel.Range range = excelApp.get_Range("A" + CurrentRow, "CN" + CurrentRow);
            range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium;
        }

        //列のテンプレートを作成
        // title  : 黄色の文字で書かれるその項目のタイトル.
        // left   : その項目の左端の列番号( "A" や "D" と指定する)
        // right  : 右端の列番号
        // return : 指定した範囲のセルが戻る "A", "D" と指定すると "A1" から"D4"が返る
        Excel.Range MakeCol(string title, string left, string right)
        {
            Excel.XlBorderWeight lineWeight = Excel.XlBorderWeight.xlMedium;

            Excel.Range r = excelApp.get_Range(left + "1", right + "4");
            r.get_Range("A1").Value2 = title;
            r.get_Range("A1").Interior.Color = 0x44FFFF;
            r.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = lineWeight;
            r.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = lineWeight;

            //とりあえず,300行分線を引いておく
            Excel.Range r2 = excelApp.get_Range(left + "1", right + "300");

            r2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = lineWeight;

            return r;
        }

        //テンプレートの生成
        void MakeTemplate()
        {
            //時間の列のテンプレート
            Excel.Range time = MakeCol("時間", "A", "B");
            time.get_Range("A2").Value2 = "再生時間";
            time.get_Range("B2").Value2 = "到達時間";

            //カウントの列のテンプレート
            Excel.Range count = MakeCol("カウント", "C", "E");
            count.get_Range("A2").Value2 = "セット";
            count.get_Range("B2").Value2 = "ゲーム";
            count.get_Range("C2").Value2 = "ポイント";

            //番号の列
            Excel.Range number = MakeCol("番号", "F", "G");
            number.get_Range("A2").Value2 = "ラリー";
            number.get_Range("B2").Value2 = "ショット";

            //種別の列
            Excel.Range kind = MakeCol("種別", "H", "K");
            kind.get_Range("A2").Value2 = "サーブ";
            kind.get_Range("B2").Value2 = "リターン";
            kind.get_Range("C2").Value2 = "ラリー";
            kind.get_Range("D2").Value2 = "endショット";

            //サーブの列
            Excel.Range serve = MakeCol("サーブ", "L", "AA");
            /*
             サーブの種類の列名を書く
             * ここは自分で書いてみてください.
             * 小さいラインの書き方は今度教えます.
             */


            //「座標」項目テンプレートの作成
            Excel.Range coordinate = MakeCol("座標", "AB", "AJ");//exelApp.get_Range("AB1", "AJ4");
            coordinate.get_Range("A2").Value2 = "バウンド";
            coordinate.get_Range("A3").Value2 = "x";
            coordinate.get_Range("B3").Value2 = "y";

            coordinate.get_Range("C2").Value2 = "打選手";
            coordinate.get_Range("C3").Value2 = "x";
            coordinate.get_Range("D3").Value2 = "y";

            coordinate.get_Range("E2").Value2 = "被打選手";
            coordinate.get_Range("E3").Value2 = "x";
            coordinate.get_Range("F3").Value2 = "y";
        }


        /*
        void SetupProcess()
        {
            foreach (var p in System.Diagnostics.Process.GetProcesses()) //Get the relevant Excel process.
            {
                if (p.MainWindowHandle == new IntPtr(excelApp.Hwnd))
                {
                    excelProcess = p;
                    break;
                }
            }

            if (excelProcess != null)
            {
                Console.WriteLine("not NULL");
            }
        }*/
    }
}
