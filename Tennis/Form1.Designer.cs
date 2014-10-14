namespace Tennis
{
    partial class Form1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.serveResult = new System.Windows.Forms.GroupBox();
            this.radioButton5 = new System.Windows.Forms.RadioButton();
            this.radioButton4 = new System.Windows.Forms.RadioButton();
            this.radioButton3 = new System.Windows.Forms.RadioButton();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.dougaPlayerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.OpenExelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.CourtPannel = new System.Windows.Forms.Panel();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.WMPlayer = new AxWMPLib.AxWindowsMediaPlayer();
            this.ClickModeMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.PlayerPositionMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.BoundPositionMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1.SuspendLayout();
            this.serveResult.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.WMPlayer)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.groupBox1.Controls.Add(this.serveResult);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Location = new System.Drawing.Point(348, 28);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(295, 526);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "groupBox1";
            // 
            // serveResult
            // 
            this.serveResult.Controls.Add(this.radioButton5);
            this.serveResult.Controls.Add(this.radioButton4);
            this.serveResult.Controls.Add(this.radioButton3);
            this.serveResult.Location = new System.Drawing.Point(15, 99);
            this.serveResult.Margin = new System.Windows.Forms.Padding(2);
            this.serveResult.Name = "serveResult";
            this.serveResult.Padding = new System.Windows.Forms.Padding(2);
            this.serveResult.Size = new System.Drawing.Size(235, 50);
            this.serveResult.TabIndex = 1;
            this.serveResult.TabStop = false;
            this.serveResult.Text = "サーブ球種";
            // 
            // radioButton5
            // 
            this.radioButton5.AutoSize = true;
            this.radioButton5.Location = new System.Drawing.Point(176, 30);
            this.radioButton5.Margin = new System.Windows.Forms.Padding(2);
            this.radioButton5.Name = "radioButton5";
            this.radioButton5.Size = new System.Drawing.Size(58, 16);
            this.radioButton5.TabIndex = 2;
            this.radioButton5.TabStop = true;
            this.radioButton5.Text = "スライス";
            this.radioButton5.UseVisualStyleBackColor = true;
            // 
            // radioButton4
            // 
            this.radioButton4.AutoSize = true;
            this.radioButton4.Location = new System.Drawing.Point(88, 30);
            this.radioButton4.Margin = new System.Windows.Forms.Padding(2);
            this.radioButton4.Name = "radioButton4";
            this.radioButton4.Size = new System.Drawing.Size(50, 16);
            this.radioButton4.TabIndex = 1;
            this.radioButton4.TabStop = true;
            this.radioButton4.Text = "スピン";
            this.radioButton4.UseVisualStyleBackColor = true;
            // 
            // radioButton3
            // 
            this.radioButton3.AutoSize = true;
            this.radioButton3.Location = new System.Drawing.Point(4, 30);
            this.radioButton3.Margin = new System.Windows.Forms.Padding(2);
            this.radioButton3.Name = "radioButton3";
            this.radioButton3.Size = new System.Drawing.Size(54, 16);
            this.radioButton3.TabIndex = 0;
            this.radioButton3.TabStop = true;
            this.radioButton3.Text = "フラット";
            this.radioButton3.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.radioButton2);
            this.groupBox2.Controls.Add(this.radioButton1);
            this.groupBox2.Location = new System.Drawing.Point(15, 34);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox2.Size = new System.Drawing.Size(219, 38);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "サーブ種別";
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(120, 18);
            this.radioButton2.Margin = new System.Windows.Forms.Padding(2);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(41, 16);
            this.radioButton2.TabIndex = 1;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "2nd";
            this.radioButton2.UseVisualStyleBackColor = true;
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Location = new System.Drawing.Point(16, 18);
            this.radioButton1.Margin = new System.Windows.Forms.Padding(2);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(39, 16);
            this.radioButton1.TabIndex = 0;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "1st";
            this.radioButton1.UseVisualStyleBackColor = true;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.dougaPlayerToolStripMenuItem,
            this.OpenExelToolStripMenuItem,
            this.ClickModeMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(4, 2, 0, 2);
            this.menuStrip1.Size = new System.Drawing.Size(1242, 26);
            this.menuStrip1.TabIndex = 2;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // dougaPlayerToolStripMenuItem
            // 
            this.dougaPlayerToolStripMenuItem.Name = "dougaPlayerToolStripMenuItem";
            this.dougaPlayerToolStripMenuItem.Size = new System.Drawing.Size(104, 22);
            this.dougaPlayerToolStripMenuItem.Text = "動画プレイヤー";
            this.dougaPlayerToolStripMenuItem.Click += new System.EventHandler(this.dougaPlayerToolStripMenuItem_Click);
            // 
            // OpenExelToolStripMenuItem
            // 
            this.OpenExelToolStripMenuItem.Name = "OpenExelToolStripMenuItem";
            this.OpenExelToolStripMenuItem.Size = new System.Drawing.Size(104, 22);
            this.OpenExelToolStripMenuItem.Text = "エクセルを開く";
            this.OpenExelToolStripMenuItem.Click += new System.EventHandler(this.OpenExelToolStripMenuItem_Click);
            // 
            // CourtPannel
            // 
            this.CourtPannel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CourtPannel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.CourtPannel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.CourtPannel.Location = new System.Drawing.Point(5, 28);
            this.CourtPannel.Margin = new System.Windows.Forms.Padding(2);
            this.CourtPannel.Name = "CourtPannel";
            this.CourtPannel.Size = new System.Drawing.Size(339, 526);
            this.CourtPannel.TabIndex = 3;
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog2";
            // 
            // WMPlayer
            // 
            this.WMPlayer.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.WMPlayer.Enabled = true;
            this.WMPlayer.Location = new System.Drawing.Point(648, 27);
            this.WMPlayer.Name = "WMPlayer";
            this.WMPlayer.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("WMPlayer.OcxState")));
            this.WMPlayer.Size = new System.Drawing.Size(582, 526);
            this.WMPlayer.TabIndex = 4;
            // 
            // ClickModeMenuItem
            // 
            this.ClickModeMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.PlayerPositionMenuItem,
            this.BoundPositionMenuItem});
            this.ClickModeMenuItem.Name = "ClickModeMenuItem";
            this.ClickModeMenuItem.Size = new System.Drawing.Size(92, 22);
            this.ClickModeMenuItem.Text = "クリック対象";
            // 
            // PlayerPositionMenuItem
            // 
            this.PlayerPositionMenuItem.Name = "PlayerPositionMenuItem";
            this.PlayerPositionMenuItem.Size = new System.Drawing.Size(152, 22);
            this.PlayerPositionMenuItem.Text = "選手位置";
            this.PlayerPositionMenuItem.Click += new System.EventHandler(this.PlayerPositionMenuItem_Click);
            // 
            // BoundPositionMenuItem
            // 
            this.BoundPositionMenuItem.Name = "BoundPositionMenuItem";
            this.BoundPositionMenuItem.Size = new System.Drawing.Size(152, 22);
            this.BoundPositionMenuItem.Text = "バウンド位置";
            this.BoundPositionMenuItem.Click += new System.EventHandler(this.BoundPositionMenuItem_Click);
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(1242, 558);
            this.Controls.Add(this.CourtPannel);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.serveResult.ResumeLayout(false);
            this.serveResult.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.WMPlayer)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem dougaPlayerToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.GroupBox serveResult;
        private System.Windows.Forms.RadioButton radioButton5;
        private System.Windows.Forms.RadioButton radioButton4;
        private System.Windows.Forms.RadioButton radioButton3;
        private System.Windows.Forms.Panel CourtPannel;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.ToolStripMenuItem OpenExelToolStripMenuItem;
        private AxWMPLib.AxWindowsMediaPlayer WMPlayer;
        private System.Windows.Forms.ToolStripMenuItem ClickModeMenuItem;
        private System.Windows.Forms.ToolStripMenuItem PlayerPositionMenuItem;
        private System.Windows.Forms.ToolStripMenuItem BoundPositionMenuItem;
    }
}

