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
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.dougaPlayerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.OpenExelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ClickModeMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.PlayerPositionMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.BoundPositionMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.CourtPannel = new System.Windows.Forms.Panel();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
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
            this.menuStrip1.Size = new System.Drawing.Size(556, 31);
            this.menuStrip1.TabIndex = 2;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // dougaPlayerToolStripMenuItem
            // 
            this.dougaPlayerToolStripMenuItem.Name = "dougaPlayerToolStripMenuItem";
            this.dougaPlayerToolStripMenuItem.Size = new System.Drawing.Size(127, 27);
            this.dougaPlayerToolStripMenuItem.Text = "動画プレイヤー";
            this.dougaPlayerToolStripMenuItem.Click += new System.EventHandler(this.dougaPlayerToolStripMenuItem_Click);
            // 
            // OpenExelToolStripMenuItem
            // 
            this.OpenExelToolStripMenuItem.Name = "OpenExelToolStripMenuItem";
            this.OpenExelToolStripMenuItem.Size = new System.Drawing.Size(127, 27);
            this.OpenExelToolStripMenuItem.Text = "エクセルを開く";
            this.OpenExelToolStripMenuItem.Click += new System.EventHandler(this.OpenExelToolStripMenuItem_Click);
            // 
            // ClickModeMenuItem
            // 
            this.ClickModeMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.PlayerPositionMenuItem,
            this.BoundPositionMenuItem});
            this.ClickModeMenuItem.Name = "ClickModeMenuItem";
            this.ClickModeMenuItem.Size = new System.Drawing.Size(112, 27);
            this.ClickModeMenuItem.Text = "クリック対象";
            // 
            // PlayerPositionMenuItem
            // 
            this.PlayerPositionMenuItem.Name = "PlayerPositionMenuItem";
            this.PlayerPositionMenuItem.Size = new System.Drawing.Size(170, 28);
            this.PlayerPositionMenuItem.Text = "選手位置";
            this.PlayerPositionMenuItem.Click += new System.EventHandler(this.PlayerPositionMenuItem_Click);
            // 
            // BoundPositionMenuItem
            // 
            this.BoundPositionMenuItem.Name = "BoundPositionMenuItem";
            this.BoundPositionMenuItem.Size = new System.Drawing.Size(170, 28);
            this.BoundPositionMenuItem.Text = "バウンド位置";
            this.BoundPositionMenuItem.Click += new System.EventHandler(this.BoundPositionMenuItem_Click);
            // 
            // CourtPannel
            // 
            this.CourtPannel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CourtPannel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.CourtPannel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.CourtPannel.Location = new System.Drawing.Point(11, 28);
            this.CourtPannel.Margin = new System.Windows.Forms.Padding(2);
            this.CourtPannel.Name = "CourtPannel";
            this.CourtPannel.Size = new System.Drawing.Size(534, 912);
            this.CourtPannel.TabIndex = 3;
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog2";
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(556, 944);
            this.Controls.Add(this.CourtPannel);
            this.Controls.Add(this.menuStrip1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem dougaPlayerToolStripMenuItem;
        private System.Windows.Forms.Panel CourtPannel;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.ToolStripMenuItem OpenExelToolStripMenuItem;

        private System.Windows.Forms.ToolStripMenuItem ClickModeMenuItem;
        private System.Windows.Forms.ToolStripMenuItem PlayerPositionMenuItem;
        private System.Windows.Forms.ToolStripMenuItem BoundPositionMenuItem;
    }
}

