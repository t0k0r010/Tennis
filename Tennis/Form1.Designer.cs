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
            this.MakeRallyToolStrip = new System.Windows.Forms.ToolStripMenuItem();
            this.CourtPannel = new System.Windows.Forms.Panel();
            this.PosLabel = new System.Windows.Forms.Label();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.TopPlayerName = new System.Windows.Forms.Label();
            this.InputPanel = new System.Windows.Forms.Panel();
            this.BottomPlayerName = new System.Windows.Forms.Label();
            this.ChangeCourtButton = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            this.CourtPannel.SuspendLayout();
            this.InputPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.dougaPlayerToolStripMenuItem,
            this.OpenExelToolStripMenuItem,
            this.ClickModeMenuItem,
            this.MakeRallyToolStrip});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(4, 2, 0, 2);
            this.menuStrip1.Size = new System.Drawing.Size(558, 24);
            this.menuStrip1.TabIndex = 2;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // dougaPlayerToolStripMenuItem
            // 
            this.dougaPlayerToolStripMenuItem.Name = "dougaPlayerToolStripMenuItem";
            this.dougaPlayerToolStripMenuItem.Size = new System.Drawing.Size(90, 20);
            this.dougaPlayerToolStripMenuItem.Text = "動画プレイヤー";
            this.dougaPlayerToolStripMenuItem.Click += new System.EventHandler(this.dougaPlayerToolStripMenuItem_Click);
            // 
            // OpenExelToolStripMenuItem
            // 
            this.OpenExelToolStripMenuItem.Name = "OpenExelToolStripMenuItem";
            this.OpenExelToolStripMenuItem.Size = new System.Drawing.Size(157, 20);
            this.OpenExelToolStripMenuItem.Text = "新しくショットデータを作成する";
            this.OpenExelToolStripMenuItem.Click += new System.EventHandler(this.OpenExelToolStripMenuItem_Click);
            // 
            // ClickModeMenuItem
            // 
            this.ClickModeMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.PlayerPositionMenuItem,
            this.BoundPositionMenuItem});
            this.ClickModeMenuItem.Name = "ClickModeMenuItem";
            this.ClickModeMenuItem.Size = new System.Drawing.Size(73, 20);
            this.ClickModeMenuItem.Text = "クリック対象";
            // 
            // PlayerPositionMenuItem
            // 
            this.PlayerPositionMenuItem.Name = "PlayerPositionMenuItem";
            this.PlayerPositionMenuItem.Size = new System.Drawing.Size(134, 22);
            this.PlayerPositionMenuItem.Text = "選手位置";
            this.PlayerPositionMenuItem.Click += new System.EventHandler(this.PlayerPositionMenuItem_Click);
            // 
            // BoundPositionMenuItem
            // 
            this.BoundPositionMenuItem.Name = "BoundPositionMenuItem";
            this.BoundPositionMenuItem.Size = new System.Drawing.Size(134, 22);
            this.BoundPositionMenuItem.Text = "バウンド位置";
            this.BoundPositionMenuItem.Click += new System.EventHandler(this.BoundPositionMenuItem_Click);
            // 
            // MakeRallyToolStrip
            // 
            this.MakeRallyToolStrip.Name = "MakeRallyToolStrip";
            this.MakeRallyToolStrip.Size = new System.Drawing.Size(116, 20);
            this.MakeRallyToolStrip.Text = "ShotDataから変換";
            this.MakeRallyToolStrip.Click += new System.EventHandler(this.MakeRallyToolStrip_Click);
            // 
            // CourtPannel
            // 
            this.CourtPannel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CourtPannel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.CourtPannel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.CourtPannel.Controls.Add(this.PosLabel);
            this.CourtPannel.Location = new System.Drawing.Point(11, 28);
            this.CourtPannel.Margin = new System.Windows.Forms.Padding(2);
            this.CourtPannel.Name = "CourtPannel";
            this.CourtPannel.Size = new System.Drawing.Size(428, 582);
            this.CourtPannel.TabIndex = 3;
            // 
            // PosLabel
            // 
            this.PosLabel.AutoSize = true;
            this.PosLabel.Location = new System.Drawing.Point(4, 4);
            this.PosLabel.Name = "PosLabel";
            this.PosLabel.Size = new System.Drawing.Size(35, 12);
            this.PosLabel.TabIndex = 0;
            this.PosLabel.Text = "label1";
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog2";
            // 
            // TopPlayerName
            // 
            this.TopPlayerName.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.TopPlayerName.AutoSize = true;
            this.TopPlayerName.Location = new System.Drawing.Point(26, 24);
            this.TopPlayerName.Name = "TopPlayerName";
            this.TopPlayerName.Size = new System.Drawing.Size(49, 12);
            this.TopPlayerName.TabIndex = 5;
            this.TopPlayerName.Text = "Player A";
            // 
            // InputPanel
            // 
            this.InputPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.InputPanel.Controls.Add(this.BottomPlayerName);
            this.InputPanel.Controls.Add(this.ChangeCourtButton);
            this.InputPanel.Controls.Add(this.TopPlayerName);
            this.InputPanel.Location = new System.Drawing.Point(452, 28);
            this.InputPanel.Name = "InputPanel";
            this.InputPanel.Size = new System.Drawing.Size(106, 582);
            this.InputPanel.TabIndex = 8;
            // 
            // BottomPlayerName
            // 
            this.BottomPlayerName.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.BottomPlayerName.AutoSize = true;
            this.BottomPlayerName.Location = new System.Drawing.Point(26, 544);
            this.BottomPlayerName.Name = "BottomPlayerName";
            this.BottomPlayerName.Size = new System.Drawing.Size(49, 12);
            this.BottomPlayerName.TabIndex = 6;
            this.BottomPlayerName.Text = "Player B";
            // 
            // ChangeCourtButton
            // 
            this.ChangeCourtButton.Location = new System.Drawing.Point(6, 278);
            this.ChangeCourtButton.Name = "ChangeCourtButton";
            this.ChangeCourtButton.Size = new System.Drawing.Size(92, 34);
            this.ChangeCourtButton.TabIndex = 7;
            this.ChangeCourtButton.Text = "ChangeCourt";
            this.ChangeCourtButton.UseVisualStyleBackColor = true;
            this.ChangeCourtButton.Click += new System.EventHandler(this.ChangeCourtButton_Click);
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(558, 614);
            this.Controls.Add(this.InputPanel);
            this.Controls.Add(this.CourtPannel);
            this.Controls.Add(this.menuStrip1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form1";
            this.Text = "Form1";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.CourtPannel.ResumeLayout(false);
            this.CourtPannel.PerformLayout();
            this.InputPanel.ResumeLayout(false);
            this.InputPanel.PerformLayout();
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
        private System.Windows.Forms.Label TopPlayerName;
        private System.Windows.Forms.Button ChangeCourtButton;
        private System.Windows.Forms.Panel InputPanel;
        private System.Windows.Forms.Label BottomPlayerName;
        private System.Windows.Forms.ToolStripMenuItem MakeRallyToolStrip;
        private System.Windows.Forms.Label PosLabel;
    }
}

