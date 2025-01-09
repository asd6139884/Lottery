namespace Lottery
{
    partial class Lottery
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            ResetButton = new Button();
            StartButton = new Button();
            ResultText = new TextBox();
            ResultImage = new PictureBox();
            ExecuteImage = new PictureBox();
            StartImage = new PictureBox();
            listBox1 = new ListBox();
            ((System.ComponentModel.ISupportInitialize)ResultImage).BeginInit();
            ((System.ComponentModel.ISupportInitialize)ExecuteImage).BeginInit();
            ((System.ComponentModel.ISupportInitialize)StartImage).BeginInit();
            SuspendLayout();
            // 
            // ResetButton
            // 
            ResetButton.Anchor = AnchorStyles.None;
            ResetButton.BackColor = Color.FromArgb(255, 128, 128);
            ResetButton.Location = new Point(354, 324);
            ResetButton.Name = "ResetButton";
            ResetButton.Size = new Size(75, 30);
            ResetButton.TabIndex = 0;
            ResetButton.Text = "Reset";
            ResetButton.UseVisualStyleBackColor = false;
            ResetButton.Click += Reset_Click;
            ResetButton.Resize += Form1_Resize;
            // 
            // StartButton
            // 
            StartButton.Anchor = AnchorStyles.None;
            StartButton.BackColor = Color.FromArgb(0, 192, 0);
            StartButton.Location = new Point(117, 295);
            StartButton.Name = "StartButton";
            StartButton.Size = new Size(75, 30);
            StartButton.TabIndex = 1;
            StartButton.Text = "Start";
            StartButton.UseVisualStyleBackColor = false;
            StartButton.Click += Start_Click;
            StartButton.Resize += Form1_Resize;
            // 
            // ResultText
            // 
            ResultText.Anchor = AnchorStyles.None;
            ResultText.BackColor = Color.Gold;
            ResultText.BorderStyle = BorderStyle.None;
            ResultText.Font = new Font("Microsoft JhengHei UI", 20F);
            ResultText.Location = new Point(32, 26);
            ResultText.Name = "ResultText";
            ResultText.ReadOnly = true;
            ResultText.Size = new Size(252, 34);
            ResultText.TabIndex = 2;
            ResultText.TextAlign = HorizontalAlignment.Center;
            ResultText.Resize += Form1_Resize;
            // 
            // ResultImage
            // 
            ResultImage.Anchor = AnchorStyles.None;
            ResultImage.Location = new Point(32, 82);
            ResultImage.Name = "ResultImage";
            ResultImage.Size = new Size(252, 193);
            ResultImage.TabIndex = 3;
            ResultImage.TabStop = false;
            ResultImage.Resize += Form1_Resize;
            // 
            // ExecuteImage
            // 
            ExecuteImage.Anchor = AnchorStyles.None;
            ExecuteImage.Location = new Point(32, 26);
            ExecuteImage.Name = "ExecuteImage";
            ExecuteImage.Size = new Size(252, 249);
            ExecuteImage.TabIndex = 4;
            ExecuteImage.TabStop = false;
            ExecuteImage.Resize += Form1_Resize;
            // 
            // StartImage
            // 
            StartImage.Anchor = AnchorStyles.None;
            StartImage.Location = new Point(32, 26);
            StartImage.Name = "StartImage";
            StartImage.Size = new Size(252, 249);
            StartImage.SizeMode = PictureBoxSizeMode.Zoom;
            StartImage.TabIndex = 5;
            StartImage.TabStop = false;
            StartImage.Resize += Form1_Resize;
            // 
            // listBox1
            // 
            listBox1.ForeColor = SystemColors.WindowText;
            listBox1.FormattingEnabled = true;
            listBox1.ItemHeight = 15;
            listBox1.Location = new Point(309, 26);
            listBox1.Name = "listBox1";
            listBox1.Size = new Size(120, 259);
            listBox1.TabIndex = 6;
            // 
            // Lottery
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.Gold;
            ClientSize = new Size(441, 361);
            Controls.Add(StartImage);
            Controls.Add(ExecuteImage);
            Controls.Add(ResultImage);
            Controls.Add(ResultText);
            Controls.Add(listBox1);
            Controls.Add(ResetButton);
            Controls.Add(StartButton);
            Name = "Lottery";
            Text = "抽獎啦~";
            ((System.ComponentModel.ISupportInitialize)ResultImage).EndInit();
            ((System.ComponentModel.ISupportInitialize)ExecuteImage).EndInit();
            ((System.ComponentModel.ISupportInitialize)StartImage).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button ResetButton;
        private Button StartButton;
        private TextBox ResultText;
        private PictureBox ResultImage;
        private PictureBox ExecuteImage;
        private PictureBox StartImage;
        private ListBox listBox1;
    }
}
