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
            Reset = new Button();
            Start = new Button();
            ResultText = new TextBox();
            ResultImage = new PictureBox();
            ExecuteImage = new PictureBox();
            StartImage = new PictureBox();
            ((System.ComponentModel.ISupportInitialize)ResultImage).BeginInit();
            ((System.ComponentModel.ISupportInitialize)ExecuteImage).BeginInit();
            ((System.ComponentModel.ISupportInitialize)StartImage).BeginInit();
            SuspendLayout();
            // 
            // Reset
            // 
            Reset.Anchor = AnchorStyles.None;
            Reset.BackColor = Color.FromArgb(255, 128, 128);
            Reset.Location = new Point(255, 324);
            Reset.Name = "Reset";
            Reset.Size = new Size(75, 30);
            Reset.TabIndex = 0;
            Reset.Text = "Reset";
            Reset.UseVisualStyleBackColor = false;
            Reset.Click += Reset_Click;
            Reset.Resize += Form1_Resize;
            // 
            // Start
            // 
            Start.Anchor = AnchorStyles.None;
            Start.BackColor = Color.FromArgb(0, 192, 0);
            Start.Location = new Point(130, 301);
            Start.Name = "Start";
            Start.Size = new Size(75, 30);
            Start.TabIndex = 1;
            Start.Text = "Start";
            Start.UseVisualStyleBackColor = false;
            Start.Click += Start_Click;
            Start.Resize += Form1_Resize;
            // 
            // ResultText
            // 
            ResultText.Anchor = AnchorStyles.None;
            ResultText.BackColor = Color.Gold;
            ResultText.BorderStyle = BorderStyle.None;
            ResultText.Font = new Font("Microsoft JhengHei UI", 20F);
            ResultText.Location = new Point(60, 32);
            ResultText.Name = "ResultText";
            ResultText.ReadOnly = true;
            ResultText.Size = new Size(206, 34);
            ResultText.TabIndex = 2;
            ResultText.TextAlign = HorizontalAlignment.Center;
            ResultText.Resize += Form1_Resize;
            // 
            // ResultImage
            // 
            ResultImage.Anchor = AnchorStyles.None;
            ResultImage.Location = new Point(36, 82);
            ResultImage.Name = "ResultImage";
            ResultImage.Size = new Size(261, 199);
            ResultImage.TabIndex = 3;
            ResultImage.TabStop = false;
            ResultImage.Resize += Form1_Resize;
            // 
            // ExecuteImage
            // 
            ExecuteImage.Anchor = AnchorStyles.None;
            ExecuteImage.Location = new Point(49, 32);
            ExecuteImage.Name = "ExecuteImage";
            ExecuteImage.Size = new Size(236, 249);
            ExecuteImage.TabIndex = 4;
            ExecuteImage.TabStop = false;
            ExecuteImage.Resize += Form1_Resize;
            // 
            // StartImage
            // 
            StartImage.Anchor = AnchorStyles.None;
            StartImage.Location = new Point(36, 32);
            StartImage.Name = "StartImage";
            StartImage.Size = new Size(261, 249);
            StartImage.SizeMode = PictureBoxSizeMode.Zoom;
            StartImage.TabIndex = 5;
            StartImage.TabStop = false;
            StartImage.Resize += Form1_Resize;
            // 
            // Lottery
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.Gold;
            ClientSize = new Size(334, 361);
            Controls.Add(Reset);
            Controls.Add(Start);
            Controls.Add(StartImage);
            Controls.Add(ExecuteImage);
            Controls.Add(ResultText);
            Controls.Add(ResultImage);
            Name = "Lottery";
            Text = "抽獎啦~";
            ((System.ComponentModel.ISupportInitialize)ResultImage).EndInit();
            ((System.ComponentModel.ISupportInitialize)ExecuteImage).EndInit();
            ((System.ComponentModel.ISupportInitialize)StartImage).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button Reset;
        private Button Start;
        private TextBox ResultText;
        private PictureBox ResultImage;
        private PictureBox ExecuteImage;
        private PictureBox StartImage;
    }
}
