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
            Result = new TextBox();
            pictureBox1 = new PictureBox();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            SuspendLayout();
            // 
            // Reset
            // 
            Reset.BackColor = Color.FromArgb(255, 128, 128);
            Reset.Location = new Point(225, 281);
            Reset.Name = "Reset";
            Reset.Size = new Size(75, 32);
            Reset.TabIndex = 0;
            Reset.Text = "Reset";
            Reset.UseVisualStyleBackColor = false;
            Reset.Click += Reset_Click;
            // 
            // Start
            // 
            Start.BackColor = Color.FromArgb(0, 192, 0);
            Start.Location = new Point(112, 231);
            Start.Name = "Start";
            Start.Size = new Size(75, 32);
            Start.TabIndex = 1;
            Start.Text = "Start";
            Start.UseVisualStyleBackColor = false;
            Start.Click += Start_Click;
            // 
            // Result
            // 
            Result.BackColor = Color.Silver;
            Result.BorderStyle = BorderStyle.None;
            Result.Font = new Font("Microsoft JhengHei UI", 20F);
            Result.Location = new Point(77, 38);
            Result.Name = "Result";
            Result.ReadOnly = true;
            Result.Size = new Size(150, 34);
            Result.TabIndex = 2;
            Result.TextAlign = HorizontalAlignment.Center;
            Result.TextChanged += Result_TextChanged;
            // 
            // pictureBox1
            // 
            pictureBox1.Location = new Point(77, 87);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(150, 129);
            pictureBox1.TabIndex = 3;
            pictureBox1.TabStop = false;
            pictureBox1.Click += pictureBox1_Click;
            // 
            // Lottery
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.Silver;
            ClientSize = new Size(314, 321);
            Controls.Add(pictureBox1);
            Controls.Add(Result);
            Controls.Add(Start);
            Controls.Add(Reset);
            Name = "Lottery";
            Text = "Lottery";
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button Reset;
        private Button Start;
        private TextBox Result;
        private PictureBox pictureBox1;
    }
}
