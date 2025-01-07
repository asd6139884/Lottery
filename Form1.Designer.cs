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
            pictureBox2 = new PictureBox();
            pictureBox3 = new PictureBox();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox3).BeginInit();
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
            Start.Location = new Point(112, 241);
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
            // pictureBox2
            // 
            pictureBox2.Location = new Point(63, 16);
            pictureBox2.Name = "pictureBox2";
            pictureBox2.Size = new Size(181, 209);
            pictureBox2.TabIndex = 4;
            pictureBox2.TabStop = false;
            // 
            // pictureBox3
            // 
            pictureBox3.Location = new Point(32, 16);
            pictureBox3.Name = "pictureBox3";
            pictureBox3.Size = new Size(241, 209);
            pictureBox3.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox3.TabIndex = 5;
            pictureBox3.TabStop = false;
            // 
            // Lottery
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.Silver;
            ClientSize = new Size(314, 321);
            Controls.Add(pictureBox3);
            Controls.Add(pictureBox2);
            Controls.Add(Result);
            Controls.Add(pictureBox1);
            Controls.Add(Start);
            Controls.Add(Reset);
            Name = "Lottery";
            Text = "Lottery";
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox2).EndInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox3).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button Reset;
        private Button Start;
        private TextBox Result;
        private PictureBox pictureBox1;
        private PictureBox pictureBox2;
        private PictureBox pictureBox3;
    }
}
