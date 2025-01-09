using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml; // EPPlus �R�W�Ŷ�


namespace Lottery
{
    public partial class Lottery : Form
    {
        private Random random; // �H���ƥͦ���
        private System.Windows.Forms.Timer loadTimer;
        private string file = "Options.xlsx";
        private DataTable dataTable; // �Ω�s����

        //�����T
        private Size LotteryOriginSize; // ������l�j�p
        private Size previousSize; // �����j�p
        private float aspectRatio; // �e����
        // Start���s
        private Size StartButtonOriginSize; // ��l�j�p
        private float StartButtonOriginFontSize; // ��l�r��j�p
        private readonly PointF StartButtonPositionRatio; // ��m���
        //  Reset���s
        private Size ResetButtonOriginSize; //  ��l�j�p
        private float ResetButtonOriginFontSize; // ��l�r��j�p
        private readonly PointF ResetButtonPositionRatio; // ��m���
        // StartImage
        private Size StartImageOriginSize;  // ��l�j�p
        private readonly PointF StartImagePositionRatio; // ��m���
        // ExecuteImage
        private Size ExecuteImageOriginSize;  // ��l�j�p
        private readonly PointF ExecuteImagePositionRatio; // ��m���
        // ResultImage
        private Size ResultImageOriginSize;  // ��l�j�p
        private readonly PointF ResultImagePositionRatio; // ��m���
        // ResultText
        private Size ResultTextOriginSize; //��l�j�p
        private float ResultTextOriginFontSize; // ��l�r��j�p
        private readonly PointF ResultTextPositionRatio; // ��m���
        // listBox1
        private Size listBox1OriginSize; // ��l�j�p
        private float listBox1OriginFontSize; // ��l�r��j�p
        private readonly PointF listBox1PositionRatio; // ��m���

        public Lottery()
        {
            InitializeComponent();

            // ���o�����T
            LotteryOriginSize = this.ClientSize;
            previousSize = this.Size;
            aspectRatio = (float)LotteryOriginSize.Width / LotteryOriginSize.Height;

            StartButtonOriginSize = StartButton.Size;
            StartButtonOriginFontSize = StartButton.Font.Size;

            ResetButtonOriginSize = ResetButton.Size;
            ResetButtonOriginFontSize = ResetButton.Font.Size;

            StartImageOriginSize = StartImage.Size;
            ExecuteImageOriginSize = ExecuteImage.Size;
            ResultImageOriginSize = ResultImage.Size;

            ResultTextOriginSize = ResultText.Size;
            ResultTextOriginFontSize = ResultText.Font.Size;

            listBox1OriginSize = listBox1.Size;
            listBox1OriginFontSize = listBox1.Font.Size;

            // �p������P�������
            StartButtonPositionRatio = new PointF(
                (float)(StartButton.Location.X + StartButtonOriginSize.Width * 0.5) / LotteryOriginSize.Width,
                (float)(StartButton.Location.Y + StartButtonOriginSize.Height * 0.5) / LotteryOriginSize.Height
            );
            ResetButtonPositionRatio = new PointF(
                (float)(ResetButton.Location.X + ResetButtonOriginSize.Width * 0.5) / LotteryOriginSize.Width,
                (float)(ResetButton.Location.Y + ResetButtonOriginSize.Height * 0.5) / LotteryOriginSize.Height
            );
            StartImagePositionRatio = new PointF(
                (float)(StartImage.Location.X + StartImageOriginSize.Width * 0.5) / LotteryOriginSize.Width,
                (float)(StartImage.Location.Y + StartImageOriginSize.Height * 0.5) / LotteryOriginSize.Height
            );
            ExecuteImagePositionRatio = new PointF(
                (float)(ExecuteImage.Location.X + ExecuteImageOriginSize.Width * 0.5) / LotteryOriginSize.Width,
                (float)(ExecuteImage.Location.Y + ExecuteImageOriginSize.Height * 0.5) / LotteryOriginSize.Height
            );
            ResultImagePositionRatio = new PointF(
                (float)(ResultImage.Location.X + ResultImageOriginSize.Width * 0.5) / LotteryOriginSize.Width,
                (float)(ResultImage.Location.Y + ResultImageOriginSize.Height * 0.5) / LotteryOriginSize.Height
            );
            ResultTextPositionRatio = new PointF(
                (float)(ResultText.Location.X + ResultTextOriginSize.Width * 0.5) / LotteryOriginSize.Width,
                (float)(ResultText.Location.Y + ResultTextOriginSize.Height * 0.5) / LotteryOriginSize.Height
            );
            listBox1PositionRatio = new PointF(
                (float)(listBox1.Location.X + listBox1OriginSize.Width * 0.5) / LotteryOriginSize.Width,
                (float)(listBox1.Location.Y + listBox1OriginSize.Height * 0.5) / LotteryOriginSize.Height
            );

            // �]�w�����ϥ�
            this.Icon = Icon.FromHandle(new Bitmap("./icon/Lottery.png").GetHicon()); // �N JPG �ഫ�� Bitmap �ó]�m���ϥ�
            this.Resize += Form1_Resize;

            random = new Random();
            dataTable = new DataTable(); // ��l�� dataTable

            LoadExcelFile(file);
            DisplayDataInListBox(dataTable);
            StartImage.Image = Image.FromFile(Path.Combine(Application.StartupPath, "icon", "Lottery2.png"));

            // ��l�� Timer
            loadTimer = new System.Windows.Forms.Timer();
            loadTimer.Interval = 3000; // �]�w 3 ��
            loadTimer.Tick += LoadTimer_Tick; // �]�w�p�ɾ��ƥ�
        }
        private void LoadExcelFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                MessageBox.Show($"�ɮ� {filePath} ���s�b�I", "���~", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // �]�m LicenseContext ���D�ӷ~�γ~
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // Ū���Ĥ@�Ӥu�@��

                    // �M���{�����C
                    dataTable.Clear();
                    dataTable.Columns.Clear();

                    // Ū�����W��
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        dataTable.Columns.Add(worksheet.Cells[1, col].Text);
                    }

                    // Ū����ƦC
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        var newRow = dataTable.NewRow();
                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            newRow[col - 1] = worksheet.Cells[row, col].Text;
                        }
                        dataTable.Rows.Add(newRow);
                    }

                    // �T�{��ƬO�_���\���J
                    if (dataTable.Rows.Count <= 0)
                    {
                        MessageBox.Show("��Ƭ��šA���ˬd Excel �ɮפ��e�C", "ĵ�i", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    // Console��ܸ��
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        Console.Write($"{column.ColumnName}\t");
                    }
                    Console.WriteLine();
                    foreach (DataRow row in dataTable.Rows)
                    {
                        foreach (var item in row.ItemArray)
                        {
                            Console.Write($"{item}\t");
                        }
                        Console.WriteLine();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ū�� Excel �ɮ׮ɵo�Ϳ��~�G{ex.Message}", "���~", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void Start_Click(object sender, EventArgs e)
        {
            if (dataTable == null || dataTable.Rows.Count == 0)
            {
                MessageBox.Show("��Ƥw�g�Χ��Υ����J�I", "����", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ResultText.Text = string.Empty;
                ResultImage.Image = null; // �M�ŹϤ�
                return;
            }

            StartButton.Enabled = false; // �T��Start���s
            ResetButton.Enabled = false; // �T��Reset���s

            StartImage.Visible = false; // �}�Y�Ϥ�����
            ExecuteImage.Visible = true; // gif���
            ResultImage.Visible = false; // �Ϥ�����
            ResultText.Visible = false; // ��r����
            ExecuteImage.Image = Image.FromFile(Path.Combine(Application.StartupPath, "icon", "loading.gif"));
            ExecuteImage.SizeMode = PictureBoxSizeMode.StretchImage; // �]�w GIF ��ܤ覡

            // �p��[�v���v����
            Console.WriteLine("---------");
            List<double> cumulativeWeights = new List<double>();
            double totalWeight = 0;
            foreach (DataRow row in dataTable.Rows)
            {
                double weight = 1; // �w�]�v���� 1
                if (row["�v��"] != DBNull.Value && double.TryParse(row["�v��"].ToString(), out double probabilityAdjustment)) // �ˬd�O�_���v��
                {
                    if (probabilityAdjustment < 1) // �v���p�� 1 �ɡA�]�w�� 1
                    {
                        weight = 1;
                    }
                    else
                    {
                        weight = probabilityAdjustment;
                    }
                }
                totalWeight += weight;
                cumulativeWeights.Add(totalWeight); // �֭p�v��
                Console.WriteLine($"Weight: {weight}, TotalWeight: {totalWeight}");
            }
            Console.WriteLine("---------");

            // �H�����
            double randomValue = random.NextDouble() * totalWeight;
            Console.WriteLine($"RandomValue: {randomValue}");
            int selectedIndex = 0;
            for (int i = 0; i < cumulativeWeights.Count; i++)
            {
                if (randomValue < cumulativeWeights[i])
                {
                    selectedIndex = i;
                    break;
                }
            }
            DataRow selectedRow = dataTable.Rows[selectedIndex];

            // �O����ܪ����G�M�Ϥ��ɮצW��
            string selectedValue = selectedRow[0]?.ToString() ?? string.Empty;
            string filePath = GetImageFilePath(selectedValue);

            // �}�l�p�ɾ��A���� 3 �����ܵ��G
            loadTimer.Tag = new { RandomIndex = selectedIndex, SelectedValue = selectedValue, FilePath = filePath }; // �O�s�������
            loadTimer.Start();
        }
        private void LoadTimer_Tick(object? sender, EventArgs e)
        {
            // ����p�ɾ�
            loadTimer.Stop();

            // ���o�p�ɾ��������
            if (loadTimer.Tag is not null)
            {
                dynamic data = loadTimer.Tag;
                int randomIndex = data.RandomIndex;
                string selectedValue = data.SelectedValue;
                string filePath = data.FilePath;

                ExecuteImage.Visible = false; // ����
                ResultImage.Visible = true; // �Ϥ����
                ResultText.Visible = true; // ��r���

                // ��ܵ��G
                ResultText.Text = selectedValue;

                // ����Ū������ܹϤ�
                try
                {
                    ResultImage.Image = Image.FromFile(filePath);
                    ResultImage.SizeMode = PictureBoxSizeMode.StretchImage; // �վ���ܤ覡
                }
                catch (Exception)
                {
                    ResultImage.Image = null; // �p�G�����~�A�M���Ϥ�
                }

                // �R���Ӧ�
                dataTable.Rows.RemoveAt(randomIndex);
                // ���s��ܸ��
                DisplayDataInListBox(dataTable);

                StartButton.Enabled = true; // �ҥ�Start���s
                ResetButton.Enabled = true; // �ҥ�Reset���s
            }
        }
        private string GetImageFilePath(string baseFileName)
        {
            string[] possibleExtensions = { ".jpg", ".png", ".bmp", ".gif", ".jpeg" };
            foreach (string ext in possibleExtensions)
            {
                string filePath = Path.Combine(Application.StartupPath, "icon", baseFileName + ext);
                if (File.Exists(filePath))
                {
                    return filePath; // ����ɮרê�^���|
                }
            }
            return string.Empty; // �p�G�䤣��Ϥ��A��^�Ŧr��
        }
        private void Reset_Click(object sender, EventArgs e)
        {
            LoadExcelFile(file);
            DisplayDataInListBox(dataTable);
            StartImage.Visible = true; // �}�Y�Ϥ����
            ExecuteImage.Visible = false; // gif����
            ResultImage.Visible = false; // �Ϥ�����
            ResultText.Visible = false; // ��r����
            MessageBox.Show("�W��w���m�A���s����I", "����", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void DisplayDataInListBox(DataTable table)
        {
            // �M�� ListBox ���{�����e
            listBox1.Items.Clear();

            // �N��Ƹ��J�� ListBox
            foreach (DataRow row in table.Rows)
            {
                listBox1.Items.Add(row["�W��"]); // �o�̥H "Name" �����ܬ���
            }
        }


        private void Form1_Resize(object? sender, EventArgs e)
        {
            // ��������b��l�Ʈ�Ĳ�o�����n���վ�
            if (this.WindowState == FormWindowState.Minimized)
                return;

            // �����e�ؤo
            int currentWidth = this.Width;
            int currentHeight = this.Height;

            // �p��s���ؤo
            int newWidth = currentWidth;
            int newHeight = currentHeight;

            // �ˬd�O�e���٬O���ק��ܸ��j
            float widthChange = Math.Abs(currentWidth - previousSize.Width);
            float heightChange = Math.Abs(currentHeight - previousSize.Height);

            // �ھ��ܤƸ��j�����רӽվ�t�@�Ӻ���
            if (widthChange > heightChange)
            {
                // �e���ܤƧ�j�A�ڦ��վ㰪��
                newHeight = (int)(currentWidth / aspectRatio);
            }
            else
            {
                // �����ܤƧ�j�A�ڦ��վ�e��
                newWidth = (int)(currentHeight * aspectRatio);
            }

            // �]�m�s���j�p
            if (newWidth != currentWidth || newHeight != currentHeight)
            {
                this.Size = new Size(newWidth, newHeight);
            }

            // �p���Y����
            float widthRatio = (float)this.ClientSize.Width / LotteryOriginSize.Width;
            float heightRatio = (float)this.ClientSize.Height / LotteryOriginSize.Height;
            float scaleRatio = Math.Min(widthRatio, heightRatio);

            // ��s����j�p�M��m
            UpdateAllControls(scaleRatio);

            // �x�s��e�j�p�ѤU�����
            previousSize = this.Size;
        }
        private void UpdateAllControls(float scaleRatio)
        {
            // ��s���s
            UpdateButton(StartButton, StartButtonPositionRatio, StartButtonOriginSize, StartButtonOriginFontSize, scaleRatio);
            UpdateButton(ResetButton, ResetButtonPositionRatio, ResetButtonOriginSize, ResetButtonOriginFontSize, scaleRatio);

            // ��s�Ϥ�
            UpdatePictureBox(StartImage, StartImagePositionRatio, StartImageOriginSize, scaleRatio);
            UpdatePictureBox(ExecuteImage, ExecuteImagePositionRatio, ExecuteImageOriginSize, scaleRatio);
            UpdatePictureBox(ResultImage, ResultImagePositionRatio, ResultImageOriginSize, scaleRatio);

            // ��s��r��
            UpdateTextBox(ResultText, ResultTextPositionRatio, ResultTextOriginSize, ResultTextOriginFontSize, scaleRatio);

            // ��s�C���
            UpdateListBox(listBox1, listBox1PositionRatio, listBox1OriginSize, listBox1OriginFontSize, scaleRatio);
        }
        private void UpdateButton(Button button, PointF PositionRatio, Size originalSize, float originalFontSize, float scaleRatio)
            {
                // �վ�r��j�p
                float newFontSize = originalFontSize * scaleRatio;
                if (Math.Abs(button.Font.Size - newFontSize) > 0.1f)
                {
                    button.Font = new Font(button.Font.FontFamily, newFontSize);
                }
                // ���Ӥ�ҽվ�j�p
                int newWidth = (int)(originalSize.Width * scaleRatio);
                int newHeight = (int)(originalSize.Height * scaleRatio);
                button.Size = new Size(newWidth, newHeight);

                // �վ��m
                int x = (int)(this.ClientSize.Width * PositionRatio.X - button.Width * 0.5);
                int y = (int)(this.ClientSize.Height * PositionRatio.Y - button.Height * 0.5);
                button.Location = new Point(x, y);
            }
        private void UpdatePictureBox(PictureBox pictureBox, PointF PositionRatio, Size originalSize, float scaleRatio)
        {
            // ���Ӥ�ҽվ�j�p
            int newWidth = (int)(originalSize.Width * scaleRatio);
            int newHeight = (int)(originalSize.Height * scaleRatio);
            pictureBox.Size = new Size(newWidth, newHeight);

            // �վ��m
            int x = (int)(this.ClientSize.Width * PositionRatio.X - pictureBox.Width * 0.5);
            int y = (int)(this.ClientSize.Height * PositionRatio.Y - pictureBox.Height * 0.5);
            pictureBox.Location = new Point(x, y);
        }
        private void UpdateTextBox(TextBox textBox, PointF PositionRatio, Size originalSize, float originalFontSize, float scaleRatio)
        {
            // �վ�r��j�p
            float newFontSize = originalFontSize * scaleRatio;
            if (Math.Abs(textBox.Font.Size - newFontSize) > 0.1f)
            {
                textBox.Font = new Font(textBox.Font.FontFamily, newFontSize);
            }
            // ���Ӥ�ҽվ�j�p
            int newWidth = (int)(originalSize.Width * scaleRatio);
            int newHeight = (int)(originalSize.Height * scaleRatio);
            textBox.Size = new Size(newWidth, newHeight);

            // �վ��m
            int x = (int)(this.ClientSize.Width * PositionRatio.X - textBox.Width * 0.5);
            int y = (int)(this.ClientSize.Height * PositionRatio.Y - textBox.Height * 0.5);
            textBox.Location = new Point(x, y);
        }
        private void UpdateListBox(ListBox listBox, PointF PositionRatio, Size originalSize, float originalFontSize, float scaleRatio)
        {
            // �վ�r��j�p
            float newFontSize = originalFontSize * scaleRatio;
            if (Math.Abs(listBox.Font.Size - newFontSize) > 0.1f)
            {
                listBox.Font = new Font(listBox.Font.FontFamily, newFontSize);
            }
            // ���Ӥ�ҽվ�j�p
            int newWidth = (int)(originalSize.Width * scaleRatio);
            int newHeight = (int)(originalSize.Height * scaleRatio);
            listBox.Size = new Size(newWidth, newHeight);
            // �վ��m
            int x = (int)(this.ClientSize.Width * PositionRatio.X - listBox.Width * 0.5);
            int y = (int)(this.ClientSize.Height * PositionRatio.Y - listBox.Height * 0.5);
            listBox.Location = new Point(x, y);
        }
    }
}
