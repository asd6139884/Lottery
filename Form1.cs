using System;
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
        private Size StartButtonOriginSize; // Start���s��l�j�p
        private float StartButtonOriginFontSize; // Start���s��l�r��j�p
        private Size ResetButtonOriginSize; //  Reset���s��l�j�p
        private float ResetButtonOriginFontSize; // Reset���s��l�r��j�p
        private Size StartImageOriginSize;  // StartImage��l�j�p
        private Point StartImageOriginLocation; // StartImage��l��m
        private Size ExecuteImageOriginSize;  // ExecuteImage��l�j�p
        private Point ExecuteImageOriginLocation; // ExecuteImage��l��m
        private Size ResultImageOriginSize;  // ResultImage��l�j�p
        private Point ResultImageOriginLocation; // ResultImage��l��m
        private Point ResultTextOriginLocation; // ResultText��l��m
        private float ResultTextOriginFontSize; // ResultText��l�r��j�p

        public Lottery()
        {
            InitializeComponent();

            // ���o�����T
            LotteryOriginSize = this.Size;
            previousSize = this.Size;
            aspectRatio = (float)LotteryOriginSize.Width / LotteryOriginSize.Height;
            StartButtonOriginSize = Start.Size;
            StartButtonOriginFontSize = Start.Font.Size;
            ResetButtonOriginSize = Reset.Size;
            ResetButtonOriginFontSize = Reset.Font.Size;
            StartImageOriginSize = StartImage.Size;
            StartImageOriginLocation = StartImage.Location;
            ExecuteImageOriginSize = ExecuteImage.Size;
            ExecuteImageOriginLocation = ExecuteImage.Location;
            ResultImageOriginSize = ResultImage.Size;
            ResultImageOriginLocation = ResultImage.Location;
            ResultTextOriginLocation = ResultText.Location;
            ResultTextOriginFontSize = ResultText.Font.Size;

            this.Icon = Icon.FromHandle(new Bitmap("./icon/Lottery.png").GetHicon()); // �N JPG �ഫ�� Bitmap �ó]�m���ϥ�
            this.Resize += Form1_Resize;

            random = new Random();
            dataTable = new DataTable(); // ��l�� dataTable

            LoadExcelFile(file);
            StartImage.Image = Image.FromFile(Path.Combine(Application.StartupPath, "icon", "Lottery2.png"));

            // ��l�� Timer
            loadTimer = new System.Windows.Forms.Timer();
            loadTimer.Interval = 3000; // �]�w 3 ��
            loadTimer.Tick += LoadTimer_Tick; // �]�w�p�ɾ��ƥ�

        }

        private void LoadExcelFile(string filePath)
        {
            if (!File.Exists(file))
            {
                MessageBox.Show($"�ɮ� {file} ���s�b�I", "���~", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // �]�m LicenseContext ���D�ӷ~�γ~
                using (var package = new ExcelPackage(new FileInfo(file)))
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

            Start.Enabled = false; // �T��Start���s
            Reset.Enabled = false; // �T��Reset���s

            StartImage.Visible = false; // �}�Y�Ϥ�����
            ExecuteImage.Visible = true; // gif���
            ResultImage.Visible = false; // �Ϥ�����
            ResultText.Visible = false; // ��r����
            ExecuteImage.Image = Image.FromFile(Path.Combine(Application.StartupPath, "icon", "loading.gif"));
            ExecuteImage.SizeMode = PictureBoxSizeMode.StretchImage; // �]�w GIF ��ܤ覡

            // �H���D��A��������ܵ��G
            int randomIndex = random.Next(dataTable.Rows.Count);
            DataRow selectedRow = dataTable.Rows[randomIndex];

            // �O����ܪ����G�M�Ϥ��ɮצW��
            string selectedValue = selectedRow[0]?.ToString() ?? string.Empty;
            string filePath = GetImageFilePath(selectedValue);

            // �}�l�p�ɾ��A���� 3 �����ܵ��G
            loadTimer.Tag = new { RandomIndex = randomIndex, SelectedValue = selectedValue, FilePath = filePath }; // �O�s�������
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

                //pictureBox2.Image = null; // ���ðʵe
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

                Start.Enabled = true; // �ҥ�Start���s
                Reset.Enabled = true; // �ҥ�Reset���s
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
            StartImage.Visible = true; // �}�Y�Ϥ����
            ExecuteImage.Visible = false; // gif����
            ResultImage.Visible = false; // �Ϥ�����
            ResultText.Visible = false; // ��r����
            MessageBox.Show("�W��w����A���s����I", "����", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            UpdateButton(Start, StartButtonOriginSize, StartButtonOriginFontSize, scaleRatio, 0.5f, 0.9f);
            UpdateButton(Reset, ResetButtonOriginSize, ResetButtonOriginFontSize, scaleRatio, 0.85f, 0.95f);

            // ��s�Ϥ�
            UpdatePictureBox(StartImage, StartImageOriginSize, StartImageOriginLocation, scaleRatio, 0.5f, 0.5f);
            UpdatePictureBox(ExecuteImage, ExecuteImageOriginSize, ExecuteImageOriginLocation, scaleRatio, 0.5f, 0.5f);
            UpdatePictureBox(ResultImage, ResultImageOriginSize, ResultImageOriginLocation, scaleRatio, 0.5f, 0.5f);

            // ��s��r��
            UpdateTextBox(ResultText, ResultTextOriginFontSize, scaleRatio, 0.5f, 0.15f);
        }

        private void UpdateButton(Button button, Size originalSize, float originalFontSize,
    float scaleRatio, float horizontalPosition, float verticalPosition)
        {
            // ���Ӥ�ҽվ�j�p
            int newWidth = (int)(originalSize.Width * scaleRatio);
            int newHeight = (int)(originalSize.Height * scaleRatio);
            button.Size = new Size(newWidth, newHeight);

            // �վ��m
            int x = (int)(this.ClientSize.Width * horizontalPosition - button.Width * 0.5);
            int y = (int)(this.ClientSize.Height * verticalPosition - button.Height * 0.5);
            button.Location = new Point(x, y);

            // �վ�r��j�p
            float newFontSize = originalFontSize * scaleRatio;
            if (Math.Abs(button.Font.Size - newFontSize) > 0.1f)
            {
                button.Font = new Font(button.Font.FontFamily, newFontSize);
            }
        }

        private void UpdatePictureBox(PictureBox pictureBox, Size originalSize, Point originalLocation, float scaleRatio,
            float horizontalPosition, float verticalPosition)
        {
            // ���Ӥ�ҽվ�j�p
            int newWidth = (int)(originalSize.Width * scaleRatio);
            int newHeight = (int)(originalSize.Height * scaleRatio);
            pictureBox.Size = new Size(newWidth, newHeight);

            // �վ��m
            int x = (int)(this.ClientSize.Width * horizontalPosition - pictureBox.Width * 0.5);
            int y = (int)(this.ClientSize.Height * verticalPosition - pictureBox.Height * 0.5);
            pictureBox.Location = new Point(x, y);
        }

        private void UpdateTextBox(TextBox textBox, float originalFontSize, float scaleRatio,
            float horizontalPosition, float verticalPosition)
        {
            // �վ�r��j�p
            float newFontSize = originalFontSize * scaleRatio;
            if (Math.Abs(textBox.Font.Size - newFontSize) > 0.1f)
            {
                textBox.Font = new Font(textBox.Font.FontFamily, newFontSize);
            }

            // �վ��m
            int x = (int)(this.ClientSize.Width * horizontalPosition - textBox.Width * 0.5);
            int y = (int)(this.ClientSize.Height * verticalPosition - textBox.Height * 0.5);
            textBox.Location = new Point(x, y);
        }
    }
}
