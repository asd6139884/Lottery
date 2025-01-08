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
            string selectedValue = selectedRow[0].ToString();
            string filePath = GetImageFilePath(selectedValue);

            // �}�l�p�ɾ��A���� 3 �����ܵ��G
            loadTimer.Tag = new { RandomIndex = randomIndex, SelectedValue = selectedValue, FilePath = filePath }; // �O�s�������
            loadTimer.Start();
        }
        private void LoadTimer_Tick(object sender, EventArgs e)
        {
            // ����p�ɾ�
            loadTimer.Stop();

            // ���o�p�ɾ��������
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
            catch (Exception ex)
            {
                ResultImage.Image = null; // �p�G�����~�A�M���Ϥ�
            }

            // �R���Ӧ�
            dataTable.Rows.RemoveAt(randomIndex);
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
        private void Result_TextChanged(object sender, EventArgs e)
        {

        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
        private void Form1_Resize(object sender, EventArgs e)
        {
            //�p��s�����ץH�������
            Size currentSize = this.Size;

            int newHeight = (int)(currentSize.Width / aspectRatio);
            int newWidth = (int)(currentSize.Height * aspectRatio);

            // �]�m�s���j�p
            Console.WriteLine("-------------------------");
            Console.WriteLine($"����Size : {currentSize.Width}x{currentSize.Height}");
            Console.WriteLine($"�e�@Size : {previousSize.Width}x{previousSize.Height}");

            if (currentSize.Width > previousSize.Width || currentSize.Height > previousSize.Height)
            {
                // �ϥΪ̦b��j
                Console.WriteLine("�ϥΪ̦b��j");
                newWidth = Math.Max(currentSize.Width, newWidth);
                newHeight = Math.Max(currentSize.Height, newHeight);
            }
            else if (currentSize.Width < previousSize.Width || currentSize.Height < previousSize.Height)
            {
                // �ϥΪ̦b�Y�p
                Console.WriteLine("�ϥΪ̦b�Y�p");
                newWidth = Math.Min(currentSize.Width, newWidth);
                newHeight = Math.Min(currentSize.Height, newHeight);
            }
            else
            {
                Console.WriteLine("Else");
                return;
            }

            Console.WriteLine($"�����x�sSize : {newWidth}x{newHeight}");
            this.Size = new Size(newWidth, newHeight); // �]�m�s���j�p
            //previousSize = new Size(newWidth, newHeight);

            // ���
            float widthRatio = (float)this.ClientSize.Width / LotteryOriginSize.Width;
            float heightRatio = (float)this.ClientSize.Height / LotteryOriginSize.Height;

            // ���e����������ҡ]�����Y��ĪG�^
            float scaleRatio = Math.Min(widthRatio, heightRatio);

            // Start���s
            Start.Width = (int)(StartButtonOriginSize.Width * widthRatio);
            Start.Height = (int)(StartButtonOriginSize.Height * heightRatio);
            Start.Location = new Point(
                (int)(this.ClientSize.Width * 0.5 - Start.Width * 0.5), // ������m
                (int)(this.ClientSize.Height * 0.9 - Start.Height * 0.5) // ������m
            );
            float StartFontSize = Math.Max(1, StartButtonOriginFontSize * scaleRatio); // �]�m�̤p�r��j�p�� 1
            Start.Font = new Font(Start.Font.FontFamily, StartFontSize); // ��s���s�r��j�p

            //Console.WriteLine($"Start���s�j�p: {Start.Size}, ��m: {Start.Location}, �r��j�p: {Start.Font.Size}");

            // Reset���s
            Reset.Width = (int)(ResetButtonOriginSize.Width * widthRatio);
            Reset.Height = (int)(ResetButtonOriginSize.Height * heightRatio);
            Reset.Location = new Point(
                (int)(this.ClientSize.Width * 0.85 - Reset.Width * 0.5),
                (int)(this.ClientSize.Height * 0.95 - Reset.Height * 0.5)
            );
            float ResetFontSize = Math.Max(1, ResetButtonOriginFontSize * scaleRatio); // �]�m�̤p�r��j�p�� 1
            Reset.Font = new Font(Reset.Font.FontFamily, ResetFontSize); // ��s���s�r��j�p

            //StartImage
            StartImage.Size = new Size(
                (int)(StartImageOriginSize.Width * widthRatio), 
                (int)(StartImageOriginSize.Height * heightRatio)
            );
            StartImage.Location = new Point(
                (int)(StartImageOriginLocation.X * widthRatio),
                (int)(StartImageOriginLocation.Y * heightRatio)
            );
            //ExecuteImage
            ExecuteImage.Size = new Size(
                (int)(ExecuteImageOriginSize.Width * widthRatio),
                (int)(ExecuteImageOriginSize.Height * heightRatio)
            );
            ExecuteImage.Location = new Point(
                (int)(ExecuteImageOriginLocation.X * widthRatio),
                (int)(ExecuteImageOriginLocation.Y * heightRatio)
            );
            //ResultImage
            ResultImage.Size = new Size(
                (int)(ResultImageOriginSize.Width * widthRatio),
                (int)(ResultImageOriginSize.Height * heightRatio)
            );
            ResultImage.Location = new Point(
                (int)(ResultImageOriginLocation.X * widthRatio),
                (int)(ResultImageOriginLocation.Y * heightRatio)
            );

            //ResultText
            ResultText.Location = new Point(
                (int)(this.ClientSize.Width * 0.5 - ResultText.Width * 0.5),
                (int)(this.ClientSize.Height * 0.1 - ResultText.Height * 0.5)
            );
            float ResultTextFontSize = Math.Max(1, ResultTextOriginFontSize * scaleRatio); // �]�m�̤p�r��j�p�� 1
            ResultText.Font = new Font(ResultText.Font.FontFamily, ResultTextFontSize); // ��s���s�r��j�p
        }
    }
}
