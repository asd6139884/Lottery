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

        public Lottery()
        {
            InitializeComponent();
            this.Icon = Icon.FromHandle(new Bitmap("./icon/Lottery.png").GetHicon()); // �N JPG �ഫ�� Bitmap �ó]�m���ϥ�
            random = new Random();
            dataTable = new DataTable(); // ��l�� dataTable

            LoadExcelFile(file);
            pictureBox3.Image = Image.FromFile(Path.Combine(Application.StartupPath, "icon", "Lottery2.png"));

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
                Result.Text = string.Empty;
                pictureBox1.Image = null; // �M�ŹϤ�
                return;
            }

            pictureBox3.Visible = false; // �}�Y�Ϥ�����
            pictureBox2.Visible = true; // gif���
            pictureBox1.Visible = false; // �Ϥ�����
            Result.Visible = false; // ��r����
            pictureBox2.Image = Image.FromFile(Path.Combine(Application.StartupPath, "icon", "loading.gif"));
            pictureBox2.SizeMode = PictureBoxSizeMode.StretchImage; // �]�w GIF ��ܤ覡

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
            pictureBox2.Visible = false; // ����
            pictureBox1.Visible = true; // �Ϥ����
            Result.Visible = true; // ��r���

            // ��ܵ��G
            Result.Text = selectedValue;

            // ����Ū������ܹϤ�
            try
            {
                pictureBox1.Image = Image.FromFile(filePath);
                pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage; // �վ���ܤ覡
            }
            catch (Exception ex)
            {
                pictureBox1.Image = null; // �p�G�����~�A�M���Ϥ�
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
            pictureBox3.Visible = true; // �}�Y�Ϥ����
            pictureBox2.Visible = false; // gif����
            pictureBox1.Visible = false; // �Ϥ�����
            Result.Visible = false; // ��r����
            MessageBox.Show("�W��w����A���s����I", "����", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void Result_TextChanged(object sender, EventArgs e)
        {

        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
