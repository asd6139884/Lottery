using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml; // EPPlus �R�W�Ŷ�


namespace Lottery
{
    public partial class Lottery : Form
    {
        private string file = "Options.xlsx";
        private Random random; // �H���ƥͦ���
        private DataTable dataTable; // �Ω�s����

        public Lottery()
        {
            InitializeComponent();
            this.Icon = Icon.FromHandle(new Bitmap("./icon/Lottery.png").GetHicon()); // �N JPG �ഫ�� Bitmap �ó]�m���ϥ�
            random = new Random();
            dataTable = new DataTable(); // ��l�� dataTable

            LoadExcelFile(file);
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
            string[] possibleExtensions = { ".jpg", ".png", ".bmp", ".gif", ".jpeg" };

            if (dataTable == null || dataTable.Rows.Count == 0)
            {
                MessageBox.Show("��Ƥw�g�Χ��Υ����J�I", "����", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Result.Text = string.Empty;
                return;
            }

            // �H���D��@��
            int randomIndex = random.Next(dataTable.Rows.Count);
            DataRow selectedRow = dataTable.Rows[randomIndex];



            // ��ܵ��G
            Result.Text = selectedRow[0].ToString();

            // Ū�������ɮצW�١]�Ϥ��ɮס^
            string filePath = string.Empty;
            foreach (string ext in possibleExtensions)
            {
                filePath = Path.Combine(Application.StartupPath, "icon", selectedRow[0].ToString() + ext);
                if (File.Exists(filePath))
                {
                    break; // ����ɮ׫�h�X�`��
                }
            }
            try
            {
                // Ū���Ϥ��ɮר���ܦb PictureBox ��
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
        private void Reset_Click(object sender, EventArgs e)
        {
            LoadExcelFile(file);
        }
        private void Result_TextChanged(object sender, EventArgs e)
        {

        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
