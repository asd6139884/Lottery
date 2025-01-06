using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml; // EPPlus 命名空間


namespace Lottery
{
    public partial class Lottery : Form
    {
        private string file = "Options.xlsx";
        private Random random; // 隨機數生成器
        private DataTable dataTable; // 用於存放資料

        public Lottery()
        {
            InitializeComponent();
            this.Icon = Icon.FromHandle(new Bitmap("./icon/Lottery.png").GetHicon()); // 將 JPG 轉換為 Bitmap 並設置為圖示
            random = new Random();
            dataTable = new DataTable(); // 初始化 dataTable

            LoadExcelFile(file);
        }

        private void LoadExcelFile(string filePath)
        {
            if (!File.Exists(file))
            {
                MessageBox.Show($"檔案 {file} 不存在！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // 設置 LicenseContext 為非商業用途
                using (var package = new ExcelPackage(new FileInfo(file)))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // 讀取第一個工作表

                    // 清除現有的列
                    dataTable.Clear();
                    dataTable.Columns.Clear();

                    // 讀取欄位名稱
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        dataTable.Columns.Add(worksheet.Cells[1, col].Text);
                    }

                    // 讀取資料列
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        var newRow = dataTable.NewRow();
                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            newRow[col - 1] = worksheet.Cells[row, col].Text;
                        }
                        dataTable.Rows.Add(newRow);
                    }

                    // 確認資料是否成功載入
                    if (dataTable.Rows.Count <= 0)
                    {
                        MessageBox.Show("資料為空，請檢查 Excel 檔案內容。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"讀取 Excel 檔案時發生錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Start_Click(object sender, EventArgs e)
        {
            string[] possibleExtensions = { ".jpg", ".png", ".bmp", ".gif", ".jpeg" };

            if (dataTable == null || dataTable.Rows.Count == 0)
            {
                MessageBox.Show("資料已經用完或未載入！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Result.Text = string.Empty;
                return;
            }

            // 隨機挑選一行
            int randomIndex = random.Next(dataTable.Rows.Count);
            DataRow selectedRow = dataTable.Rows[randomIndex];



            // 顯示結果
            Result.Text = selectedRow[0].ToString();

            // 讀取對應檔案名稱（圖片檔案）
            string filePath = string.Empty;
            foreach (string ext in possibleExtensions)
            {
                filePath = Path.Combine(Application.StartupPath, "icon", selectedRow[0].ToString() + ext);
                if (File.Exists(filePath))
                {
                    break; // 找到檔案後退出循環
                }
            }
            try
            {
                // 讀取圖片檔案並顯示在 PictureBox 中
                pictureBox1.Image = Image.FromFile(filePath);
                pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage; // 調整顯示方式
            }
            catch (Exception ex)
            {
                pictureBox1.Image = null; // 如果有錯誤，清除圖片
            }


            // 刪除該行
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
