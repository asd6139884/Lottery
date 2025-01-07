using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml; // EPPlus 命名空間


namespace Lottery
{
    public partial class Lottery : Form
    {
        private Random random; // 隨機數生成器
        private System.Windows.Forms.Timer loadTimer;
        private string file = "Options.xlsx";
        private DataTable dataTable; // 用於存放資料

        public Lottery()
        {
            InitializeComponent();
            this.Icon = Icon.FromHandle(new Bitmap("./icon/Lottery.png").GetHicon()); // 將 JPG 轉換為 Bitmap 並設置為圖示
            random = new Random();
            dataTable = new DataTable(); // 初始化 dataTable

            LoadExcelFile(file);
            pictureBox3.Image = Image.FromFile(Path.Combine(Application.StartupPath, "icon", "Lottery2.png"));

            // 初始化 Timer
            loadTimer = new System.Windows.Forms.Timer();
            loadTimer.Interval = 3000; // 設定 3 秒
            loadTimer.Tick += LoadTimer_Tick; // 設定計時器事件
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
            if (dataTable == null || dataTable.Rows.Count == 0)
            {
                MessageBox.Show("資料已經用完或未載入！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Result.Text = string.Empty;
                pictureBox1.Image = null; // 清空圖片
                return;
            }

            pictureBox3.Visible = false; // 開頭圖片隱藏
            pictureBox2.Visible = true; // gif顯示
            pictureBox1.Visible = false; // 圖片隱藏
            Result.Visible = false; // 文字隱藏
            pictureBox2.Image = Image.FromFile(Path.Combine(Application.StartupPath, "icon", "loading.gif"));
            pictureBox2.SizeMode = PictureBoxSizeMode.StretchImage; // 設定 GIF 顯示方式

            // 隨機挑選，但先不顯示結果
            int randomIndex = random.Next(dataTable.Rows.Count);
            DataRow selectedRow = dataTable.Rows[randomIndex];

            // 記錄選擇的結果和圖片檔案名稱
            string selectedValue = selectedRow[0].ToString();
            string filePath = GetImageFilePath(selectedValue);

            // 開始計時器，等待 3 秒後顯示結果
            loadTimer.Tag = new { RandomIndex = randomIndex, SelectedValue = selectedValue, FilePath = filePath }; // 保存相關資料
            loadTimer.Start();
        }

        private void LoadTimer_Tick(object sender, EventArgs e)
        {
            // 停止計時器
            loadTimer.Stop();

            // 取得計時器中的資料
            dynamic data = loadTimer.Tag;
            int randomIndex = data.RandomIndex;
            string selectedValue = data.SelectedValue;
            string filePath = data.FilePath;

            //pictureBox2.Image = null; // 隱藏動畫
            pictureBox2.Visible = false; // 隱藏
            pictureBox1.Visible = true; // 圖片顯示
            Result.Visible = true; // 文字顯示

            // 顯示結果
            Result.Text = selectedValue;

            // 嘗試讀取並顯示圖片
            try
            {
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
        private string GetImageFilePath(string baseFileName)
        {
            string[] possibleExtensions = { ".jpg", ".png", ".bmp", ".gif", ".jpeg" };
            foreach (string ext in possibleExtensions)
            {
                string filePath = Path.Combine(Application.StartupPath, "icon", baseFileName + ext);
                if (File.Exists(filePath))
                {
                    return filePath; // 找到檔案並返回路徑
                }
            }
            return string.Empty; // 如果找不到圖片，返回空字串
        }



        private void Reset_Click(object sender, EventArgs e)
        {
            LoadExcelFile(file);
            pictureBox3.Visible = true; // 開頭圖片顯示
            pictureBox2.Visible = false; // gif隱藏
            pictureBox1.Visible = false; // 圖片隱藏
            Result.Visible = false; // 文字隱藏
            MessageBox.Show("名單已重制，重新抽獎！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void Result_TextChanged(object sender, EventArgs e)
        {

        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
