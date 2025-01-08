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

        //物件資訊
        private Size LotteryOriginSize; // 視窗初始大小
        private Size previousSize; // 視窗大小
        private float aspectRatio; // 寬高比
        private Size StartButtonOriginSize; // Start按鈕初始大小
        private float StartButtonOriginFontSize; // Start按鈕初始字體大小
        private Size ResetButtonOriginSize; //  Reset按鈕初始大小
        private float ResetButtonOriginFontSize; // Reset按鈕初始字體大小
        private Size StartImageOriginSize;  // StartImage初始大小
        private Point StartImageOriginLocation; // StartImage初始位置
        private Size ExecuteImageOriginSize;  // ExecuteImage初始大小
        private Point ExecuteImageOriginLocation; // ExecuteImage初始位置
        private Size ResultImageOriginSize;  // ResultImage初始大小
        private Point ResultImageOriginLocation; // ResultImage初始位置
        private Point ResultTextOriginLocation; // ResultText初始位置
        private float ResultTextOriginFontSize; // ResultText初始字體大小

        public Lottery()
        {
            InitializeComponent();

            // 取得物件資訊
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

            this.Icon = Icon.FromHandle(new Bitmap("./icon/Lottery.png").GetHicon()); // 將 JPG 轉換為 Bitmap 並設置為圖示
            this.Resize += Form1_Resize;

            random = new Random();
            dataTable = new DataTable(); // 初始化 dataTable

            LoadExcelFile(file);
            StartImage.Image = Image.FromFile(Path.Combine(Application.StartupPath, "icon", "Lottery2.png"));

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
                ResultText.Text = string.Empty;
                ResultImage.Image = null; // 清空圖片
                return;
            }

            Start.Enabled = false; // 禁用Start按鈕
            Reset.Enabled = false; // 禁用Reset按鈕

            StartImage.Visible = false; // 開頭圖片隱藏
            ExecuteImage.Visible = true; // gif顯示
            ResultImage.Visible = false; // 圖片隱藏
            ResultText.Visible = false; // 文字隱藏
            ExecuteImage.Image = Image.FromFile(Path.Combine(Application.StartupPath, "icon", "loading.gif"));
            ExecuteImage.SizeMode = PictureBoxSizeMode.StretchImage; // 設定 GIF 顯示方式

            // 隨機挑選，但先不顯示結果
            int randomIndex = random.Next(dataTable.Rows.Count);
            DataRow selectedRow = dataTable.Rows[randomIndex];

            // 記錄選擇的結果和圖片檔案名稱
            string selectedValue = selectedRow[0]?.ToString() ?? string.Empty;
            string filePath = GetImageFilePath(selectedValue);

            // 開始計時器，等待 3 秒後顯示結果
            loadTimer.Tag = new { RandomIndex = randomIndex, SelectedValue = selectedValue, FilePath = filePath }; // 保存相關資料
            loadTimer.Start();
        }
        private void LoadTimer_Tick(object? sender, EventArgs e)
        {
            // 停止計時器
            loadTimer.Stop();

            // 取得計時器中的資料
            if (loadTimer.Tag is not null)
            {
                dynamic data = loadTimer.Tag;
                int randomIndex = data.RandomIndex;
                string selectedValue = data.SelectedValue;
                string filePath = data.FilePath;

                //pictureBox2.Image = null; // 隱藏動畫
                ExecuteImage.Visible = false; // 隱藏
                ResultImage.Visible = true; // 圖片顯示
                ResultText.Visible = true; // 文字顯示

                // 顯示結果
                ResultText.Text = selectedValue;

                // 嘗試讀取並顯示圖片
                try
                {
                    ResultImage.Image = Image.FromFile(filePath);
                    ResultImage.SizeMode = PictureBoxSizeMode.StretchImage; // 調整顯示方式
                }
                catch (Exception)
                {
                    ResultImage.Image = null; // 如果有錯誤，清除圖片
                }

                // 刪除該行
                dataTable.Rows.RemoveAt(randomIndex);

                Start.Enabled = true; // 啟用Start按鈕
                Reset.Enabled = true; // 啟用Reset按鈕
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
                    return filePath; // 找到檔案並返回路徑
                }
            }
            return string.Empty; // 如果找不到圖片，返回空字串
        }
        private void Reset_Click(object sender, EventArgs e)
        {
            LoadExcelFile(file);
            StartImage.Visible = true; // 開頭圖片顯示
            ExecuteImage.Visible = false; // gif隱藏
            ResultImage.Visible = false; // 圖片隱藏
            ResultText.Visible = false; // 文字隱藏
            MessageBox.Show("名單已重制，重新抽獎！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void Form1_Resize(object? sender, EventArgs e)
        {
            // 防止視窗在初始化時觸發不必要的調整
            if (this.WindowState == FormWindowState.Minimized)
                return;

            // 獲取當前尺寸
            int currentWidth = this.Width;
            int currentHeight = this.Height;

            // 計算新的尺寸
            int newWidth = currentWidth;
            int newHeight = currentHeight;

            // 檢查是寬度還是高度改變較大
            float widthChange = Math.Abs(currentWidth - previousSize.Width);
            float heightChange = Math.Abs(currentHeight - previousSize.Height);

            // 根據變化較大的維度來調整另一個維度
            if (widthChange > heightChange)
            {
                // 寬度變化更大，據此調整高度
                newHeight = (int)(currentWidth / aspectRatio);
            }
            else
            {
                // 高度變化更大，據此調整寬度
                newWidth = (int)(currentHeight * aspectRatio);
            }

            // 設置新的大小
            if (newWidth != currentWidth || newHeight != currentHeight)
            {
                this.Size = new Size(newWidth, newHeight);
            }

            // 計算縮放比例
            float widthRatio = (float)this.ClientSize.Width / LotteryOriginSize.Width;
            float heightRatio = (float)this.ClientSize.Height / LotteryOriginSize.Height;
            float scaleRatio = Math.Min(widthRatio, heightRatio);

            // 更新控制項大小和位置
            UpdateAllControls(scaleRatio);

            // 儲存當前大小供下次比較
            previousSize = this.Size;
        }
        private void UpdateAllControls(float scaleRatio)
        {
            // 更新按鈕
            UpdateButton(Start, StartButtonOriginSize, StartButtonOriginFontSize, scaleRatio, 0.5f, 0.9f);
            UpdateButton(Reset, ResetButtonOriginSize, ResetButtonOriginFontSize, scaleRatio, 0.85f, 0.95f);

            // 更新圖片
            UpdatePictureBox(StartImage, StartImageOriginSize, StartImageOriginLocation, scaleRatio, 0.5f, 0.5f);
            UpdatePictureBox(ExecuteImage, ExecuteImageOriginSize, ExecuteImageOriginLocation, scaleRatio, 0.5f, 0.5f);
            UpdatePictureBox(ResultImage, ResultImageOriginSize, ResultImageOriginLocation, scaleRatio, 0.5f, 0.5f);

            // 更新文字框
            UpdateTextBox(ResultText, ResultTextOriginFontSize, scaleRatio, 0.5f, 0.15f);
        }

        private void UpdateButton(Button button, Size originalSize, float originalFontSize,
    float scaleRatio, float horizontalPosition, float verticalPosition)
        {
            // 按照比例調整大小
            int newWidth = (int)(originalSize.Width * scaleRatio);
            int newHeight = (int)(originalSize.Height * scaleRatio);
            button.Size = new Size(newWidth, newHeight);

            // 調整位置
            int x = (int)(this.ClientSize.Width * horizontalPosition - button.Width * 0.5);
            int y = (int)(this.ClientSize.Height * verticalPosition - button.Height * 0.5);
            button.Location = new Point(x, y);

            // 調整字體大小
            float newFontSize = originalFontSize * scaleRatio;
            if (Math.Abs(button.Font.Size - newFontSize) > 0.1f)
            {
                button.Font = new Font(button.Font.FontFamily, newFontSize);
            }
        }

        private void UpdatePictureBox(PictureBox pictureBox, Size originalSize, Point originalLocation, float scaleRatio,
            float horizontalPosition, float verticalPosition)
        {
            // 按照比例調整大小
            int newWidth = (int)(originalSize.Width * scaleRatio);
            int newHeight = (int)(originalSize.Height * scaleRatio);
            pictureBox.Size = new Size(newWidth, newHeight);

            // 調整位置
            int x = (int)(this.ClientSize.Width * horizontalPosition - pictureBox.Width * 0.5);
            int y = (int)(this.ClientSize.Height * verticalPosition - pictureBox.Height * 0.5);
            pictureBox.Location = new Point(x, y);
        }

        private void UpdateTextBox(TextBox textBox, float originalFontSize, float scaleRatio,
            float horizontalPosition, float verticalPosition)
        {
            // 調整字體大小
            float newFontSize = originalFontSize * scaleRatio;
            if (Math.Abs(textBox.Font.Size - newFontSize) > 0.1f)
            {
                textBox.Font = new Font(textBox.Font.FontFamily, newFontSize);
            }

            // 調整位置
            int x = (int)(this.ClientSize.Width * horizontalPosition - textBox.Width * 0.5);
            int y = (int)(this.ClientSize.Height * verticalPosition - textBox.Height * 0.5);
            textBox.Location = new Point(x, y);
        }
    }
}
