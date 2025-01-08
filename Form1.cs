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
            catch (Exception ex)
            {
                ResultImage.Image = null; // 如果有錯誤，清除圖片
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
            StartImage.Visible = true; // 開頭圖片顯示
            ExecuteImage.Visible = false; // gif隱藏
            ResultImage.Visible = false; // 圖片隱藏
            ResultText.Visible = false; // 文字隱藏
            MessageBox.Show("名單已重制，重新抽獎！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void Result_TextChanged(object sender, EventArgs e)
        {

        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
        private void Form1_Resize(object sender, EventArgs e)
        {
            //計算新的高度以維持比例
            Size currentSize = this.Size;

            int newHeight = (int)(currentSize.Width / aspectRatio);
            int newWidth = (int)(currentSize.Height * aspectRatio);

            // 設置新的大小
            Console.WriteLine("-------------------------");
            Console.WriteLine($"視窗Size : {currentSize.Width}x{currentSize.Height}");
            Console.WriteLine($"前一Size : {previousSize.Width}x{previousSize.Height}");

            if (currentSize.Width > previousSize.Width || currentSize.Height > previousSize.Height)
            {
                // 使用者在放大
                Console.WriteLine("使用者在放大");
                newWidth = Math.Max(currentSize.Width, newWidth);
                newHeight = Math.Max(currentSize.Height, newHeight);
            }
            else if (currentSize.Width < previousSize.Width || currentSize.Height < previousSize.Height)
            {
                // 使用者在縮小
                Console.WriteLine("使用者在縮小");
                newWidth = Math.Min(currentSize.Width, newWidth);
                newHeight = Math.Min(currentSize.Height, newHeight);
            }
            else
            {
                Console.WriteLine("Else");
                return;
            }

            Console.WriteLine($"視窗儲存Size : {newWidth}x{newHeight}");
            this.Size = new Size(newWidth, newHeight); // 設置新的大小
            //previousSize = new Size(newWidth, newHeight);

            // 比例
            float widthRatio = (float)this.ClientSize.Width / LotteryOriginSize.Width;
            float heightRatio = (float)this.ClientSize.Height / LotteryOriginSize.Height;

            // 取寬高的平均比例（平衡縮放效果）
            float scaleRatio = Math.Min(widthRatio, heightRatio);

            // Start按鈕
            Start.Width = (int)(StartButtonOriginSize.Width * widthRatio);
            Start.Height = (int)(StartButtonOriginSize.Height * heightRatio);
            Start.Location = new Point(
                (int)(this.ClientSize.Width * 0.5 - Start.Width * 0.5), // 水平位置
                (int)(this.ClientSize.Height * 0.9 - Start.Height * 0.5) // 垂直位置
            );
            float StartFontSize = Math.Max(1, StartButtonOriginFontSize * scaleRatio); // 設置最小字體大小為 1
            Start.Font = new Font(Start.Font.FontFamily, StartFontSize); // 更新按鈕字體大小

            //Console.WriteLine($"Start按鈕大小: {Start.Size}, 位置: {Start.Location}, 字體大小: {Start.Font.Size}");

            // Reset按鈕
            Reset.Width = (int)(ResetButtonOriginSize.Width * widthRatio);
            Reset.Height = (int)(ResetButtonOriginSize.Height * heightRatio);
            Reset.Location = new Point(
                (int)(this.ClientSize.Width * 0.85 - Reset.Width * 0.5),
                (int)(this.ClientSize.Height * 0.95 - Reset.Height * 0.5)
            );
            float ResetFontSize = Math.Max(1, ResetButtonOriginFontSize * scaleRatio); // 設置最小字體大小為 1
            Reset.Font = new Font(Reset.Font.FontFamily, ResetFontSize); // 更新按鈕字體大小

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
            float ResultTextFontSize = Math.Max(1, ResultTextOriginFontSize * scaleRatio); // 設置最小字體大小為 1
            ResultText.Font = new Font(ResultText.Font.FontFamily, ResultTextFontSize); // 更新按鈕字體大小
        }
    }
}
