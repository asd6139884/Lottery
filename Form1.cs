using System;
using System.Collections;
using System.Collections.Generic;
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
        // Start按鈕
        private Size StartButtonOriginSize; // 初始大小
        private float StartButtonOriginFontSize; // 初始字體大小
        private readonly PointF StartButtonPositionRatio; // 位置比例
        //  Reset按鈕
        private Size ResetButtonOriginSize; //  初始大小
        private float ResetButtonOriginFontSize; // 初始字體大小
        private readonly PointF ResetButtonPositionRatio; // 位置比例
        // StartImage
        private Size StartImageOriginSize;  // 初始大小
        private readonly PointF StartImagePositionRatio; // 位置比例
        // ExecuteImage
        private Size ExecuteImageOriginSize;  // 初始大小
        private readonly PointF ExecuteImagePositionRatio; // 位置比例
        // ResultImage
        private Size ResultImageOriginSize;  // 初始大小
        private readonly PointF ResultImagePositionRatio; // 位置比例
        // ResultText
        private Size ResultTextOriginSize; //初始大小
        private float ResultTextOriginFontSize; // 初始字體大小
        private readonly PointF ResultTextPositionRatio; // 位置比例
        // listBox1
        private Size listBox1OriginSize; // 初始大小
        private float listBox1OriginFontSize; // 初始字體大小
        private readonly PointF listBox1PositionRatio; // 位置比例

        public Lottery()
        {
            InitializeComponent();

            // 取得物件資訊
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

            // 計算水平與垂直比例
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

            // 設定視窗圖示
            this.Icon = Icon.FromHandle(new Bitmap("./icon/Lottery.png").GetHicon()); // 將 JPG 轉換為 Bitmap 並設置為圖示
            this.Resize += Form1_Resize;

            random = new Random();
            dataTable = new DataTable(); // 初始化 dataTable

            LoadExcelFile(file);
            DisplayDataInListBox(dataTable);
            StartImage.Image = Image.FromFile(Path.Combine(Application.StartupPath, "icon", "Lottery2.png"));

            // 初始化 Timer
            loadTimer = new System.Windows.Forms.Timer();
            loadTimer.Interval = 3000; // 設定 3 秒
            loadTimer.Tick += LoadTimer_Tick; // 設定計時器事件
        }
        private void LoadExcelFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                MessageBox.Show($"檔案 {filePath} 不存在！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // 設置 LicenseContext 為非商業用途
                using (var package = new ExcelPackage(new FileInfo(filePath)))
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

                    // Console顯示資料
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

            StartButton.Enabled = false; // 禁用Start按鈕
            ResetButton.Enabled = false; // 禁用Reset按鈕

            StartImage.Visible = false; // 開頭圖片隱藏
            ExecuteImage.Visible = true; // gif顯示
            ResultImage.Visible = false; // 圖片隱藏
            ResultText.Visible = false; // 文字隱藏
            ExecuteImage.Image = Image.FromFile(Path.Combine(Application.StartupPath, "icon", "loading.gif"));
            ExecuteImage.SizeMode = PictureBoxSizeMode.StretchImage; // 設定 GIF 顯示方式

            // 計算加權機率分布
            Console.WriteLine("---------");
            List<double> cumulativeWeights = new List<double>();
            double totalWeight = 0;
            foreach (DataRow row in dataTable.Rows)
            {
                double weight = 1; // 預設權重為 1
                if (row["權重"] != DBNull.Value && double.TryParse(row["權重"].ToString(), out double probabilityAdjustment)) // 檢查是否有權重
                {
                    if (probabilityAdjustment < 1) // 權重小於 1 時，設定為 1
                    {
                        weight = 1;
                    }
                    else
                    {
                        weight = probabilityAdjustment;
                    }
                }
                totalWeight += weight;
                cumulativeWeights.Add(totalWeight); // 累計權重
                Console.WriteLine($"Weight: {weight}, TotalWeight: {totalWeight}");
            }
            Console.WriteLine("---------");

            // 隨機選擇
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

            // 記錄選擇的結果和圖片檔案名稱
            string selectedValue = selectedRow[0]?.ToString() ?? string.Empty;
            string filePath = GetImageFilePath(selectedValue);

            // 開始計時器，等待 3 秒後顯示結果
            loadTimer.Tag = new { RandomIndex = selectedIndex, SelectedValue = selectedValue, FilePath = filePath }; // 保存相關資料
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
                // 重新顯示資料
                DisplayDataInListBox(dataTable);

                StartButton.Enabled = true; // 啟用Start按鈕
                ResetButton.Enabled = true; // 啟用Reset按鈕
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
            DisplayDataInListBox(dataTable);
            StartImage.Visible = true; // 開頭圖片顯示
            ExecuteImage.Visible = false; // gif隱藏
            ResultImage.Visible = false; // 圖片隱藏
            ResultText.Visible = false; // 文字隱藏
            MessageBox.Show("名單已重置，重新抽獎！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void DisplayDataInListBox(DataTable table)
        {
            // 清空 ListBox 的現有內容
            listBox1.Items.Clear();

            // 將資料載入到 ListBox
            foreach (DataRow row in table.Rows)
            {
                listBox1.Items.Add(row["名稱"]); // 這裡以 "Name" 欄位顯示為例
            }
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
            UpdateButton(StartButton, StartButtonPositionRatio, StartButtonOriginSize, StartButtonOriginFontSize, scaleRatio);
            UpdateButton(ResetButton, ResetButtonPositionRatio, ResetButtonOriginSize, ResetButtonOriginFontSize, scaleRatio);

            // 更新圖片
            UpdatePictureBox(StartImage, StartImagePositionRatio, StartImageOriginSize, scaleRatio);
            UpdatePictureBox(ExecuteImage, ExecuteImagePositionRatio, ExecuteImageOriginSize, scaleRatio);
            UpdatePictureBox(ResultImage, ResultImagePositionRatio, ResultImageOriginSize, scaleRatio);

            // 更新文字框
            UpdateTextBox(ResultText, ResultTextPositionRatio, ResultTextOriginSize, ResultTextOriginFontSize, scaleRatio);

            // 更新列表框
            UpdateListBox(listBox1, listBox1PositionRatio, listBox1OriginSize, listBox1OriginFontSize, scaleRatio);
        }
        private void UpdateButton(Button button, PointF PositionRatio, Size originalSize, float originalFontSize, float scaleRatio)
            {
                // 調整字體大小
                float newFontSize = originalFontSize * scaleRatio;
                if (Math.Abs(button.Font.Size - newFontSize) > 0.1f)
                {
                    button.Font = new Font(button.Font.FontFamily, newFontSize);
                }
                // 按照比例調整大小
                int newWidth = (int)(originalSize.Width * scaleRatio);
                int newHeight = (int)(originalSize.Height * scaleRatio);
                button.Size = new Size(newWidth, newHeight);

                // 調整位置
                int x = (int)(this.ClientSize.Width * PositionRatio.X - button.Width * 0.5);
                int y = (int)(this.ClientSize.Height * PositionRatio.Y - button.Height * 0.5);
                button.Location = new Point(x, y);
            }
        private void UpdatePictureBox(PictureBox pictureBox, PointF PositionRatio, Size originalSize, float scaleRatio)
        {
            // 按照比例調整大小
            int newWidth = (int)(originalSize.Width * scaleRatio);
            int newHeight = (int)(originalSize.Height * scaleRatio);
            pictureBox.Size = new Size(newWidth, newHeight);

            // 調整位置
            int x = (int)(this.ClientSize.Width * PositionRatio.X - pictureBox.Width * 0.5);
            int y = (int)(this.ClientSize.Height * PositionRatio.Y - pictureBox.Height * 0.5);
            pictureBox.Location = new Point(x, y);
        }
        private void UpdateTextBox(TextBox textBox, PointF PositionRatio, Size originalSize, float originalFontSize, float scaleRatio)
        {
            // 調整字體大小
            float newFontSize = originalFontSize * scaleRatio;
            if (Math.Abs(textBox.Font.Size - newFontSize) > 0.1f)
            {
                textBox.Font = new Font(textBox.Font.FontFamily, newFontSize);
            }
            // 按照比例調整大小
            int newWidth = (int)(originalSize.Width * scaleRatio);
            int newHeight = (int)(originalSize.Height * scaleRatio);
            textBox.Size = new Size(newWidth, newHeight);

            // 調整位置
            int x = (int)(this.ClientSize.Width * PositionRatio.X - textBox.Width * 0.5);
            int y = (int)(this.ClientSize.Height * PositionRatio.Y - textBox.Height * 0.5);
            textBox.Location = new Point(x, y);
        }
        private void UpdateListBox(ListBox listBox, PointF PositionRatio, Size originalSize, float originalFontSize, float scaleRatio)
        {
            // 調整字體大小
            float newFontSize = originalFontSize * scaleRatio;
            if (Math.Abs(listBox.Font.Size - newFontSize) > 0.1f)
            {
                listBox.Font = new Font(listBox.Font.FontFamily, newFontSize);
            }
            // 按照比例調整大小
            int newWidth = (int)(originalSize.Width * scaleRatio);
            int newHeight = (int)(originalSize.Height * scaleRatio);
            listBox.Size = new Size(newWidth, newHeight);
            // 調整位置
            int x = (int)(this.ClientSize.Width * PositionRatio.X - listBox.Width * 0.5);
            int y = (int)(this.ClientSize.Height * PositionRatio.Y - listBox.Height * 0.5);
            listBox.Location = new Point(x, y);
        }
    }
}
