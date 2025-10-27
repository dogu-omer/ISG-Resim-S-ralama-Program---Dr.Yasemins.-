using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace FotoToExcel_SablonB8
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnSelectTemplate_Click(object? sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog();
            ofd.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtTemplate.Text = ofd.FileName;
                try
                {
                    using var fs = new FileStream(ofd.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using var wb = new XSSFWorkbook(fs);
                    cmbSheet.Items.Clear();
                    for (int i = 0; i < wb.NumberOfSheets; i++)
                    {
                        cmbSheet.Items.Add(wb.GetSheetName(i));
                    }
                    if (cmbSheet.Items.Count > 0)
                        cmbSheet.SelectedIndex = 0;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Şablon okunamadı: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnSelectFolder_Click(object? sender, EventArgs e)
        {
            using var fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                txtFolder.Text = fbd.SelectedPath;
            }
        }

        private void btnRun_Click(object? sender, EventArgs e)
        {
            var template = txtTemplate.Text;
            var folder = txtFolder.Text;
            if (string.IsNullOrWhiteSpace(template) || !File.Exists(template))
            {
                MessageBox.Show("Geçerli bir Excel şablon (*.xlsx) seçin.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
            {
                MessageBox.Show("Geçerli bir fotoğraf klasörü seçin.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string sheetName = cmbSheet.SelectedItem?.ToString() ?? string.Empty;

            using var sfd = new SaveFileDialog();
            sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
            sfd.FileName = "FotografRaporu_Sablondan.xlsx";
            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                int targetPx = (int)nudSize.Value; // 90 default
                int startRow = (int)nudStartRow.Value; // 8 default
                string colLetter = txtColumn.Text.Trim().ToUpperInvariant();
                if (string.IsNullOrWhiteSpace(colLetter)) colLetter = "B";
                int colIndex = ColumnLetterToIndex(colLetter);

                CreateFromTemplate(template, sfd.FileName, sheetName, folder, targetPx, colIndex, startRow);
                MessageBox.Show("Bitti: " + sfd.FileName, "Tamamlandı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void CreateFromTemplate(string templatePath, string outPath, string sheetName,
                                               string imageFolder, int targetPx, int colIndex, int startRowIndexExcel)
        {
            using var fs = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var wb = new XSSFWorkbook(fs);

            ISheet sheet = null!;
            if (!string.IsNullOrEmpty(sheetName))
                sheet = wb.GetSheet(sheetName) ?? wb.GetSheetAt(0);
            else
                sheet = wb.GetSheetAt(0);

            sheet.SetColumnWidth(colIndex, PixelsToColumnWidth(targetPx));

            var patterns = new[] { "*.jpg", "*.jpeg", "*.png", "*.bmp", "*.gif" };
            var files = patterns.SelectMany(p => Directory.GetFiles(imageFolder, p, SearchOption.TopDirectoryOnly))
                                .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
                                .ToList();
            if (files.Count == 0)
                throw new InvalidOperationException("Klasörde uygun resim yok.");

            var drawing = sheet.CreateDrawingPatriarch();

            int row0 = startRowIndexExcel - 1; // B8 -> 7
            for (int i = 0; i < files.Count; i++)
            {
                int rowIndex = row0 + i;
                var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                row.HeightInPoints = PixelsToPoints(targetPx);

                byte[] img = ResizeToSquare(files[i], targetPx);
                int picType = GetPictureTypeByExt(files[i]);
                int picIdx = wb.AddPicture(img, picType);

                var anchor = wb.GetCreationHelper().CreateClientAnchor();
                anchor.Col1 = colIndex;
                anchor.Row1 = rowIndex;
                anchor.Col2 = colIndex + 1;
                anchor.Row2 = rowIndex + 1;
                anchor.AnchorType = AnchorType.MoveDontResize;

                var picture = drawing.CreatePicture(anchor, picIdx);
                picture.Resize(1.0);
            }

            using var os = new FileStream(outPath, FileMode.Create, FileAccess.Write);
            wb.Write(os);
        }

        private static int ColumnLetterToIndex(string col)
        {
            int sum = 0;
            for (int i = 0; i < col.Length; i++)
            {
                sum *= 26;
                sum += (col[i] - 'A' + 1);
            }
            return sum - 1;
        }

        private static int PixelsToColumnWidth(int pixels)
        {
            double widthChars = (pixels - 5.0) / 7.0;
            return (int)Math.Round(widthChars * 256, MidpointRounding.AwayFromZero);
        }

        private static float PixelsToPoints(int pixels, float dpi = 96f)
        {
            return pixels * 72f / dpi;
        }

        private static byte[] ResizeToSquare(string path, int targetPx)
        {
            using var src = (System.Drawing.Image)System.Drawing.Image.FromFile(path);
            using var bmp = new Bitmap(targetPx, targetPx);
            using (var g = Graphics.FromImage(bmp))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;

                float ratio = Math.Min((float)targetPx / src.Width, (float)targetPx / src.Height);
                int w = (int)(src.Width * ratio);
                int h = (int)(src.Height * ratio);
                int x = (targetPx - w) / 2;
                int y = (targetPx - h) / 2;
                var fitRect = new Rectangle(x, y, w, h);

                g.Clear(Color.White);
                g.DrawImage(src, fitRect);
            }

            using var ms = new MemoryStream();
            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            return ms.ToArray();
        }

        private static int GetPictureTypeByExt(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            return ext switch
            {
                ".jpg" or ".jpeg" => (int)PictureType.JPEG,
                ".png" => (int)PictureType.PNG,
                ".bmp" => (int)PictureType.PNG,
                ".gif" => (int)PictureType.PNG,
                _ => (int)PictureType.PNG
            };
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }
    }
}
