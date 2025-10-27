using System.Windows.Forms;

namespace FotoToExcel_SablonB8
{
    partial class MainForm
    {
        private Button btnSelectTemplate;
        private TextBox txtTemplate;
        private ComboBox cmbSheet;
        private Button btnSelectFolder;
        private TextBox txtFolder;
        private NumericUpDown nudSize;
        private NumericUpDown nudStartRow;
        private TextBox txtColumn;
        private Button btnRun;
        private Label lbl1, lbl2, lbl3, lbl4, lbl5, lbl6;

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            btnSelectTemplate = new Button();
            txtTemplate = new TextBox();
            cmbSheet = new ComboBox();
            btnSelectFolder = new Button();
            txtFolder = new TextBox();
            nudSize = new NumericUpDown();
            nudStartRow = new NumericUpDown();
            txtColumn = new TextBox();
            btnRun = new Button();
            lbl1 = new Label();
            lbl2 = new Label();
            lbl3 = new Label();
            lbl4 = new Label();
            lbl5 = new Label();
            lbl6 = new Label();
            label1 = new Label();
            linkLabel1 = new LinkLabel();
            label2 = new Label();
            ((System.ComponentModel.ISupportInitialize)nudSize).BeginInit();
            ((System.ComponentModel.ISupportInitialize)nudStartRow).BeginInit();
            SuspendLayout();
            // 
            // btnSelectTemplate
            // 
            btnSelectTemplate.Location = new Point(24, 20);
            btnSelectTemplate.Name = "btnSelectTemplate";
            btnSelectTemplate.Size = new Size(140, 30);
            btnSelectTemplate.TabIndex = 0;
            btnSelectTemplate.Text = "Şablon Seç (.xlsx)";
            btnSelectTemplate.Click += btnSelectTemplate_Click;
            // 
            // txtTemplate
            // 
            txtTemplate.Location = new Point(180, 23);
            txtTemplate.Name = "txtTemplate";
            txtTemplate.Size = new Size(520, 23);
            txtTemplate.TabIndex = 1;
            // 
            // cmbSheet
            // 
            cmbSheet.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbSheet.Location = new Point(180, 57);
            cmbSheet.Name = "cmbSheet";
            cmbSheet.Size = new Size(220, 23);
            cmbSheet.TabIndex = 3;
            // 
            // btnSelectFolder
            // 
            btnSelectFolder.Location = new Point(24, 96);
            btnSelectFolder.Name = "btnSelectFolder";
            btnSelectFolder.Size = new Size(140, 30);
            btnSelectFolder.TabIndex = 4;
            btnSelectFolder.Text = "Foto Klasörü";
            btnSelectFolder.Click += btnSelectFolder_Click;
            // 
            // txtFolder
            // 
            txtFolder.Location = new Point(180, 99);
            txtFolder.Name = "txtFolder";
            txtFolder.Size = new Size(520, 23);
            txtFolder.TabIndex = 5;
            // 
            // nudSize
            // 
            nudSize.Location = new Point(180, 138);
            nudSize.Maximum = new decimal(new int[] { 600, 0, 0, 0 });
            nudSize.Minimum = new decimal(new int[] { 40, 0, 0, 0 });
            nudSize.Name = "nudSize";
            nudSize.Size = new Size(100, 23);
            nudSize.TabIndex = 7;
            nudSize.Value = new decimal(new int[] { 90, 0, 0, 0 });
            // 
            // nudStartRow
            // 
            nudStartRow.Location = new Point(180, 170);
            nudStartRow.Maximum = new decimal(new int[] { 100000, 0, 0, 0 });
            nudStartRow.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            nudStartRow.Name = "nudStartRow";
            nudStartRow.Size = new Size(100, 23);
            nudStartRow.TabIndex = 9;
            nudStartRow.Value = new decimal(new int[] { 8, 0, 0, 0 });
            // 
            // txtColumn
            // 
            txtColumn.Location = new Point(180, 202);
            txtColumn.Name = "txtColumn";
            txtColumn.Size = new Size(100, 23);
            txtColumn.TabIndex = 11;
            txtColumn.Text = "B";
            // 
            // btnRun
            // 
            btnRun.Location = new Point(180, 240);
            btnRun.Name = "btnRun";
            btnRun.Size = new Size(220, 34);
            btnRun.TabIndex = 12;
            btnRun.Text = "Şablondan Dışa Aktar";
            btnRun.Click += btnRun_Click;
            // 
            // lbl1
            // 
            lbl1.AutoSize = true;
            lbl1.Location = new Point(24, 60);
            lbl1.Name = "lbl1";
            lbl1.Size = new Size(38, 15);
            lbl1.TabIndex = 2;
            lbl1.Text = "Sayfa:";
            // 
            // lbl2
            // 
            lbl2.AutoSize = true;
            lbl2.Location = new Point(24, 140);
            lbl2.Name = "lbl2";
            lbl2.Size = new Size(64, 15);
            lbl2.TabIndex = 6;
            lbl2.Text = "Boyut (px):";
            // 
            // lbl3
            // 
            lbl3.AutoSize = true;
            lbl3.Location = new Point(24, 172);
            lbl3.Name = "lbl3";
            lbl3.Size = new Size(127, 15);
            lbl3.TabIndex = 8;
            lbl3.Text = "Başlangıç Satırı (B8=8):";
            // 
            // lbl4
            // 
            lbl4.AutoSize = true;
            lbl4.Location = new Point(24, 204);
            lbl4.Name = "lbl4";
            lbl4.Size = new Size(59, 15);
            lbl4.TabIndex = 10;
            lbl4.Text = "Sütun (B):";
            // 
            // lbl5
            // 
            lbl5.Location = new Point(24, 285);
            lbl5.Name = "lbl5";
            lbl5.Size = new Size(676, 18);
            lbl5.TabIndex = 13;
            lbl5.Text = "Not: Şablon dosyasını BOZMADAN yeni bir dosya oluşturur.";
            // 
            // lbl6
            // 
            lbl6.Location = new Point(24, 305);
            lbl6.Name = "lbl6";
            lbl6.Size = new Size(676, 36);
            lbl6.TabIndex = 14;
            lbl6.Text = "B8’den başlar, tek sütunda alt alta 90x90 (değiştirilebilir).";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(50, 341);
            label1.Name = "label1";
            label1.Size = new Size(146, 15);
            label1.TabIndex = 15;
            label1.Text = "Ömer DOĞU Yük.Bil.Müh. ";
            // 
            // linkLabel1
            // 
            linkLabel1.AutoSize = true;
            linkLabel1.LinkColor = Color.SpringGreen;
            linkLabel1.LinkVisited = true;
            linkLabel1.Location = new Point(59, 399);
            linkLabel1.Name = "linkLabel1";
            linkLabel1.Size = new Size(120, 15);
            linkLabel1.TabIndex = 16;
            linkLabel1.TabStop = true;
            linkLabel1.Tag = "ÖMER DOĞU ";
            linkLabel1.Text = "www.omerdogu.com";
            linkLabel1.VisitedLinkColor = Color.FromArgb(192, 0, 0);
            linkLabel1.LinkClicked += linkLabel1_LinkClicked;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(50, 368);
            label2.Name = "label2";
            label2.Size = new Size(150, 15);
            label2.TabIndex = 17;
            label2.Text = "B Sınıfı İş Güvenliği Uzmanı";
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.DarkCyan;
            ClientSize = new Size(767, 460);
            Controls.Add(label2);
            Controls.Add(linkLabel1);
            Controls.Add(label1);
            Controls.Add(btnSelectTemplate);
            Controls.Add(txtTemplate);
            Controls.Add(lbl1);
            Controls.Add(cmbSheet);
            Controls.Add(btnSelectFolder);
            Controls.Add(txtFolder);
            Controls.Add(lbl2);
            Controls.Add(nudSize);
            Controls.Add(lbl3);
            Controls.Add(nudStartRow);
            Controls.Add(lbl4);
            Controls.Add(txtColumn);
            Controls.Add(btnRun);
            Controls.Add(lbl5);
            Controls.Add(lbl6);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "MainForm";
            Tag = "www.omerdogu.com";
            Text = "Dr.YaseminS. Karadeniz Kızı";
            Load += MainForm_Load;
            ((System.ComponentModel.ISupportInitialize)nudSize).EndInit();
            ((System.ComponentModel.ISupportInitialize)nudStartRow).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        private Label label1;
        private LinkLabel linkLabel1;
        private Label label2;
    }
}
