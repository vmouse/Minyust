namespace MinUstPdf
{
    partial class Form1
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.start = new System.Windows.Forms.Button();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.pathCSV = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.pathPDF = new System.Windows.Forms.TextBox();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.pb = new System.Windows.Forms.ToolStripProgressBar();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.label4 = new System.Windows.Forms.Label();
            this.minSumm = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.viewReport = new System.Windows.Forms.TextBox();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // start
            // 
            this.start.Location = new System.Drawing.Point(12, 271);
            this.start.Name = "start";
            this.start.Size = new System.Drawing.Size(75, 23);
            this.start.TabIndex = 0;
            this.start.Text = "Старт";
            this.start.UseVisualStyleBackColor = true;
            this.start.Click += new System.EventHandler(this.button1_Click);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(405, 12);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(586, 303);
            this.richTextBox1.TabIndex = 1;
            this.richTextBox1.Text = "";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "\"Файлы отчетов|*.pdf|Все файлы|*.*\"";
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(251, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "1. Укажите путь, где стоит сохранить CSV файл";
            // 
            // pathCSV
            // 
            this.pathCSV.Location = new System.Drawing.Point(15, 41);
            this.pathCSV.Name = "pathCSV";
            this.pathCSV.ReadOnly = true;
            this.pathCSV.Size = new System.Drawing.Size(248, 20);
            this.pathCSV.TabIndex = 3;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(269, 41);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 20);
            this.button2.TabIndex = 4;
            this.button2.Text = "Обзор";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 80);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(331, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "2. Укажите папку с PDF файлами или путь к индексному файлу";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(269, 96);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 20);
            this.button3.TabIndex = 7;
            this.button3.Text = "Обзор";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // pathPDF
            // 
            this.pathPDF.Location = new System.Drawing.Point(15, 96);
            this.pathPDF.Name = "pathPDF";
            this.pathPDF.ReadOnly = true;
            this.pathPDF.Size = new System.Drawing.Size(248, 20);
            this.pathPDF.TabIndex = 6;
            this.pathPDF.TextChanged += new System.EventHandler(this.pathPDF_TextChanged);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.pb});
            this.statusStrip1.Location = new System.Drawing.Point(0, 322);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(360, 22);
            this.statusStrip1.TabIndex = 9;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // pb
            // 
            this.pb.Name = "pb";
            this.pb.Size = new System.Drawing.Size(200, 16);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.FileName = "parser.csv";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 137);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(83, 13);
            this.label4.TabIndex = 13;
            this.label4.Text = "3. Фильтрация";
            // 
            // minSumm
            // 
            this.minSumm.Location = new System.Drawing.Point(15, 176);
            this.minSumm.Name = "minSumm";
            this.minSumm.Size = new System.Drawing.Size(248, 20);
            this.minSumm.TabIndex = 14;
            this.minSumm.Text = "-";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 156);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(258, 13);
            this.label5.TabIndex = 15;
            this.label5.Text = "3.1 Минимальная сумма разделов 1 и 2 (тыс.руб)";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(12, 199);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(80, 13);
            this.label6.TabIndex = 17;
            this.label6.Text = "3.2 Вид отчета";
            // 
            // viewReport
            // 
            this.viewReport.Location = new System.Drawing.Point(15, 219);
            this.viewReport.Name = "viewReport";
            this.viewReport.ReadOnly = true;
            this.viewReport.Size = new System.Drawing.Size(248, 20);
            this.viewReport.TabIndex = 16;
            this.viewReport.Text = "ОН0002";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(360, 344);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.viewReport);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.minSumm);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.pathPDF);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.pathCSV);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.start);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Парсер";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button start;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox pathCSV;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox pathPDF;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripProgressBar pb;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox minSumm;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox viewReport;
    }
}

