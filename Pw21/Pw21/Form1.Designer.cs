
namespace Pw21
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
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
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Title title1 = new System.Windows.Forms.DataVisualization.Charting.Title();
            this.dataGridViewGroup = new System.Windows.Forms.DataGridView();
            this.labelGroup = new System.Windows.Forms.Label();
            this.labelTerm = new System.Windows.Forms.Label();
            this.groupBoxObjects = new System.Windows.Forms.GroupBox();
            this.comboBoxObjects = new System.Windows.Forms.ComboBox();
            this.labelName = new System.Windows.Forms.Label();
            this.buttonVed = new System.Windows.Forms.Button();
            this.buttonBall = new System.Windows.Forms.Button();
            this.comboBoxGroup = new System.Windows.Forms.ComboBox();
            this.chartBall = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.buttonChart = new System.Windows.Forms.Button();
            this.buttonExcel = new System.Windows.Forms.Button();
            this.buttonWord = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewGroup)).BeginInit();
            this.groupBoxObjects.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chartBall)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridViewGroup
            // 
            this.dataGridViewGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewGroup.Location = new System.Drawing.Point(18, 48);
            this.dataGridViewGroup.Name = "dataGridViewGroup";
            this.dataGridViewGroup.Size = new System.Drawing.Size(643, 337);
            this.dataGridViewGroup.TabIndex = 0;
            // 
            // labelGroup
            // 
            this.labelGroup.AutoSize = true;
            this.labelGroup.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelGroup.Location = new System.Drawing.Point(14, 14);
            this.labelGroup.Name = "labelGroup";
            this.labelGroup.Size = new System.Drawing.Size(67, 20);
            this.labelGroup.TabIndex = 1;
            this.labelGroup.Text = "Группа";
            // 
            // labelTerm
            // 
            this.labelTerm.AutoSize = true;
            this.labelTerm.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelTerm.Location = new System.Drawing.Point(235, 17);
            this.labelTerm.Name = "labelTerm";
            this.labelTerm.Size = new System.Drawing.Size(83, 17);
            this.labelTerm.TabIndex = 2;
            this.labelTerm.Text = "3 семестр";
            // 
            // groupBoxObjects
            // 
            this.groupBoxObjects.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.groupBoxObjects.Controls.Add(this.comboBoxObjects);
            this.groupBoxObjects.Controls.Add(this.labelName);
            this.groupBoxObjects.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.groupBoxObjects.Location = new System.Drawing.Point(667, 48);
            this.groupBoxObjects.Name = "groupBoxObjects";
            this.groupBoxObjects.Size = new System.Drawing.Size(155, 142);
            this.groupBoxObjects.TabIndex = 3;
            this.groupBoxObjects.TabStop = false;
            this.groupBoxObjects.Text = "Предметы";
            // 
            // comboBoxObjects
            // 
            this.comboBoxObjects.FormattingEnabled = true;
            this.comboBoxObjects.Items.AddRange(new object[] {
            "ТРПП",
            "ОАиП",
            "ОБЖ",
            "Мат. метод",
            "Граф. диз"});
            this.comboBoxObjects.Location = new System.Drawing.Point(9, 42);
            this.comboBoxObjects.Name = "comboBoxObjects";
            this.comboBoxObjects.Size = new System.Drawing.Size(131, 23);
            this.comboBoxObjects.TabIndex = 4;
            this.comboBoxObjects.SelectedIndexChanged += new System.EventHandler(this.comboBoxObjects_SelectedIndexChanged);
            // 
            // labelName
            // 
            this.labelName.AutoSize = true;
            this.labelName.Location = new System.Drawing.Point(6, 26);
            this.labelName.Name = "labelName";
            this.labelName.Size = new System.Drawing.Size(67, 15);
            this.labelName.TabIndex = 2;
            this.labelName.Text = "Название:";
            // 
            // buttonVed
            // 
            this.buttonVed.BackColor = System.Drawing.SystemColors.ControlDark;
            this.buttonVed.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonVed.Location = new System.Drawing.Point(676, 210);
            this.buttonVed.Name = "buttonVed";
            this.buttonVed.Size = new System.Drawing.Size(146, 33);
            this.buttonVed.TabIndex = 4;
            this.buttonVed.Text = "Заполнить ведомость";
            this.buttonVed.UseVisualStyleBackColor = false;
            this.buttonVed.Click += new System.EventHandler(this.buttonVed_Click);
            // 
            // buttonBall
            // 
            this.buttonBall.BackColor = System.Drawing.SystemColors.ControlDark;
            this.buttonBall.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonBall.Location = new System.Drawing.Point(676, 267);
            this.buttonBall.Name = "buttonBall";
            this.buttonBall.Size = new System.Drawing.Size(146, 33);
            this.buttonBall.TabIndex = 5;
            this.buttonBall.Text = "Средний балл";
            this.buttonBall.UseVisualStyleBackColor = false;
            this.buttonBall.Click += new System.EventHandler(this.buttonBall_Click);
            // 
            // comboBoxGroup
            // 
            this.comboBoxGroup.FormattingEnabled = true;
            this.comboBoxGroup.Items.AddRange(new object[] {
            "ИП-21-7К",
            "ИП-21-8К",
            "ИП-21-9К"});
            this.comboBoxGroup.Location = new System.Drawing.Point(87, 14);
            this.comboBoxGroup.Name = "comboBoxGroup";
            this.comboBoxGroup.Size = new System.Drawing.Size(121, 21);
            this.comboBoxGroup.TabIndex = 6;
            this.comboBoxGroup.SelectedIndexChanged += new System.EventHandler(this.comboBoxGroup_SelectedIndexChanged);
            // 
            // chartBall
            // 
            this.chartBall.BackColor = System.Drawing.SystemColors.ActiveBorder;
            chartArea1.Name = "ChartArea1";
            this.chartBall.ChartAreas.Add(chartArea1);
            this.chartBall.Location = new System.Drawing.Point(18, 402);
            this.chartBall.Name = "chartBall";
            this.chartBall.Palette = System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Grayscale;
            series1.ChartArea = "ChartArea1";
            series1.Name = "Series1";
            this.chartBall.Series.Add(series1);
            this.chartBall.Size = new System.Drawing.Size(643, 200);
            this.chartBall.TabIndex = 7;
            this.chartBall.Text = "chart1";
            title1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold);
            title1.Name = "Title";
            title1.Text = "Динамика среднего балла";
            this.chartBall.Titles.Add(title1);
            // 
            // buttonChart
            // 
            this.buttonChart.BackColor = System.Drawing.SystemColors.ControlDark;
            this.buttonChart.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonChart.Location = new System.Drawing.Point(676, 327);
            this.buttonChart.Name = "buttonChart";
            this.buttonChart.Size = new System.Drawing.Size(146, 33);
            this.buttonChart.TabIndex = 9;
            this.buttonChart.Text = "График";
            this.buttonChart.UseVisualStyleBackColor = false;
            this.buttonChart.Click += new System.EventHandler(this.buttonChart_Click);
            // 
            // buttonExcel
            // 
            this.buttonExcel.BackColor = System.Drawing.SystemColors.ControlDark;
            this.buttonExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonExcel.Location = new System.Drawing.Point(676, 386);
            this.buttonExcel.Name = "buttonExcel";
            this.buttonExcel.Size = new System.Drawing.Size(146, 33);
            this.buttonExcel.TabIndex = 10;
            this.buttonExcel.Text = "Экспорт в Excel";
            this.buttonExcel.UseVisualStyleBackColor = false;
            this.buttonExcel.Click += new System.EventHandler(this.buttonExcel_Click);
            // 
            // buttonWord
            // 
            this.buttonWord.BackColor = System.Drawing.SystemColors.ControlDark;
            this.buttonWord.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonWord.Location = new System.Drawing.Point(676, 447);
            this.buttonWord.Name = "buttonWord";
            this.buttonWord.Size = new System.Drawing.Size(146, 33);
            this.buttonWord.TabIndex = 11;
            this.buttonWord.Text = "Экспорт в Word";
            this.buttonWord.UseVisualStyleBackColor = false;
            this.buttonWord.Click += new System.EventHandler(this.buttonWord_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLight;
            this.ClientSize = new System.Drawing.Size(834, 611);
            this.Controls.Add(this.buttonWord);
            this.Controls.Add(this.buttonExcel);
            this.Controls.Add(this.buttonChart);
            this.Controls.Add(this.chartBall);
            this.Controls.Add(this.comboBoxGroup);
            this.Controls.Add(this.buttonBall);
            this.Controls.Add(this.buttonVed);
            this.Controls.Add(this.groupBoxObjects);
            this.Controls.Add(this.labelTerm);
            this.Controls.Add(this.labelGroup);
            this.Controls.Add(this.dataGridViewGroup);
            this.Name = "Form1";
            this.Text = "Сводная ведомость";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewGroup)).EndInit();
            this.groupBoxObjects.ResumeLayout(false);
            this.groupBoxObjects.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chartBall)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridViewGroup;
        private System.Windows.Forms.Label labelGroup;
        private System.Windows.Forms.Label labelTerm;
        private System.Windows.Forms.GroupBox groupBoxObjects;
        private System.Windows.Forms.ComboBox comboBoxObjects;
        private System.Windows.Forms.Label labelName;
        private System.Windows.Forms.Button buttonVed;
        private System.Windows.Forms.Button buttonBall;
        private System.Windows.Forms.ComboBox comboBoxGroup;
        private System.Windows.Forms.DataVisualization.Charting.Chart chartBall;
        private System.Windows.Forms.Button buttonChart;
        private System.Windows.Forms.Button buttonExcel;
        private System.Windows.Forms.Button buttonWord;
    }
}

