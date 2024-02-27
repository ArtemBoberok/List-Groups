using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Pw21
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        Random rand = new Random();
        int n = 0;  // Строки
        int m = 1;  // Столбцы

        private void comboBoxGroup_SelectedIndexChanged(object sender, EventArgs e)  // Заполнение таблицы студентами
        {
            n = 0;        
            dataGridViewGroup.TopLeftHeaderCell.Value = "ФИО";  // Ячейка левого верхнего угла таблицы
            dataGridViewGroup.RowHeadersWidth = 150;        	// Ширина столбца заголовков строк
            string Group = comboBoxGroup.SelectedItem.ToString();
            int i = 0;

            try
            {
                StreamReader sr = new StreamReader(Group + ".txt");
                while (!sr.EndOfStream)
                {
                    sr.ReadLine();
                    n++;
                }
                sr.Close();
                dataGridViewGroup.RowCount = n + 1;  // Строки

                sr = new StreamReader(Group + ".txt");
                while (!sr.EndOfStream)
                {
                    dataGridViewGroup.Rows[i].HeaderCell.Value = sr.ReadLine();
                    i++;
                }
                sr.Close();
                dataGridViewGroup.Rows[i].HeaderCell.Value = "Средний балл:";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка чтения \n" + ex.ToString());
            }          
        }

        private void comboBoxObjects_SelectedIndexChanged(object sender, EventArgs e)  // Заполнение
        {
            dataGridViewGroup.ColumnCount = m;
            dataGridViewGroup.Columns[m-1].HeaderCell.Value = comboBoxObjects.SelectedItem.ToString();
            comboBoxObjects.Items.Remove(comboBoxObjects.SelectedItem);
            m += 1;
        }

        private void buttonVed_Click(object sender, EventArgs e)
        {
            buttonVed.Enabled = false;
            for (int i = 0; i < n; i++)        //Цикл по строкам таблицы
            {
                for (int j = 0; j < m-1; j++)  //Цикл по столбцам таблицы
                {
                    dataGridViewGroup.Rows[i].Cells[j].Value = rand.Next(2, 6);
                }
            }

            for (int i = 0; i < n; i++)          //Все строки
                for (int j = 0; j < m - 1; j++)  //Все столбцы
                {
                    dataGridViewGroup[j, i].ToolTipText = "Ячейка";  //Текст всплывающей подсказки
                    dataGridViewGroup[j, i].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridViewGroup[j, i].Style.SelectionBackColor = Color.Orange; //Фон выделенной ячейки
                    dataGridViewGroup[j, i].Style.SelectionForeColor = Color.White;  //Текст выделенной ячейки
                    if (dataGridViewGroup[j, i].Value.ToString() == "2")             //Проверка значения в ячейке
                    {
                        dataGridViewGroup[j, i].Style.BackColor = Color.Red;         //Цвет фон для 5
                        dataGridViewGroup[j, i].Style.ForeColor = Color.Yellow;      //Цвет текста для 5
                    }
                    if (dataGridViewGroup[j, i].Value.ToString() == "5")             //Проверка значения в ячейке
                    {
                        dataGridViewGroup[j, i].Style.BackColor = Color.Green;       //Цвет фон для 5
                        dataGridViewGroup[j, i].Style.ForeColor = Color.Yellow;      //Цвет текста для 5
                    }
                }
        }

        private void buttonBall_Click(object sender, EventArgs e)
        {
            buttonBall.Enabled = false;
            double[] ball = new double[dataGridViewGroup.ColumnCount];
            for (int i = 0; i < dataGridViewGroup.ColumnCount; i++)
            {
                double result = 0;
                for (int j = 0; j < dataGridViewGroup.RowCount - 1; j++)
                {
                    result+= Convert.ToDouble(dataGridViewGroup[i, j].Value);
                }
                ball[i] = result / (dataGridViewGroup.RowCount - 1);
                dataGridViewGroup[i, dataGridViewGroup.RowCount - 1].Value = Math.Round(ball[i], 2);
            }
        }

        private void buttonChart_Click(object sender, EventArgs e)
        {
            buttonChart.Enabled = false;
            chartBall.Series[0].Points.Clear();
            for(int i =0; i < m -1; i++)
            {
                chartBall.Series[0].Points.AddXY(dataGridViewGroup.Columns[i].HeaderCell.Value, dataGridViewGroup[i, dataGridViewGroup.RowCount - 1].Value);
            }
        }

        private void buttonExcel_Click(object sender, EventArgs e)
        {
            var application = new Excel.Application();
            application.SheetsInNewWorkbook = 1;
            Excel.Workbook workbook = application.Workbooks.Add(Type.Missing);
            Excel.Worksheet worksheet = application.Worksheets.Item[1];
            worksheet.Name = "Сводная ведомость";

            worksheet.Cells[1,1]= dataGridViewGroup.TopLeftHeaderCell.Value;

            for (int i = 0; i < dataGridViewGroup.ColumnCount; i++)
            {
                string cn = dataGridViewGroup.Columns[i].HeaderText;
                worksheet.Cells[1, i + 2] = cn;
                for (int j = 0; j < dataGridViewGroup.RowCount; j++)
                {
                    worksheet.Cells[j + 2, 1] = dataGridViewGroup.Rows[j].HeaderCell.Value;
                    worksheet.Cells[j + 2, i + 2] = dataGridViewGroup.Rows[j].Cells[i].Value;
                }
            }
            application.Visible = true;
            buttonChart.Enabled = false;
        }

        private void buttonWord_Click(object sender, EventArgs e)
        {
            var application = new Word.Application();
            Word.Document document = application.Documents.Add();
            Word.Paragraph userParagraph = document.Paragraphs.Add();
            Word.Range userRange = userParagraph.Range;
            userRange.Text = "Сводная ведомость " + comboBoxGroup.SelectedItem.ToString();
            userRange.InsertParagraph();

            // Добавление таблицы
            Word.Paragraph tableParagraph = document.Paragraphs.Add();
            Word.Range tableRange = userParagraph.Range;
            Word.Table numbersTable = document.Tables.Add(tableRange, n + 2, m);
            numbersTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            numbersTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            Word.Range cellRange;
            cellRange = numbersTable.Cell(1, 1).Range;
            cellRange.Text = "ФИО";
            for (int i = 0; i < m - 1; i++)
            {
                cellRange = numbersTable.Cell(1, i + 2).Range;
                cellRange.Text = dataGridViewGroup.Columns[i].HeaderText;
            }

            // Выравнивание по центру и жирный шрифт
            numbersTable.Rows[1].Range.Bold = 1;
            numbersTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            // Заполнение таблицы
            for (int i = 0; i < n + 1; i++)
            {
                for (int j = 0; j < m - 1; j++)
                {
                    cellRange = numbersTable.Cell(i + 2, 1).Range;
                    cellRange.Text = dataGridViewGroup.Rows[i].HeaderCell.Value.ToString();
                    cellRange = numbersTable.Cell(i + 2, j + 2).Range;
                    cellRange.Text = dataGridViewGroup[j, i].Value.ToString();
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
            }

            // Сохранение документа
            application.Visible = true;
            document.SaveAs2(@"X:\Test.docx");

            buttonChart.Enabled = false;
        }
    }
}