using System;
using System.Windows.Forms;
using Word1 = Microsoft.Office.Interop.Word;

namespace Word
{
    public partial class Word : Form
    {
        public Word()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e) { }

        private void button1_Click(object sender, EventArgs e)
        {
            String[] names = { "Андрей - раб", "Леха - раб", "Андрей - раб", "Андрей - раб", "Андрей - раб" };
            String[] number = { "79876543210", "79876543210", "79876543210", "79876543210", "79876543210" };

            var application = new Word1.Application();
            Word1.Document document = application.Documents.Add();
            Word1.Paragraph userParagraph = document.Paragraphs.Add();
            Word1.Range userRange = userParagraph.Range;
            userRange.Text = "Справочник телефонов";
            userRange.InsertParagraph();

            Word1.Paragraph tableParagraph = document.Paragraphs.Add();
            Word1.Range tableRange = userParagraph.Range;
            Word1.Table numbersTable = document.Tables.Add(tableRange, names.Length, 2);
            numbersTable.Borders.InsideLineStyle = Word1.WdLineStyle.wdLineStyleSingle;
            numbersTable.Borders.OutsideLineStyle = Word1.WdLineStyle.wdLineStyleSingle;

            Word1.Range cellRange;
            cellRange = numbersTable.Cell(1, 1).Range;
            cellRange.Text = "ФИО";
            cellRange = numbersTable.Cell(1, 2).Range;
            cellRange.Text = "Номер телефона";

            // Выравнивание по центру и жирный шрифт
            numbersTable.Rows[1].Range.Bold = 1;
            numbersTable.Rows[1].Range.ParagraphFormat.Alignment = Word1.WdParagraphAlignment.wdAlignParagraphCenter;

            // Заполнение таблицы
            for (int i = 0; i < names.Length; i++)
            {
                cellRange = numbersTable.Cell(i + 2, 1).Range;
                cellRange.Text = names[i];
                cellRange.ParagraphFormat.Alignment = Word1.WdParagraphAlignment.wdAlignParagraphCenter;
                cellRange = numbersTable.Cell(i + 2, 2).Range;
                cellRange.Text = number[i];
                cellRange.ParagraphFormat.Alignment = Word1.WdParagraphAlignment.wdAlignParagraphCenter;
            }

            // Сохранение документа
            application.Visible = true;
            document.SaveAs2(@"X:\Test.docx");
        }
    }
}
