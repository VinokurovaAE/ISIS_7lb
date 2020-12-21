using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ISIS_5lb
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.RowCount = 5;
            dataGridView1.ColumnCount = 5;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Word.Application winword = new Word.Application();
            winword.Visible = false;
            object missing = System.Reflection.Missing.Value;
            Word.Document document = winword.Documents.Add(ref missing);
            Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
            object start = 0;
            object end = 0;
            Word.Range rng = document.Range(ref start, ref end);
            rng.Text = richTextBox1.Text;
            rng.Font.Italic = 1;                
            para1.Range.InsertParagraphAfter();
            winword.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Excel.Application ObjExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook;
            Excel.Worksheet ObjWorkSheet;
            ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
            ObjWorkSheet = ObjWorkBook.Sheets[1];
            ObjWorkSheet.Cells.NumberFormat = "@";
            for (int i=0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    ObjWorkSheet.Cells[i + 1, j + 1] = dataGridView1[j, i].Value;
                }
            }
            ObjExcel.Visible = true;
        }
    }
}
