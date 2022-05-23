using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelObj = Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing.Printing;

namespace WindowsFormsApp1
{
    public partial class Умножение : Form
    {
        public Умножение()
        {
            InitializeComponent();
        }
        public double[,] A;
        public double[,] B;
        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            if (numericUpDown2.Value == 0)
            {
                MessageBox.Show("Задайте количество столбцов!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown2.Value == 1)
            {
                MessageBox.Show("Задайте корректное количество столбцов!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown1.Value == 0)
            {
                MessageBox.Show("Задайте количество строк!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown1.Value == 1)
            {
                MessageBox.Show("Задайте корректное количество строк!", "Ошибка", MessageBoxButtons.OK);
            }
            if ((numericUpDown1.Value != 0 & numericUpDown1.Value != 1) & (numericUpDown2.Value != 0 & numericUpDown1.Value != 1) )
            {
                int b = (int)numericUpDown2.Value;
                for (int i = 0; i < b; i++)
                {
                    dataGridView1.Columns.Add("", "");
                }
                int a = (int)numericUpDown1.Value;
                for (int j = 0; j < a; j++)
                {
                    dataGridView1.Rows.Add();
                }
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();
            if (numericUpDown4.Value == 0)
            {
                MessageBox.Show("Задайте количество столбцов!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown4.Value == 1)
            {
                MessageBox.Show("Задайте корректное количество столбцов!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown3.Value == 0)
            {
                MessageBox.Show("Задайте количество строк!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown3.Value == 1)
            {
                MessageBox.Show("Задайте корректное количество строк!", "Ошибка", MessageBoxButtons.OK);
            }
            if ((numericUpDown3.Value != 0 & numericUpDown1.Value != 1) & (numericUpDown4.Value != 0 & numericUpDown1.Value != 1) )
            {
                int a = (int)numericUpDown4.Value;
                for (int i = 0; i < a; i++)
                {
                dataGridView2.Columns.Add("", "");
                }
                int b = (int)numericUpDown3.Value;
                for (int j = 0; j < b; j++)
                {
                dataGridView2.Rows.Add();
                }
            }         
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (numericUpDown4.Value == 0)
            {
                MessageBox.Show("Задайте количество столбцов!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown4.Value == 1)
            {
                MessageBox.Show("Задайте корректное количество столбцов!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown3.Value == 0)
            {
                MessageBox.Show("Задайте количество строк!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown3.Value == 1)
            {
                MessageBox.Show("Задайте корректное количество строк!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown1.Value == 0)
            {
                MessageBox.Show("Задайте количество столбцов!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown1.Value == 1)
            {
                MessageBox.Show("Задайте корректное количество столбцов!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown2.Value == 0)
            {
                MessageBox.Show("Задайте количество строк!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown2.Value == 1)
            {
                MessageBox.Show("Задайте корректное количество строк!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown1.Value == 0)
            {
                MessageBox.Show("Задайте количество столбцов!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown1.Value == 1)
            {
                MessageBox.Show("Задайте корректное количество столбцов!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown2.Value != numericUpDown3.Value)
            {
                MessageBox.Show("Умножение невозможно!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown1.Value != numericUpDown4.Value)
            {
                MessageBox.Show("Умножение невозможно!", "Ошибка", MessageBoxButtons.OK);
            }
            dataGridView3.Rows.Clear();
            dataGridView3.Columns.Clear();
            double[,] A = new double[dataGridView1.RowCount, dataGridView1.ColumnCount];
            double[,] B = new double[dataGridView2.RowCount, dataGridView2.ColumnCount];
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
                for (int j = 0; j < dataGridView1.RowCount; j++)
                    A[j, i] = Convert.ToInt32(dataGridView1[i, j].Value);
            for (int i = 0; i < dataGridView2.ColumnCount; i++)
                for (int j = 0; j < dataGridView2.RowCount; j++)
                    B[j, i] = Convert.ToInt32(dataGridView2[i, j].Value);
            for (int i = 0; i < dataGridView2.ColumnCount; i++)
            {
                dataGridView3.Columns.Add("", "");
            }
            for (int j = 0; j < dataGridView1.RowCount; j++)
            {
                dataGridView3.Rows.Add();
            }
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                {
                    double s = 0;
                    for (int k = 0; k < dataGridView2.RowCount; k++)
                        s += A[i, k] * B[k, j];
                        dataGridView3[j, i].Value = s;
                }
            }        
        }
        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();
            if (numericUpDown2.Value == 0)
            {
                MessageBox.Show("Задайте количество столбцов!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown2.Value == 1)
            {
                MessageBox.Show("Задайте корректное количество столбцов!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown1.Value == 0)
            {
                MessageBox.Show("Задайте количество строк!", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown1.Value == 1)
            {
                MessageBox.Show("Задайте корректное количество строк!", "Ошибка", MessageBoxButtons.OK);
            }
            if ((numericUpDown1.Value != 0 & numericUpDown1.Value != 1) & (numericUpDown2.Value != 0 & numericUpDown1.Value != 1))
            {
                int a = (int)numericUpDown1.Value;
                int b = (int)numericUpDown2.Value;
                for (int i = 0; i < a; i++)
                {
                    dataGridView1.Columns.Add(" ", " ");
                }
                for (int i = 0; i < b; i++)
                {
                    dataGridView2.Columns.Add(" ", " ");
                }
                for (int j = 0; j < b; j++)
                {
                    dataGridView1.Rows.Add();
                }
                for (int j = 0; j < a; j++)
                {
                    dataGridView2.Rows.Add();
                }
                A = new double[a, b];
                B = new double[b, a];
                Random rnd = new Random();
                for (int i = 0; i < a; i++)
                {
                    for (int j = 0; j < b; j++)
                    {
                        A[i, j] = rnd.Next(-100, 100);
                        dataGridView1[i, j].Value = A[i, j];
                    }
                }
                for (int i = 0; i < b; i++)
                {
                    for (int j = 0; j < a; j++)
                    {
                        B[i, j] = rnd.Next(-100, 100);
                        dataGridView2[i, j].Value = B[i, j];
                    }
                }
            }
        }
        private void вПервуюМатрицуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Microsoft Excel (*.xls*)|*.xls*";
            ofd.Title = "Выберите документ для загрузки данных";
            if (ofd.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("Вы не выбрали файл для открытия", "Загрузка данных...", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ofd.FileName + ";Extended Properties='Excel 12.0 XML;HDR=YES;IMEX=1';";
            System.Data.OleDb.OleDbConnection con = new System.Data.OleDb.OleDbConnection(constr);
            con.Open();
            DataSet ds = new DataSet();
            DataTable schemaTable = con.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];
            string select = String.Format("SELECT * FROM [{0}]", sheet1);
            System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(select, con);
            ad.Fill(ds);
            DataTable dt = ds.Tables[0];
            con.Close();
            con.Dispose();
            dataGridView1.DataSource = dt;
        }
        private void воВторуюМатрицуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Microsoft Excel (*.xls*)|*.xls*";
            ofd.Title = "Выберите документ для загрузки данных";
            if (ofd.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("Вы не выбрали файл для открытия", "Загрузка данных...", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ofd.FileName + ";Extended Properties='Excel 12.0 XML;HDR=YES;IMEX=1';";
            System.Data.OleDb.OleDbConnection con = new System.Data.OleDb.OleDbConnection(constr);
            con.Open();
            DataSet ds = new DataSet();
            DataTable schemaTable = con.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];
            string select = String.Format("SELECT * FROM [{0}]", sheet1);
            System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(select, con);
            ad.Fill(ds);
            DataTable dt = ds.Tables[0];
            con.Close();
            con.Dispose();
            dataGridView2.DataSource = dt;
        }
        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void печатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintDocument printDocument = new PrintDocument();
            PrintDialog printDialog = new PrintDialog();
            printDialog.Document = printDocument;
            if (printDialog.ShowDialog() == DialogResult.OK)
                printDialog.Document.Print();
        }
        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveTable(dataGridView3);
        }
        void SaveTable(DataGridView saving)
        {
            string path = System.IO.Directory.GetCurrentDirectory() + @"\" + "Save_channel.xlsx";
            Excel.Application excelapp = new Excel.Application();
            Excel.Workbook workbook = excelapp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            for (int i = 1; i < saving.RowCount + 1; i++)
            {
                for (int j = 1; j < saving.ColumnCount + 1; j++)
                {
                    worksheet.Rows[i].Columns[j] = saving.Rows[i - 1].Cells[j - 1].Value;
                }
            }
            excelapp.AlertBeforeOverwriting = false;
            workbook.SaveAs(path);
            excelapp.Quit();
        }
    }
}