using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using ExcelObj = Microsoft.Office.Interop.Excel;
using System.Drawing.Printing;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public partial class Вычитание : Form
    {
        public Вычитание()
        {
            InitializeComponent();
        }
        double[,] ar;
        int n, m;
        double[,] ar2;
        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();
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
            if ((numericUpDown1.Value != 0 & numericUpDown1.Value != 1)& (numericUpDown2.Value != 0 & numericUpDown1.Value != 1))
            {
                n = (int)numericUpDown1.Value; 
                for (int i = 0; i < n; i++)
                {
                    dataGridView1.Columns.Add(" ", " ");
                }
                for (int i = 0; i < n; i++)
                {
                    dataGridView2.Columns.Add(" ", " ");
                }
                m = (int)numericUpDown2.Value;
                for (int j = 0; j < m; j++)
                {
                    dataGridView1.Rows.Add();
                }
                for (int j = 0; j < m; j++)
                {
                    dataGridView2.Rows.Add();
                }
            }
            ar = new double[n, m];
            ar2 = new double[n, m];
        }
        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();
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
            if ((numericUpDown1.Value != 0 & numericUpDown1.Value != 1) & (numericUpDown2.Value != 0 & numericUpDown1.Value != 1))
            {
                n = (int)numericUpDown1.Value;
                for (int i = 0; i < n; i++)
                {
                    dataGridView1.Columns.Add(" ", " ");
                }
                for (int i = 0; i < n; i++)
                {
                    dataGridView2.Columns.Add(" ", " ");
                }
                m = (int)numericUpDown2.Value;
                for (int j = 0; j < m; j++)
                {
                    dataGridView1.Rows.Add();
                }
                for (int j = 0; j < m; j++)
                {
                    dataGridView2.Rows.Add();
                }
                ar = new double[n, m];
                ar2 = new double[n, m];
                Random rnd = new Random();
                for (int i = 0; i < n; i++)
                {
                    for (int j = 0; j < m; j++)
                    {
                        ar[i, j] = rnd.Next(-100, 100);
                        dataGridView1[i, j].Value = ar[i, j];
                        ar2[i, j] = rnd.Next(-100, 100);
                        dataGridView2[i, j].Value = ar2[i, j];
                    }
                }
            }
        }
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
        }
        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
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
            for (int i = 1; i < saving.RowCount +1; i++)
            {
                for (int j = 1; j< saving.ColumnCount+1; j++)
                {
                    worksheet.Rows[i].Columns[j] = saving.Rows[i - 1].Cells[j - 1].Value;
                }
            }
            excelapp.AlertBeforeOverwriting = false;
            workbook.SaveAs(path);
            excelapp.Quit();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView3.Rows.Clear();
            dataGridView3.Columns.Clear();
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
            
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < m; j++)
                {
                    ar[i, j] = Convert.ToDouble(dataGridView1[i,j].Value);
                }
            }
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < m; j++)
                {
                    ar2[i, j] = Convert.ToDouble(dataGridView2[i,j].Value);
                }
            }
            dataGridView3.ColumnCount = n;
            dataGridView3.RowCount = m;
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < m; j++)
                {
                    dataGridView3[i,j].Value = ar[i, j] - ar2[i, j];
                }
            }
        }
    }
}