using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Printing;
using ExcelObj = Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public partial class Обратная_матрица : Form
    {
        public int MaxOrder;
        double[,] inv;
        double[,] m;
        double[] y;
        double[] x;
        public bool CalcInverse()
        {
            MaxOrder = (int)numericUpDown1.Value;
            int i, k, l;
            bool calcError = false;
            double tempEl;
            i = 1;
            for (k = 0; k <= MaxOrder - 2; k++)
            {
                if (Math.Abs(m[k, k]) > 1e-8)
                {
                    for (i = k; i <= MaxOrder - 2; i++)
                    {
                        if (!calcError)
                        {
                            if (Math.Abs(m[i + 1, k]) > 1e-8)
                            {
                                tempEl = m[i + 1, k];
                                for (l = 0; l <= MaxOrder - 1; l++)
                                {
                                    if (l >= k)
                                        m[i + 1, l] = m[i + 1, l] * m[k, k] - m[k, l] * tempEl;
                                    inv[i + 1, l] = inv[i + 1, l] * m[k, k] - inv[k, l] * tempEl;
                                }
                            }
                        }
                    }

                    if (Math.Abs(m[MaxOrder - 1 - k, MaxOrder - 1 - k]) > 1e-8)
                    {
                        for (i = k; i <= MaxOrder - 2; i++)
                        {
                            if (!calcError)
                            {
                                if (Math.Abs(m[MaxOrder - i - 2, MaxOrder - 1 - k]) > 1e-8)
                                {
                                    tempEl = m[MaxOrder - i - 2, MaxOrder - k - 1];
                                    for (l = MaxOrder - 1; l >= 0; l--)
                                    {
                                        if (l <= MaxOrder - k - 1)
                                            m[MaxOrder - i - 2, l] = m[MaxOrder - i - 2, l] * m[MaxOrder - 1 - k, MaxOrder - 1 - k] - m[MaxOrder - k - 1, l] * tempEl;
                                        inv[MaxOrder - i - 2, l] = inv[MaxOrder - i - 2, l] * m[MaxOrder - 1 - k, MaxOrder - 1 - k] - inv[MaxOrder - k - 1, l] * tempEl;
                                    }
                                }
                            }
                        }
                    }
                    else
                        calcError = true;
                }
                else
                    calcError = true;
            }
            for (k = 0; k < MaxOrder; k++)
            {
                for (i = 0; i < MaxOrder; i++)
                {
                    inv[k, i] = inv[k, i] / m[k, k];
                }
            }
            return !calcError;
        }
        public Обратная_матрица()
        {
            InitializeComponent();
        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            MaxOrder = (int)numericUpDown1.Value;
            if (numericUpDown1.Value == 0)
            {
                MessageBox.Show("Задайте размер матрицы", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown1.Value == 1)
            {
                MessageBox.Show("Задайте размер матрицы", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown1.Value != 0 & numericUpDown1.Value != 1)
            {
                GMatrixA.ColumnCount = MaxOrder;
                GMatrixA.RowCount = MaxOrder;
            
            int i, k;
            DataGridViewCell cell;
            m = new double[MaxOrder, MaxOrder];
            y = new double[MaxOrder];
            x = new double[MaxOrder];
            inv = new double[MaxOrder, MaxOrder];
                for (i = 0; i < MaxOrder; i++)
                {
                    for (k = 0; k < MaxOrder; k++)
                    {
                    cell = GMatrixA[i, k];
                    cell.Value = Convert.ToString(m[k, i]);
                    inv[i, k] = 0.0;
                    if (i == k)
                        inv[i, k] = 1.0;
                    }
                } 
            }
        }
        private void button2_Click_1(object sender, EventArgs e)
        {
            MaxOrder = (int)numericUpDown1.Value;
            if (numericUpDown1.Value == 0)
            {
                MessageBox.Show("Задайте размер матрицы", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown1.Value == 1)
            {
                MessageBox.Show("Задайте размер матрицы", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown1.Value != 0 & numericUpDown1.Value != 1)
            {
                GMatrixA.ColumnCount = MaxOrder;
                GMatrixA.RowCount = MaxOrder;
            int i, k;
            DataGridViewCell cell;
            for (i = 0; i < MaxOrder; i++)
            {
                for (k = 0; k < MaxOrder; k++)
                {
                    cell = GMatrixA[i, k];
                    m[k, i] = Convert.ToDouble(cell.Value);
                }
            }
            if (CalcInverse())
            {
                for (i = 0; i < MaxOrder; i++)
                {
                    for (k = 0; k < MaxOrder; k++)
                    {
                        cell = GMatrixA[i, k];
                        cell.Value = Convert.ToString(Math.Round(inv[k, i], MaxOrder));
                    }
                }
            }
            else
                MessageBox.Show("Матрица вырождена. Обратной матрицы не существует!","Ошибка", MessageBoxButtons.OK);
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            GMatrixA.Rows.Clear();
            GMatrixA.Columns.Clear();
            if (numericUpDown1.Value == 0)
            {
                MessageBox.Show("Задайте размер матрицы", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown1.Value == 1)
            {
                MessageBox.Show("Задайте размер матрицы", "Ошибка", MessageBoxButtons.OK);
            }
            if (numericUpDown1.Value != 0 & numericUpDown1.Value != 1)
            {
                GMatrixA.ColumnCount = MaxOrder;
                GMatrixA.RowCount = MaxOrder;
                MaxOrder = (int)numericUpDown1.Value;
                int i, k;
                for (i = 0; i < MaxOrder; i++)
                {
                    GMatrixA.Columns.Add(" ", " ");
                }
                for (int j = 0; j < MaxOrder; j++)
                {
                    GMatrixA.Rows.Add();
                }
                inv = new double[MaxOrder, MaxOrder];
                DataGridViewCell cell;
                m = new double[MaxOrder, MaxOrder];
                y = new double[MaxOrder];
                x = new double[MaxOrder];
                inv = new double[MaxOrder, MaxOrder];
                for (i = 0; i < MaxOrder; i++)
                {
                    for (k = 0; k < MaxOrder; k++)
                    {
                        cell = GMatrixA[i, k];
                        cell.Value = Convert.ToString(m[k, i]);
                        inv[i, k] = 0.0;
                        if (i == k)
                            inv[i, k] = 1.0;
                    }
                }
                Random rnd = new Random();
                for (i = 0; i < MaxOrder; i++)
                {
                    for (int j = 0; j < MaxOrder; j++)
                    {
                        m[i, j] = rnd.Next(-100, 100);
                        GMatrixA[i, j].Value = m[i, j];
                    }
                }
            }
        }
        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
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
            GMatrixA.DataSource = dt;
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
            SaveTable(GMatrixA);
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