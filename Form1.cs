using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Обратная_матрица o = new Обратная_матрица();
            o.ShowDialog();

        }
        private void button2_Click(object sender, EventArgs e)
        {
            Сложение с = new Сложение();
            с.ShowDialog();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Вычитание v = new Вычитание();
            v.ShowDialog();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            Умножение u = new Умножение();
            u.ShowDialog();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            Транспонирование t = new Транспонирование();
            t.ShowDialog();
        }
    }
}