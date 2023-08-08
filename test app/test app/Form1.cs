using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace test_app
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DataGridViewComboBoxColumn comboboxcell = new DataGridViewComboBoxColumn();
            dataGridView1.RowCount = 3;
            dataGridView1.ColumnCount = 3;

            //for (int column = 0; column < dataGridView1.ColumnCount; column++)
            //{
            //    for (int row = 0; row < dataGridView1.RowCount; row++)
            //    {
            //        dataGridView1[column, row].Value = ++row;
            //        dataGridView1[column, row].Value = ++row;
            //        dataGridView1[column, row].Value = ++row;
            //    }

            //}
            comboboxcell.Items.Add("aa");
            comboboxcell.DisplayIndex = 3;

            dataGridView1.Columns.Insert(3,comboboxcell);

        }
    }
}
