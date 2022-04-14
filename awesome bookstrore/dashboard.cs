using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using CsvHelper;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace awesome_bookstrore
{
    public partial class dashboard : Form
    {
        int book_number = 0;
        double price = 0;
        //private const string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db_users.mdb";
        //public dashboard()
        //{
        //    InitializeComponent();
            
        //}
        public dashboard(string username)
        {
            InitializeComponent();
            label1.Text = username;
        }

        
        //OleDbConnection conn = new OleDbConnection(ConnectionString);
        //OleDbCommand cmd = new OleDbCommand();
        //OleDbDataAdapter da = new OleDbDataAdapter();

        private void dashboard_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            //cartBindingSource.DataSource=new List<Cart>();
            //label4.Text =book_number.ToString();
            //label20.Text =price.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            book_number++;
            price = price + 22.50;
            label4.Text = book_number.ToString();
            label20.Text = price.ToString();
            dataGridView1.Rows.Add("Kozmos: Evrenin ve Yaşamın Sırları", 22.50 );
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            book_number++;
            price = price + 55.50;
            label4.Text = book_number.ToString();
            label20.Text = price.ToString();
            dataGridView1.Rows.Add("Eminim Şaka Yapıyorsunuz Bay Feynman", 55.50 );
        }

        private void button3_Click(object sender, EventArgs e)
        {
            book_number++;
            price = price + 1885.50;
            label4.Text = book_number.ToString();
            label20.Text = price.ToString();
            dataGridView1.Rows.Add("Intro to Python for Computer Science and Data Science", 1885.50 );
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        

        private void button25_Click(object sender, EventArgs e)
        {
            //using (SaveFileDialog sfd = new SaveFileDialog() {Filter = "CSV|*.csv*",ValidateNames=true })
            //{
            //    if(sfd.ShowDialog() == DialogResult.OK)
            //    {
            //        using(var sw = new StreamWriter(sfd.FileName))
            //        {
            //            var writer = new CsvWriter(sw, CultureInfo.CurrentCulture);
            //            writer.WriteHeader(typeof(Cart));
            //            foreach(Cart cart in cartBindingSource.DataSource as List<Cart>)
            //            {
            //                writer.WriteRecord(cart);
            //            }
            //        }
            //        MessageBox.Show("Your Books have been successfully saved","Message",MessageBoxButtons.OK,MessageBoxIcon.Information);
            //    }
            //}
            try
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
                Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

                int StartCol = 1;
                int StartRow = 1;

                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];
                    myRange.Value2 = dataGridView1.Columns[j].HeaderText;
                }

                StartRow++;

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];
                        myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                    }
                }
            }
            catch (Exception hata)
            {
                MessageBox.Show(hata.StackTrace);
            }
        }

        private void button26_Click(object sender, EventArgs e)
        {
            //using(OpenFileDialog ofd = new OpenFileDialog() { Filter ="CSV|*.csv*",ValidateNames =true })
            //{
            //    if(ofd.ShowDialog() == DialogResult.OK)
            //    {
            //        var sr = new StreamReader(new FileStream(ofd.FileName, FileMode.Open));
            //        var csv = new CsvReader(sr,CultureInfo.CurrentCulture);
            //        cartBindingSource.DataSource = csv.GetRecord<Cart>().ToString();
            //    }
            //}
        }

        private void button13_Click(object sender, EventArgs e)
        {
            int rowIndex = dataGridView1.CurrentCell.RowIndex;
            double price2 = 0;
            dataGridView1.Rows.RemoveAt(rowIndex);
            book_number--;
            label4.Text = book_number.ToString();
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                price2 = price2 + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
            }
            label20.Text = price2.ToString();
        }

        private void button7_Click_1(object sender, EventArgs e)
        {

        }
    }
}
