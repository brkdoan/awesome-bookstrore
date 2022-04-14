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

namespace awesome_bookstrore
{
    public partial class Form1 : Form
    {
        private const string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db_users.mdb";

        public Form1()
        {
            InitializeComponent();
            
        }
        OleDbConnection conn = new OleDbConnection(ConnectionString);
        OleDbCommand cmd = new OleDbCommand();
        OleDbDataAdapter da = new OleDbDataAdapter();

        private void Form1_Load(object sender, EventArgs e)
        {

        }

      

        private void btnRegister_Click(object sender, EventArgs e)
        {
            if(textUsername.Text ==""&&textPassword.Text == "" && textConfirmpsw.Text == "")
            {
                MessageBox.Show("Username and Password are empty", "Registration Faild", MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            else if(textPassword.Text == textConfirmpsw.Text)
            {
                conn.Open();
                string register = "INSERT INTO tbl_users VALUES ('" + textUsername.Text + "','" + textPassword.Text + "')";
                cmd = new OleDbCommand(register,conn);
                //cmd.Connection.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
                //cmd.Connection.Close();
                
                textUsername.Text = "";
                textPassword.Text = "";
                textConfirmpsw.Text = "";

                MessageBox.Show("Your Account has been Successfully Created","Registration Success",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Password does not match, Please Re-enter", "Registration Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textPassword.Text = "";
                textConfirmpsw.Text = "";
                textPassword.Focus();
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                textPassword.PasswordChar = '\0';
                textConfirmpsw.PasswordChar = '\0';

            }
            else
            {
                textPassword.PasswordChar = '*';
                textConfirmpsw.PasswordChar = '*';
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textUsername.Text = "";
            textPassword.Text = "";
            textConfirmpsw.Text = "";
            textUsername.Focus();
        }

        private void label6_Click(object sender, EventArgs e)
        {
            new FormLogin().Show();
            this.Hide();
        }
    }
}
