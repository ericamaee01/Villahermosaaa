using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Villahermosaaa.Resources;

namespace Villahermosaaa
{
    public partial class Login : Form
    {
       
       
        public Login()
        {
            InitializeComponent();
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\ACT-STUDENT\\Downloads\\book1.xlsx");
            Worksheet sheet = book.Worksheets[0];
            bool loginSuccess = false;

            for (int i = 2; i <= sheet.LastRow; i++) // Skip header row
            {
                string storedUsername = sheet.Range[i, 11].Value?.Trim();
                string storedPassword = sheet.Range[i, 12].Value?.Trim();

                if (storedUsername == txtUsername.Text.Trim() && storedPassword == txtPassword.Text.Trim())
                {
                    string profilePath = sheet.Range[i, 14].Value;
                    string name = txtUsername.Text.Trim();


                   
                    MessageBox.Show("Login successful");

                    this.Hide();
                    Dashboard dashboard = new Dashboard(name, profilePath); //  Now it matches your constructor
                    dashboard.ShowDialog();
                    this.Show();


                }

            }
            if (!loginSuccess)
            {
                // Validate username and password fields
                if (string.IsNullOrWhiteSpace(txtPassword.Text) && string.IsNullOrWhiteSpace(txtUsername.Text))
                {
                    MessageBox.Show("Username and password cannot be empty.");
                }
                else if (string.IsNullOrWhiteSpace(txtUsername.Text))
                {
                    MessageBox.Show("Username cannot be empty.");
                }
                else if (string.IsNullOrWhiteSpace(txtPassword.Text))
                {
                    MessageBox.Show("Password cannot be empty.");
                }
                else
                {
                    MessageBox.Show("Invalid username or password.", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }


        }

        private void btnCLEAR_Click(object sender, EventArgs e)
        {
            txtUsername.Clear();
            txtPassword.Clear();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox1.Checked)
            {
                txtPassword.UseSystemPasswordChar = true;
            }
            else
            {
                txtPassword.UseSystemPasswordChar = false;
            }
        }
    }
}
