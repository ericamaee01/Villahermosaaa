﻿using Spire.Xls;
using Spire.Xls.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Villahermosaaa.Resources;

namespace Villahermosaaa
{
    public partial class Form2 : Form
    {
        private string currentUserName;
        Logs logs = new Logs();

        public Form2(string userName)
        {
            InitializeComponent();
            LoadExcelFile();
            currentUserName = userName;

            dataGridView1.DefaultCellStyle.ForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.BackColor = Color.White;

            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.LightPink;
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
        }
        public void LoadExcelFile()
        {
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx");
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();
            dataGridView1.DataSource = dt;


        }
            

        
        public void insertdata(string name, string gender, string hobbies, string favColor,
                       string address, string email, string birthdate, string age,
                       string course, string saying, string username, string password,
                       string status, string profilePicture)
        {
            DataTable dt = (DataTable)dataGridView1.DataSource;
            DataRow newRow = dt.NewRow();

            newRow[0] = name;
            newRow[1] = gender;
            newRow[2] = hobbies;
            newRow[3] = favColor;
            newRow[4] = address;
            newRow[5] = email;
            newRow[6] = birthdate;
            newRow[7] = age;
            newRow[8] = course;
            newRow[9] = saying;
            newRow[10] = username;
            newRow[11] = password;
            newRow[12] = status;
            newRow[13] = profilePicture;

            dt.Rows.Add(newRow);

        }

        public void Update(string name, string gender, string hobbies, string favColor,
                      string address, string email, string birthdate, string age,
                      string course, string saying, string username, string password,
                      string status, string profilePicture)
        {
            DataTable dt = (DataTable)dataGridView1.DataSource;
            DataRow newRow = dt.NewRow();

            newRow[0] = name;
            newRow[1] = gender;
            newRow[2] = hobbies;
            newRow[3] = favColor;
            newRow[4] = address;
            newRow[5] = email;
            newRow[6] = birthdate;
            newRow[7] = age;
            newRow[8] = course;
            newRow[9] = saying;
            newRow[10] = username;
            newRow[11] = password;
            newRow[12] = status;
            newRow[13] = profilePicture;

            dt.Rows.Add(newRow);
        }


        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int r = dataGridView1.CurrentCell?.RowIndex ?? -1;
            if (r < 0 || r >= dataGridView1.Rows.Count) return;

            Form1 f1 = Application.OpenForms["Form1"] as Form1;

            if (f1 == null)
            {
                if (string.IsNullOrEmpty(currentUserName)) return;
                f1 = new Form1(currentUserName);
                f1.Show();
            }

            // Check each cell and assign safely
            var cells = dataGridView1.Rows[r].Cells;

            f1.txtName.Text = cells[0]?.Value?.ToString() ?? "";

            string gender = cells[1]?.Value?.ToString() ?? "";
            f1.radMale.Checked = gender == "Male";
            f1.radFemale.Checked = gender == "Female";

            string hobbies = cells[2]?.Value?.ToString() ?? "";
            string[] h = hobbies.Split(',');
            f1.cbCooking.Checked = h.Contains("Dancing");
            f1.cbSinging.Checked = h.Contains("Singing");
            f1.cbDancing.Checked = h.Contains("Reading");

            f1.txtAddress.Text = cells[3]?.Value?.ToString() ?? "";
            f1.cboFavColor.SelectedItem = cells[4]?.Value?.ToString() ?? "";
            f1.txtEmail.Text = cells[5]?.Value?.ToString() ?? "";

            string birthdate = cells[6]?.Value?.ToString();
            f1.dtpBday.Value = DateTime.TryParse(birthdate, out DateTime dob) ? dob : DateTime.Now;

            f1.txtAge.Text = cells[7]?.Value?.ToString() ?? "";
            f1.cbocourse.SelectedItem = cells[8]?.Value?.ToString() ?? "";
            f1.txtSaying.Text = cells[9]?.Value?.ToString() ?? "";
            f1.txtUsername.Text = cells[10]?.Value?.ToString() ?? "";
            f1.txtPassword.Text = cells[11]?.Value?.ToString() ?? "";

            string picPath = cells[13]?.Value?.ToString() ?? "";
            f1.txtProfilePicture.Text = System.IO.File.Exists(picPath) ? picPath : "";

            f1.btnADD.Visible = false;
            f1.btnUPDATE.Visible = true;
        }



        private void btnDELETE_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                int selectedIndex = dataGridView1.SelectedRows[0].Index;

                // Update status in DataGridView
                dataGridView1.Rows[selectedIndex].Cells[12].Value = "0";

                // Load the Excel file
                Workbook book = new Workbook();
                book.LoadFromFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx");
                Worksheet sheet = book.Worksheets[0];


                string username = dataGridView1.Rows[selectedIndex].Cells[10].Value.ToString();

                for (int i = 2; i <= sheet.LastRow; i++)
                {
                    if (sheet.Range[i, 11].Value == username)
                    {
                        sheet.Range[i, 13].Value = "0";
                        break;
                    }
                }

                // Save changes
                book.SaveToFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx");

                MessageBox.Show("Deleted. Status marked as '0'", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                Mylogs logs = new Mylogs();
                logs.insertLogs(currentUserName, "Deleted a student");
            }
            else
            {
                MessageBox.Show("Please select a row to delete.", "Delete Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {

            dataGridView1.ClearSelection();
            bool itemFound = false;

            try
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[0].Value.ToString().Equals(txtSearch.Text, StringComparison.OrdinalIgnoreCase))
                    {
                        row.Selected = true;
                        itemFound = true;
                        break;
                    }
                }

                if (!itemFound)
                {
                    throw new Exception("Item was not in the list.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Search not Found: " + ex.Message);
            }
            finally
            {
                txtSearch.Clear();
            }
        }

        private void btnCLOSE_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {

        }

       
    }
}

