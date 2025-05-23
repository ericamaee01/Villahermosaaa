﻿using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Villahermosaaa.Resources
{
    public partial class Inactive : Form
    {
        private string currentUserName;
        public Inactive(string userName)
        {
            InitializeComponent();
            currentUserName = userName;
        }

        private void btnDELETE_Click(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count > 0)
            {
                int selectedIndex = dataGridView2.SelectedRows[0].Index;

                // Update status in DataGridView
                dataGridView2.Rows[selectedIndex].Cells[12].Value = "1";

                // Load the Excel file
                Workbook book = new Workbook();
                book.LoadFromFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx");
                Worksheet sheet = book.Worksheets[0];


                string username = dataGridView2.Rows[selectedIndex].Cells[10].Value.ToString();

                for (int i = 2; i <= sheet.LastRow; i++)
                {
                    if (sheet.Range[i, 11].Value == username)
                    {
                        sheet.Range[i, 13].Value = "1";
                        break;
                    }
                }


                // Save changes
                book.SaveToFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx");

                MessageBox.Show("User activated. Status marked as '1'", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                Mylogs logs = new Mylogs();
                logs.insertLogs(currentUserName, "Deleted an inactive user");
            }
            else
            {
                MessageBox.Show("Please select a row to delete.", "Delete Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            dataGridView2.ClearSelection();
            bool itemFound = false;

            try
            {
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    // Skip new row placeholder
                    if (row.IsNewRow) continue;

                    if (row.Cells[0].Value != null &&
                        row.Cells[0].Value.ToString().IndexOf(txtSearch.Text.Trim(), StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        row.Selected = true;
                        // Optional: Scroll to the first match only
                        if (!itemFound)
                        {
                            dataGridView2.FirstDisplayedScrollingRowIndex = row.Index;
                        }
                        itemFound = true;
                    }
                }

                if (!itemFound)
                {
                    MessageBox.Show("Item was not found in the list. Please try again.", "Search Failed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred during search: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int r = dataGridView2.CurrentCell.RowIndex;

            // Check if Form1 is already open
            Form1 f1 = Application.OpenForms["Form1"] as Form1;

            // Create instance if not open
            if (f1 == null)
            {
                f1 = new Form1(currentUserName);
                f1.Show();
            }

            // Populate Form2 with the data from the selected row in Form1's DataGridView
            if (dataGridView2.Rows[r].Cells[0].Value != null)
                f1.txtName.Text = dataGridView2.Rows[r].Cells[0].Value.ToString();

            // Gender
            string gender = dataGridView2.Rows[r].Cells[1].Value.ToString();
            f1.radMale.Checked = gender == "Male";
            f1.radFemale.Checked = gender == "Female";

            // Hobbies
            string hobbies = dataGridView2.Rows[r].Cells[2].Value.ToString();
            string[] h = hobbies.Split(',');
            f1.cbCooking.Checked = h.Contains("Cooking");
            f1.cbSinging.Checked = h.Contains("Singing");
            f1.cbDancing.Checked = h.Contains("Dancing");

            // Address
            f1.txtAddress.Text = dataGridView2.Rows[r].Cells[3].Value.ToString();

            // Favorite Color
            f1.cboFavColor.SelectedItem = dataGridView2.Rows[r].Cells[4].Value.ToString();

            // Email
            f1.txtEmail.Text = dataGridView2.Rows[r].Cells[5].Value.ToString();

            // Birthdate
            f1.dtpBday.Value = DateTime.TryParse(dataGridView2.Rows[r].Cells[6].Value.ToString(), out DateTime dob) ? dob : DateTime.Now;

            // Age
            f1.txtAge.Text = dataGridView2.Rows[r].Cells[7].Value.ToString();

            // Course
            f1.cbocourse.SelectedItem = dataGridView2.Rows[r].Cells[8].Value.ToString();

            // Saying
            f1.txtSaying.Text = dataGridView2.Rows[r].Cells[9].Value.ToString();

            // Username
            f1.txtUsername.Text = dataGridView2.Rows[r].Cells[10].Value.ToString();

            // Password
            f1.txtPassword.Text = dataGridView2.Rows[r].Cells[11].Value.ToString();

            // Profile Picture
            string picPath = dataGridView2.Rows[r].Cells[13].Value.ToString();
            if (System.IO.File.Exists(picPath))
            {
                f1.txtProfilePicture.Text = picPath;
            }
            else
            {
                f1.txtProfilePicture.Text = "";
            }

            // Hide ADD button, show UPDATE button
            f1.btnADD.Visible = false;
            f1.btnUPDATE.Visible = true;
        }
    }
}
