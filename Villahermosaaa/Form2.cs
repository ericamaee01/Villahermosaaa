using Spire.Xls;
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

namespace Villahermosaaa
{
    public partial class Form2 : Form
    {
        internal object txtName;

        public Form2()
        {
            InitializeComponent();
            LoadExcelFile();
        }
        public void LoadExcelFile()
        {
            Workbook book = new Workbook();
            string Filelocation = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;

            string folder = "Book1";
            string file = "Book1.xlsx";
            string path = Path.Combine(Filelocation, folder, file);
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
            int r = dataGridView1.CurrentCell.RowIndex;

            // Check if Form1 is already open
            Form1 f1 = Application.OpenForms["Form1"] as Form1;

            // Create instance if not open
            if (f1 == null)
            {
                f1 = new Form1();
                f1.Show();
            }

            // Populate Form2 with the data from the selected row in Form1's DataGridView
            if (dataGridView1.Rows[r].Cells[0].Value != null)
                f1.txtName.Text = dataGridView1.Rows[r].Cells[0].Value.ToString();

            // Gender
            string gender = dataGridView1.Rows[r].Cells[1].Value.ToString();
            f1.radMale.Checked = gender == "Male";
            f1.radFemale.Checked = gender == "Female";

            // Hobbies
            string hobbies = dataGridView1.Rows[r].Cells[2].Value.ToString();
            string[] h = hobbies.Split(',');
            f1.cbCooking.Checked = h.Contains("Cooking");
            f1.cbSinging.Checked = h.Contains("Singing");
            f1.cbDancing.Checked = h.Contains("Dancing");

            // Address
            f1.txtAddress.Text = dataGridView1.Rows[r].Cells[3].Value.ToString();

            // Favorite Color
            f1.cboFavColor.SelectedItem = dataGridView1.Rows[r].Cells[4].Value.ToString();

            // Email
            f1.txtEmail.Text = dataGridView1.Rows[r].Cells[5].Value.ToString();

            // Birthdate
            f1.dtpBday.Value = DateTime.TryParse(dataGridView1.Rows[r].Cells[6].Value.ToString(), out DateTime dob) ? dob : DateTime.Now;

            // Age
            f1.txtAge.Text = dataGridView1.Rows[r].Cells[7].Value.ToString();

            // Course
            f1.cbocourse.SelectedItem = dataGridView1.Rows[r].Cells[8].Value.ToString();

            // Saying
            f1.txtSaying.Text = dataGridView1.Rows[r].Cells[9].Value.ToString();

            // Username
            f1.txtUsername.Text = dataGridView1.Rows[r].Cells[10].Value.ToString();

            // Password
            f1.txtPassword.Text = dataGridView1.Rows[r].Cells[11].Value.ToString();

            // Profile Picture
            string picPath = dataGridView1.Rows[r].Cells[13].Value.ToString();
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

        private void btnDELETE_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                int selectedIndex = dataGridView1.SelectedRows[0].Index;
                dataGridView1.Rows.RemoveAt(selectedIndex);
                MessageBox.Show("Row deleted successfully.", "Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);

                dataGridView1.Rows[selectedIndex].Cells[12].Value = "0";
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

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}

