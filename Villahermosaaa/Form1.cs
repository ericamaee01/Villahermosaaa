﻿using Spire.Xls;
using Spire.Xls.Core;
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
using Villahermosaaa.Resources;

namespace Villahermosaaa
{
    public partial class Form1 : Form
    {
        Form2 form2;

        string[] student = new string[5];
        int i = 0;
        private string currentUserName;


        public Form1(string username)
        {
            InitializeComponent();
            currentUserName = username;

        }

        private void btnADD_Click(object sender, EventArgs e)
        {
            string name = txtName.Text.Trim();
            string gender = "";
            string hobbies = "";
            string favcolor = cboFavColor.Text.Trim();
            string address = txtAddress.Text.Trim();
            string email = txtEmail.Text.Trim();
            string birthdate = dtpBday.Text.Trim();
            string age = txtAge.Text.Trim();
            string course = cbocourse.Text.Trim();
            string saying = txtSaying.Text.Trim();
            string username = txtUsername.Text.Trim();
            string password = txtPassword.Text.Trim();
            string profilePicture = txtProfilePicture.Text.Trim();

            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show("Name cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtName.Focus();
                return;
            }
            else if (int.TryParse(name, out _))
            {
                MessageBox.Show("Name cannot be a number.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtName.Focus();
                return;
            }

            gender = radFemale.Checked ? radFemale.Text : radMale.Checked ? radMale.Text : "";
            if (string.IsNullOrEmpty(gender))
            {
                MessageBox.Show("Please select a gender.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (cbCooking.Checked) hobbies += cbCooking.Text + ", ";
            if (cbSinging.Checked) hobbies += cbSinging.Text + ", ";
            if (cbDancing.Checked) hobbies += cbDancing.Text + ", ";
            hobbies = hobbies.TrimEnd(',', ' ');

            if (string.IsNullOrEmpty(hobbies))
            {
                MessageBox.Show("Please select at least one hobby.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrEmpty(favcolor))
            {
                MessageBox.Show("Please select a favorite color.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cboFavColor.Focus();
                return;
            }

            if (string.IsNullOrEmpty(address))
            {
                MessageBox.Show("Address cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtAddress.Focus();
                return;
            }

            if (string.IsNullOrEmpty(email) || !email.Contains("@") || !email.Contains("."))
            {
                MessageBox.Show("Please enter a valid email.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtEmail.Focus();
                return;
            }

            if (string.IsNullOrEmpty(birthdate))
            {
                MessageBox.Show("Please select a birthdate.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dtpBday.Focus();
                return;
            }

            if (string.IsNullOrEmpty(age) || !int.TryParse(age, out int ageValue) || ageValue <= 0)
            {
                MessageBox.Show("Please enter a valid age.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtAge.Focus();
                return;
            }

            if (string.IsNullOrEmpty(course))
            {
                MessageBox.Show("Please select a course.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cbocourse.Focus();
                return;
            }

            if (string.IsNullOrEmpty(saying))
            {
                MessageBox.Show("Saying cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtSaying.Focus();
                return;
            }

            if (string.IsNullOrEmpty(username))
            {
                MessageBox.Show("Username cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtUsername.Focus();
                return;
            }

            if (string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Password cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtPassword.Focus();
                return;
            }

            if (string.IsNullOrEmpty(profilePicture))
            {
                MessageBox.Show("Please browse and select a profile picture.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtProfilePicture.Focus();
                return;
            }
            Workbook checkBook = new Workbook();
            checkBook.LoadFromFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx");
            Worksheet checkSheet = checkBook.Worksheets[0];

            for (int i = 2; i <= checkSheet.LastRow; i++) // skip header
            {
                string existingUsername = checkSheet.Range[i, 11].Value?.Trim();
                string existingPassword = checkSheet.Range[i, 12].Value?.Trim();

                if (string.Equals(existingUsername, username, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(existingPassword, password, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("A user with the same username and password already exists.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            //Insert data into form2
            Form2 form2 = new Form2(currentUserName);
            form2.insertdata(name, gender, hobbies, favcolor, address, email, birthdate, age, course, saying, username, password, "1", profilePicture);

            // Save to Excel
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx");
            Worksheet sh = book.Worksheets[0];
            int r = sh.LastRow + 1;

            sh.Range[r, 1].Value = name;
            sh.Range[r, 2].Value = gender;
            sh.Range[r, 3].Value = hobbies;
            sh.Range[r, 4].Value = address;
            sh.Range[r, 5].Value = favcolor;
            sh.Range[r, 6].Value = email;
            sh.Range[r, 7].Value = birthdate;
            sh.Range[r, 8].Value = age;
            sh.Range[r, 9].Value = course;
            sh.Range[r, 10].Value = saying;
            sh.Range[r, 11].Value = username;
            sh.Range[r, 12].Value = password;
            sh.Range[r, 13].Value = "1"; // active flag
            sh.Range[r, 14].Value = profilePicture; // picture path

            book.SaveToFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx", ExcelVersion.Version2016);

            MessageBox.Show("Successfully added!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);


            Mylogs logs = new Mylogs();
            logs.insertLogs(currentUserName, $"Added new user: {name}");

            // Reset form
            btnADD.Visible = true;
            btnUPDATE.Visible = false;

            txtName.Clear(); txtSaying.Clear(); txtAddress.Clear(); txtAge.Clear();
            txtEmail.Clear(); txtUsername.Clear(); txtPassword.Clear(); txtProfilePicture.Clear();

            cboFavColor.SelectedIndex = -1;
            cbocourse.SelectedIndex = -1;
            radMale.Checked = false;
            radFemale.Checked = false;
            cbCooking.Checked = false;
            cbSinging.Checked = false;
            cbDancing.Checked = false;

            txtAge.ReadOnly = true;
            txtName.Focus();

        }

        private void btnDISPLAY_Click(object sender, EventArgs e)
        {
            if (form2 == null || form2.IsDisposed)
            {
                form2 = new Form2(currentUserName);
            }
            form2.Show();
            form2.BringToFront();
        }

        private void btnUPDATE_Click(object sender, EventArgs e)
        {
            string name = txtName.Text.Trim();
            string gender = "";
            string hobbies = "";
            string favcolor = cboFavColor.Text.Trim();
            string address = txtAddress.Text.Trim();
            string email = txtEmail.Text.Trim();
            string birthdate = dtpBday.Text.Trim();
            string age = txtAge.Text.Trim();
            string course = cbocourse.Text.Trim();
            string saying = txtSaying.Text.Trim();
            string username = txtUsername.Text.Trim();
            string password = txtPassword.Text.Trim();
            string profilePicture = txtProfilePicture.Text.Trim();

            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show("Name cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtName.Focus();
                return;
            }
            else if (int.TryParse(name, out _))
            {
                MessageBox.Show("Name cannot be a number.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtName.Focus();
                return;
            }

            // Gender selection
            gender = radFemale.Checked ? radFemale.Text : radMale.Checked ? radMale.Text : "";
            if (string.IsNullOrEmpty(gender))
            {
                MessageBox.Show("Please select a gender.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Hobbies selection
            if (cbCooking.Checked) hobbies += cbCooking.Text + ", ";
            if (cbSinging.Checked) hobbies += cbSinging.Text + ", ";
            if (cbDancing.Checked) hobbies += cbDancing.Text + ", ";
            hobbies = hobbies.TrimEnd(',', ' ');
            if (string.IsNullOrEmpty(hobbies))
            {
                MessageBox.Show("Please select at least one hobby.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrEmpty(favcolor))
            {
                MessageBox.Show("Please select a favorite color.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cboFavColor.Focus();
                return;
            }

            if (string.IsNullOrEmpty(address))
            {
                MessageBox.Show("Address cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtAddress.Focus();
                return;
            }

            if (string.IsNullOrEmpty(email) || !email.Contains("@") || !email.Contains("."))
            {
                MessageBox.Show("Please enter a valid email.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtEmail.Focus();
                return;
            }

            if (string.IsNullOrEmpty(birthdate))
            {
                MessageBox.Show("Please select a birthdate.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dtpBday.Focus();
                return;
            }

            if (string.IsNullOrEmpty(age) || !int.TryParse(age, out int ageValue) || ageValue <= 0)
            {
                MessageBox.Show("Please enter a valid age.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtAge.Focus();
                return;
            }

            if (string.IsNullOrEmpty(course))
            {
                MessageBox.Show("Please select a course.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cbocourse.Focus();
                return;
            }

            if (string.IsNullOrEmpty(saying))
            {
                MessageBox.Show("Saying cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtSaying.Focus();
                return;
            }

            if (string.IsNullOrEmpty(username))
            {
                MessageBox.Show("Username cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtUsername.Focus();
                return;
            }

            if (string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Password cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtPassword.Focus();
                return;
            }

            if (string.IsNullOrEmpty(profilePicture))
            {
                MessageBox.Show("Please browse and select a profile picture.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtProfilePicture.Focus();
                return;
            }

            Workbook checkBook = new Workbook();
            checkBook.LoadFromFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx");
            Worksheet checkSheet = checkBook.Worksheets[0];

            for (int i = 2; i <= checkSheet.LastRow; i++) // skip header
            {
                string existingUsername = checkSheet.Range[i, 11].Value?.Trim();
                string existingPassword = checkSheet.Range[i, 12].Value?.Trim();

                if (string.Equals(existingUsername, username, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(existingPassword, password, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("A user with the same username and password already exists.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // Load the Excel file to update
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx");
            Worksheet sh = book.Worksheets[0];

            // Search for the existing row based on username (assuming it's in column 11)
            bool isUpdated = false;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 11].Value.ToString() == username) // Username column
                {
                    // If record is found, update it
                    sh.Range[i, 1].Value = name;
                    sh.Range[i, 2].Value = gender;
                    sh.Range[i, 3].Value = hobbies;
                    sh.Range[i, 4].Value = address;
                    sh.Range[i, 5].Value = favcolor;
                    sh.Range[i, 6].Value = email;
                    sh.Range[i, 7].Value = birthdate;
                    sh.Range[i, 8].Value = age;
                    sh.Range[i, 9].Value = course;
                    sh.Range[i, 10].Value = saying;
                    sh.Range[i, 12].Value = password;
                    sh.Range[i, 13].Value = "1"; // active flag
                    sh.Range[i, 14].Value = profilePicture; // Profile picture path

                    isUpdated = true;
                    break;
                }
            }

            // If no existing record was found, add new record
            if (!isUpdated)
            {
                int newRow = sh.LastRow + 1;
                sh.Range[newRow, 1].Value = name;
                sh.Range[newRow, 2].Value = gender;
                sh.Range[newRow, 3].Value = hobbies;
                sh.Range[newRow, 4].Value = address;
                sh.Range[newRow, 5].Value = favcolor;
                sh.Range[newRow, 6].Value = email;
                sh.Range[newRow, 7].Value = birthdate;
                sh.Range[newRow, 8].Value = age;
                sh.Range[newRow, 9].Value = course;
                sh.Range[newRow, 10].Value = saying;
                sh.Range[newRow, 11].Value = username;
                sh.Range[newRow, 12].Value = password;
                sh.Range[newRow, 13].Value = "1"; // active flag
                sh.Range[newRow, 14].Value = profilePicture; // Profile picture path
            }

            // Save changes to Excel
            book.SaveToFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx", ExcelVersion.Version2016);

            MessageBox.Show(isUpdated ? "Successfully updated!" : "Successfully added!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

            Mylogs logs = new Mylogs();
            logs.insertLogs(currentUserName, "Updated a student");


            // Reset form
            btnADD.Visible = true;
            btnUPDATE.Visible = false;

            txtName.Clear(); txtSaying.Clear(); txtAddress.Clear(); txtAge.Clear();
            txtEmail.Clear(); txtUsername.Clear(); txtPassword.Clear(); txtProfilePicture.Clear();

            cboFavColor.SelectedIndex = -1;
            cbocourse.SelectedIndex = -1;
            radMale.Checked = false;
            radFemale.Checked = false;
            cbCooking.Checked = false;
            cbSinging.Checked = false;
            cbDancing.Checked = false;

            txtAge.ReadOnly = true;
            txtName.Focus();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {

            OpenFileDialog d = new OpenFileDialog();
            if (d.ShowDialog() == DialogResult.OK)
            {
                txtProfilePicture.Text = d.FileName;

            }

            string profilePath = txtProfilePicture.Text.Trim();

        }

        private void dtpBday_ValueChanged(object sender, EventArgs e)
        {
            DateTime birthDate = DateTime.Parse(dtpBday.Text);
            int age = DateTime.Now.Year - birthDate.Year;

            if (DateTime.Now < birthDate.AddYears(age))
            {
                age--;
            }

            txtAge.Text = age.ToString();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        
    }
}





  