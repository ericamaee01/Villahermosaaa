using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Villahermosaaa
{
    public partial class Form1 : Form
    {
        Form2 form2;

        string[] student = new string[5];
        int i = 0;


        public Form1()
        {
            InitializeComponent();
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

            if (radFemale.Checked)
                gender = radFemale.Text.Trim();
            else if (radMale.Checked)
                gender = radMale.Text.Trim();

            if (string.IsNullOrEmpty(gender))
            {
                MessageBox.Show("Please select a gender.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (cbCooking.Checked) hobbies += cbCooking.Text.Trim() + ", ";
            if (cbSinging.Checked) hobbies += cbSinging.Text.Trim() + ", ";
            if (cbDancing.Checked) hobbies += cbDancing.Text.Trim() + ", ";
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

            i++;

            Form2 form2 = new Form2();
            form2.insertdata(name, gender, hobbies, favcolor, address, email, birthdate, age, course, saying, username, password, "1", profilePicture);

            // Save to Excel
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\ACT-STUDENT\\Downloads\\book1.xlsx");
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
            sh.Range[r, 13].Value = "1";
            sh.Range[r, 14].Value = profilePicture;

            // ✅ Proper save
            book.SaveToFile("C:\\Users\\ACT-STUDENT\\Downloads\\book1.xlsx", ExcelVersion.Version2016);

            MessageBox.Show("Successfully added!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

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
                form2 = new Form2();
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

            if (radFemale.Checked)
                gender = radFemale.Text.Trim();
            else if (radMale.Checked)
                gender = radMale.Text.Trim();

            if (string.IsNullOrEmpty(gender))
            {
                MessageBox.Show("Please select a gender.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (cbCooking.Checked) hobbies += cbCooking.Text.Trim() + ", ";
            if (cbSinging.Checked) hobbies += cbSinging.Text.Trim() + ", ";
            if (cbDancing.Checked) hobbies += cbDancing.Text.Trim() + ", ";
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

            i++;

            Form2 form2 = new Form2();
            form2.Update(name, gender, hobbies, favcolor, address, email, birthdate, age, course, saying, username, password, "1", profilePicture);

            // Save to Excel
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\ACT-STUDENT\\Downloads\\book1.xlsx");
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
            sh.Range[r, 13].Value = "1";
            sh.Range[r, 14].Value = profilePicture;

            // ✅ Proper save
            book.SaveToFile("C:\\Users\\ACT-STUDENT\\Downloads\\book1.xlsx", ExcelVersion.Version2016);

            MessageBox.Show("Successfully updated!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

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





  