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

namespace Villahermosaaa.Resources
{
    public partial class Dashboard : Form
    {
        private string currentUserName;
        public Dashboard(string name, string path)
        {
            
            InitializeComponent();
            currentUserName = name;

            // Load the Excel file
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx");
            Worksheet sh = book.Worksheets[0];

            lblName.Text = "Welcome! " + name;

            try
            {
                pictureBox1.Image = Image.FromFile(path);
                pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            }
            catch (Exception ex)
            {
                pictureBox1.Image = null;
                MessageBox.Show("Error loading profile picture:\n" + ex.Message, "Image Error");
            }



            // Count active and inactive students
            int activeStudentCount = 0;
            int inactiveStudentCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                string status = sh.Range[i, 13].Value?.ToString().Trim();
                if (status == "1") activeStudentCount++;
                else if (status == "0") inactiveStudentCount++;
            }
            lblActive.Text = activeStudentCount.ToString();
            lblInactive.Text = inactiveStudentCount.ToString();

            // Count the male students
            int maleGenderCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 2].Value.ToString() == "Male")
                {
                    maleGenderCount++;
                    lblMale.Text = maleGenderCount.ToString();
                }
            }

            // Count the female students
            int femaleGenderCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 2].Value.ToString() == "Female")
                {
                    femaleGenderCount++;
                    lblFemale.Text = femaleGenderCount.ToString();
                }
            }

            // Count hobbies
            int CookingHobbiesCount = 0;
            int SingingHobbiesCount = 0;
            int DancingHobbiesCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                string hobby = sh.Range[i, 3].Value.ToString();
                if (hobby == "Cooking") CookingHobbiesCount++;
                if (hobby == "Singing") SingingHobbiesCount++;
                if (hobby == "Dancing") DancingHobbiesCount++;
            }
            lblCooking.Text = CookingHobbiesCount.ToString();
            lblSinging.Text = SingingHobbiesCount.ToString();
            lblDancing.Text = DancingHobbiesCount.ToString();

            // Count favorite colors
            int BlackColorCount = 0;
            int PinkColorCount = 0;
            int PurpleColorCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                string color = sh.Range[i, 5].Value.ToString();
                if (color == "Black") BlackColorCount++;
                if (color == "Pink") PinkColorCount++;
                if (color == "Purple") PurpleColorCount++;
            }
            lblBlack.Text = BlackColorCount.ToString();
            lblPink.Text = PinkColorCount.ToString();
            lblPurple.Text = PurpleColorCount.ToString();

            // Count courses
            int bsitCourseCount = 0;
            int bsbaCourseCount = 0;
            int bsedCourseCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                string course = sh.Range[i, 9].Value.ToString();
                if (course == "BSIT") bsitCourseCount++;
                if (course == "BSED") bsedCourseCount++;
                if (course == "BSBA") bsbaCourseCount++;
            }
            lblBSIT.Text = bsitCourseCount.ToString();
            lblBSBA.Text = bsedCourseCount.ToString();
            lblBSED.Text = bsbaCourseCount.ToString();


        }



        public void loadform(object Form)
        {
            // create instance
            Form form = Form as Form;
            //configure the form to be loaded
            form.TopLevel = false;
            form.Dock = DockStyle.Fill;

            // Clear existing controls
            this.mainpanel.Controls.Clear();

            //add the form to the panel and display it
            this.mainpanel.Controls.Add(form);
            this.mainpanel.Tag = form;
            form.Show();
        }



        private void btnLogout_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Are you sure you want to log out?", "Confirm Logout", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                Mylogs logs = new Mylogs();
                logs.insertLogs(currentUserName, "Successfully logged out!");

                this.Close();
                Login login = new Login();
                login.ShowDialog();

                if (result == DialogResult.No)
                {
                    return;
                }
            }
        }

      
        private void btnActive_Click(object sender, EventArgs e)
        {
            Active active = new Active(currentUserName);

            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx");
            Worksheet sheet = book.Worksheets[0];

            DataTable dt = sheet.ExportDataTable();
            DataTable filtered = dt.Clone();

            foreach (DataRow dr in dt.Rows)
            {
                if (dr[12].ToString().Trim() == "1")
                {
                    filtered.ImportRow(dr);
                }
            }

            active.dataGridView1.DataSource = filtered;
            active.dataGridView1.Refresh();

            active.dataGridView1.DefaultCellStyle.ForeColor = Color.Black;
            active.dataGridView1.DefaultCellStyle.BackColor = Color.White;

            active.dataGridView1.DefaultCellStyle.SelectionBackColor = Color.LightPink;
            active.dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;

            loadform(active);
        }
    

        private void btnInactive_Click(object sender, EventArgs e)
        {
            // Prepare the form instance
            Inactive inactive = new Inactive(currentUserName);

            // Load Excel data
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx");
            Worksheet sheet = book.Worksheets[0];

            // Export and filter data
            DataTable dt = sheet.ExportDataTable();
            DataTable filtered = dt.Clone();

            foreach (DataRow dr in dt.Rows)
            {
                if (dr[12].ToString().Trim() == "0")
                {
                    filtered.ImportRow(dr);
                }
            }
            inactive.dataGridView2.DataSource = filtered;

            inactive.dataGridView2.DefaultCellStyle.ForeColor = Color.Black;
            inactive.dataGridView2.DefaultCellStyle.BackColor = Color.White;

            inactive.dataGridView2.DefaultCellStyle.SelectionBackColor = Color.LightPink;
            inactive.dataGridView2.DefaultCellStyle.SelectionForeColor = Color.Black;
            loadform(inactive);
        }

        private void btnLogs_Click(object sender, EventArgs e)
        {
            Mylogs logs = new Mylogs();

            Logs logsForm = new Logs();

            // Load Excel file
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx");
            Worksheet sheet = book.Worksheets[1]; // Sheet2 for logs

            // Export and filter data
            DataTable dt = sheet.ExportDataTable();
            DataTable filtered = dt.Clone();

            foreach (DataRow dr in dt.Rows)
            {
                // No filtering needed here, just copy all rows for logs
                filtered.ImportRow(dr);
            }

            logsForm.dataGridView2.DataSource = filtered;

            logsForm.dataGridView2.DefaultCellStyle.ForeColor = Color.Black;
            logsForm.dataGridView2.DefaultCellStyle.BackColor = Color.White;

            logsForm.dataGridView2.DefaultCellStyle.SelectionBackColor = Color.LightPink;
            logsForm.dataGridView2.DefaultCellStyle.SelectionForeColor = Color.Black;
            loadform(logsForm);
        }

        private void btndashboard_Click(object sender, EventArgs e)
        {

            loadform(new home());

            // Load the Excel file
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx");
            Worksheet sh = book.Worksheets[0];

            // Count active and inactive students
            int activeStudentCount = 0;
            int inactiveStudentCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                string status = sh.Range[i, 13].Value?.ToString().Trim();
                if (status == "1") activeStudentCount++;
                else if (status == "0") inactiveStudentCount++;
            }
            lblActive.Text = activeStudentCount.ToString();
            lblInactive.Text = inactiveStudentCount.ToString();

            // Count the male students
            int maleGenderCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 2].Value.ToString() == "Male")
                {
                    maleGenderCount++;
                    lblMale.Text = maleGenderCount.ToString();
                }
            }

            // Count the female students
            int femaleGenderCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 2].Value.ToString() == "Female")
                {
                    femaleGenderCount++;
                    lblFemale.Text = femaleGenderCount.ToString();
                }
            }

            // Count hobbies
            int CookingHobbiesCount = 0;
            int SingingHobbiesCount = 0;
            int DancingHobbiesCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                string hobby = sh.Range[i, 3].Value.ToString();
                if (hobby == "Cooking") CookingHobbiesCount++;
                if (hobby == "Singing") SingingHobbiesCount++;
                if (hobby == "Dancing") DancingHobbiesCount++;
            }
            lblCooking.Text = CookingHobbiesCount.ToString();
            lblSinging.Text = SingingHobbiesCount.ToString();
            lblDancing.Text = DancingHobbiesCount.ToString();

            // Count favorite colors
            int BlackColorCount = 0;
            int PinkColorCount = 0;
            int PurpleColorCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                string color = sh.Range[i, 5].Value.ToString();
                if (color == "Pink") BlackColorCount++;
                if (color == "Black") PinkColorCount++;
                if (color == "White") PurpleColorCount++;
            }
            lblBlack.Text = BlackColorCount.ToString();
            lblPink.Text = PinkColorCount.ToString();
            lblPurple.Text = PurpleColorCount.ToString();

            // Count courses
            int bsitCourseCount = 0;
            int bsedCourseCount = 0;
            int bsbaCourseCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                string course = sh.Range[i, 9].Value.ToString();
                if (course == "BSIT") bsitCourseCount++;
                if (course == "BSED") bsedCourseCount++;
                if (course == "BSBA") bsbaCourseCount++;
            }
            lblBSIT.Text = bsitCourseCount.ToString();
            lblBSED.Text = bsedCourseCount.ToString();
            lblBSBA.Text = bsbaCourseCount.ToString();

        }

    }

}
        