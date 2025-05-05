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
        public Dashboard(string name, string path)
        {
            InitializeComponent();
            lblName.Text = "Welcome!, " + name;

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

            // Load the Excel file to count active students
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\ACT-STUDENT\\Downloads\\book1.xlsx");
            Worksheet sh = book.Worksheets[0];

            int activeStudentCount = 0;

            // Loop through the rows and check for active status (column 13 holds the active status)
            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                // If the status in column 13 is "1" (active)
                if (sh.Range[i, 13].Value.ToString() == "1")
                {
                    activeStudentCount++;
                    lblActive.Text = activeStudentCount.ToString();
                }


            }
            int inactiveStudentCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                // If the status in column 13 is "0" (inactive)
                if (sh.Range[i, 13].Value.ToString() == "0")
                {
                    inactiveStudentCount++;
                    lblInactive.Text = inactiveStudentCount.ToString();
                }

            }
            int maleGenderCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 2].Value.ToString() == "Male")
                {
                    maleGenderCount++;
                    lblMale.Text = maleGenderCount.ToString();
                }

            }
            int FemaleGenderCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 2].Value.ToString() == "Female")
                {
                    FemaleGenderCount++;
                    lblFemale.Text = FemaleGenderCount.ToString();
                }

            }
            int CookingHobbiesCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 3].Value.ToString() == "Dancing")
                {
                    CookingHobbiesCount++;
                    lblCooking.Text = CookingHobbiesCount.ToString();
                }

            }
            int singingHobbiesCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 3].Value.ToString() == "Singing")
                {
                    singingHobbiesCount++;
                    lblSinging.Text = singingHobbiesCount.ToString();
                }

            }
            int DancingHobbiesCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 3].Value.ToString() == "Reading")
                {
                    DancingHobbiesCount++;
                    lblDancing.Text = DancingHobbiesCount.ToString();
                }

            }
            int pinkColorCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 5].Value.ToString() == "Pink")
                {
                    pinkColorCount++;
                    lblPink.Text = pinkColorCount.ToString();
                }

            }
            int blackColorCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 5].Value.ToString() == "Black")
                {
                    blackColorCount++;
                    lblBlack.Text = blackColorCount.ToString();
                }

            }
            int purpleColorCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 5].Value.ToString() == "White")
                {
                    purpleColorCount++;
                    lblPurple.Text = purpleColorCount.ToString();
                }

            }
            int BsitCourseCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 9].Value.ToString() == "BSIT")
                {
                    BsitCourseCount++;
                    lblBSIT.Text = BsitCourseCount.ToString();
                }

            }
            int bsedCourseCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 9].Value.ToString() == "BSED")
                {
                    bsedCourseCount++;
                    BSED.Text = bsedCourseCount.ToString();
                }

            }
            int bsbaCourseCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 9].Value.ToString() == "BSBA")
                {
                    bsbaCourseCount++;
                    BSBA.Text = bsbaCourseCount.ToString();
                }
            }
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
                // Show login form
                Login loginForm = new Login();
                loginForm.Show();

                // Close or hide main form
                this.Hide(); // or this.Close();

                if (result != DialogResult.Yes)
                {
                    this.Close();
                }
            }

        }

        private void btndashboard_Click(object sender, EventArgs e)
        {
            loadform(new home ());
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void label_Click(object sender, EventArgs e)
        {

        }
    }
}
        