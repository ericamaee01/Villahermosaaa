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

namespace Villahermosaaa.Resources
{
    public partial class home : Form
    {
        public home()
        {
            InitializeComponent();

            // Load the Excel file to count active students
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx");
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
            int femaleGenderCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 2].Value.ToString() == "Female")
                {
                    femaleGenderCount++;
                    lblFemale.Text = femaleGenderCount.ToString();
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
            int SingingHobbiesCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 3].Value.ToString() == "Singing")
                {
                    SingingHobbiesCount++;
                    lblSinging.Text = SingingHobbiesCount.ToString();
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
            int BlackColorCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 5].Value.ToString() == "Pink")
                {
                    BlackColorCount++;
                    lblBlack.Text = BlackColorCount.ToString();
                }

            }
            int PinkColorCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 5].Value.ToString() == "Black")
                {
                    PinkColorCount++;
                    lblPink.Text = PinkColorCount.ToString();
                }

            }
            int PurpleColorCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 5].Value.ToString() == "White")
                {
                    PurpleColorCount++;
                    lblPurple.Text = PurpleColorCount.ToString();
                }

            }
            int bsitCourseCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 9].Value.ToString() == "BSIT")
                {
                    bsitCourseCount++;
                    lblBSIT.Text = bsitCourseCount.ToString();
                }

            }
            int bsedCourseCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 9].Value.ToString() == "BSED")
                {
                    bsedCourseCount++;
                    lblBSED.Text = bsedCourseCount.ToString();
                }

            }
            int bsbaCourseCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 9].Value.ToString() == "BSBA")
                {
                    bsbaCourseCount++;
                    lblBSBA.Text = bsbaCourseCount.ToString();
                }
            }
        }

        private void home_Load(object sender, EventArgs e)
        {

        }

        private void lblDancing_Click(object sender, EventArgs e)
        {

        }
    }

}