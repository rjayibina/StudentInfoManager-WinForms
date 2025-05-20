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
using Spire.Xls;

namespace EVEDRI_Lab_Act2
{
    public partial class Dashboard: Form
    {
        Workbook book = new Workbook();

        private Form2 f2;

        private string _userName;

        private Form1 _form1;

        Logs logs = new Logs();

        public Dashboard(string userName, Form1 form1)
        {
            InitializeComponent();
            _userName = userName;
            _form1 = form1;

            lblActiveCount.Text = showCount(13, "Active").ToString();
            lblInactiveCount.Text = showCount(13, "Inactive").ToString();
            lblMaleCount.Text = showCount(2, "Male").ToString();
            lblFemaleCount.Text = showCount(2, "Female").ToString();
            lblBasketballCount.Text = showCount(3, "Basketball").ToString();
            lblVolleyballCount.Text = showCount(3, "Volleyball").ToString();
            lblSoccerCount.Text = showCount(3, "Soccer").ToString();
            lblRedCount.Text = showCount(8, "Red").ToString();
            lblBlueCount.Text = showCount(8, "Blue").ToString();
            lblBlackCount.Text = showCount(8, "Black").ToString();
            lblBSITCount.Text = showCount(10, "BSIT").ToString();
            lblBSCpECount.Text = showCount(10, "BSCpE").ToString();
            lblBSCSCount.Text = showCount(10, "BSCS").ToString();

            btnActive.MouseEnter += btnActive_MouseEnter;
            btnActive.MouseLeave += btnActive_MouseLeave;
            btnInactive.MouseEnter += btnInactive_MouseEnter;
            btnInactive.MouseLeave += btnInactive_MouseLeave;
            btnLogs.MouseEnter += btnLogs_MouseEnter;
            btnLogs.MouseLeave += btnLogs_MouseLeave;
            btnLogout.MouseEnter += btnLogout_MouseEnter;
            btnLogout.MouseLeave += btnLogout_MouseLeave;
            btnAdd.MouseEnter += btnAdd_MouseEnter;
            btnAdd.MouseLeave += btnAdd_MouseLeave;
            btnViewAll.MouseEnter += btnViewAll_MouseEnter;
            btnViewAll.MouseLeave += btnViewAll_MouseLeave;

            panelActive.Paint += panelActive_Paint;
            panelActive.MouseEnter += panelActive_MouseEnter;
            panelActive.MouseLeave += panelActive_MouseLeave;
            panelInactive.Paint += panelInactive_Paint;
            panelInactive.MouseEnter += panelInactive_MouseEnter;
            panelInactive.MouseLeave += panelInactive_MouseLeave;
            panelGender.Paint += panelGender_Paint;
            panelGender.MouseEnter += panelGender_MouseEnter;
            panelGender.MouseLeave += panelGender_MouseLeave;
            panelHobbies.MouseEnter += panelHobbies_MouseEnter;
            panelHobbies.MouseLeave += panelHobbies_MouseLeave;
            panelHobbies.Paint += panelHobbies_Paint;
            panelColor.Paint += panelColor_Paint;
            panelColor.MouseEnter += panelColor_MouseEnter;
            panelColor.MouseLeave += panelColor_MouseLeave;
            panelProgram.Paint += panelProgram_Paint;
            panelProgram.MouseEnter += panelProgram_MouseEnter;
            panelProgram.MouseLeave += panelProgram_MouseLeave;
        }

        private void panelProgram_Paint(object sender, PaintEventArgs e)
        {
            Color borderColor = Color.FromArgb(21, 112, 239);
            int borderWidth = 1;

            using (Pen borderPen = new Pen(borderColor, borderWidth))
            {
                Rectangle rect = new Rectangle(0, 0, panelProgram.Width - 1, panelProgram.Height - 1);
                e.Graphics.DrawRectangle(borderPen, rect);
            }
        }

        private void panelProgram_MouseEnter(object sender, EventArgs e)
        {
            panelProgram.BackColor = Color.FromArgb(21, 112, 239);

            lblProgram.ForeColor = Color.White;
            lblBSIT.ForeColor = Color.White;
            lblBSCpE.ForeColor = Color.White;
            lblBSCS.ForeColor = Color.White;

            lblBSITCount.ForeColor = Color.White;
            lblBSCpECount.ForeColor = Color.White;
            lblBSCSCount.ForeColor = Color.White;

            lblProgram.BackColor = Color.Transparent;
            lblBSIT.BackColor = Color.Transparent;
            lblBSCpE.BackColor = Color.Transparent;
            lblBSCS.BackColor = Color.Transparent;

            lblBSITCount.BackColor = Color.Transparent;
            lblBSCpECount.BackColor = Color.Transparent;
            lblBSCSCount.BackColor = Color.Transparent;
        }

        private void panelProgram_MouseLeave(object sender, EventArgs e)
        {
            panelProgram.BackColor = Color.Transparent;

            lblProgram.ForeColor = Color.Black;
            lblBSIT.ForeColor = Color.Black;
            lblBSCpE.ForeColor = Color.Black;
            lblBSCS.ForeColor = Color.Black;

            lblBSITCount.ForeColor = Color.Black;
            lblBSCpECount.ForeColor = Color.Black;
            lblBSCSCount.ForeColor = Color.Black;

            lblProgram.BackColor = Color.Transparent;
            lblBSIT.BackColor = Color.Transparent;
            lblBSCpE.BackColor = Color.Transparent;
            lblBSCS.BackColor = Color.Transparent;

            lblBSITCount.BackColor = Color.Transparent;
            lblBSCpECount.BackColor = Color.Transparent;
            lblBSCSCount.BackColor = Color.Transparent;
        }

        private void panelColor_MouseEnter(object sender, EventArgs e)
        {
            panelColor.BackColor = Color.FromArgb(21, 112, 239);

            lblColor.ForeColor = Color.White;
            lblRed.ForeColor = Color.White;
            lblBlue.ForeColor = Color.White;
            lblBlack.ForeColor = Color.White;

            lblRedCount.ForeColor = Color.White;
            lblBlueCount.ForeColor = Color.White;
            lblBlackCount.ForeColor = Color.White;

            // Optional: Ensure labels remain transparent
            lblColor.BackColor = Color.Transparent;
            lblRed.BackColor = Color.Transparent;
            lblBlue.BackColor = Color.Transparent;
            lblBlack.BackColor = Color.Transparent;
            lblRedCount.BackColor = Color.Transparent;
            lblBlueCount.BackColor = Color.Transparent;
            lblBlackCount.BackColor = Color.Transparent;
        }

        private void panelColor_MouseLeave(object sender, EventArgs e)
        {
            panelColor.BackColor = Color.Transparent;

            lblColor.ForeColor = Color.Black;
            lblRed.ForeColor = Color.Black;
            lblBlue.ForeColor = Color.Black;
            lblBlack.ForeColor = Color.Black;

            lblRedCount.ForeColor = Color.Black;
            lblBlueCount.ForeColor = Color.Black;
            lblBlackCount.ForeColor = Color.Black;

            // Optional: Ensure labels remain transparent
            lblColor.BackColor = Color.Transparent;
            lblRed.BackColor = Color.Transparent;
            lblBlue.BackColor = Color.Transparent;
            lblBlack.BackColor = Color.Transparent;
            lblRedCount.BackColor = Color.Transparent;
            lblBlueCount.BackColor = Color.Transparent;
            lblBlackCount.BackColor = Color.Transparent;
        }


        private void panelColor_Paint(object sender, PaintEventArgs e)
        {
            Color borderColor = Color.FromArgb(21, 112, 239);
            int borderWidth = 1;

            using (Pen borderPen = new Pen(borderColor, borderWidth))
            {
                Rectangle rect = new Rectangle(0, 0, panelColor.Width - 1, panelColor.Height - 1);
                e.Graphics.DrawRectangle(borderPen, rect);
            }
        }


        private void panelHobbies_Paint(object sender, PaintEventArgs e)
        {
            Color borderColor = Color.FromArgb(21, 112, 239);
            int borderWidth = 1;

            using (Pen borderPen = new Pen(borderColor, borderWidth))
            {
                // Draw rectangle border 1px inside to avoid clipping
                Rectangle rect = new Rectangle(0, 0, panelHobbies.Width - 1, panelHobbies.Height - 1);
                e.Graphics.DrawRectangle(borderPen, rect);
            }
        }

        private void panelHobbies_MouseEnter(object sender, EventArgs e)
        {
            Color highlight = Color.FromArgb(21, 112, 239);
            panelHobbies.BackColor = highlight;

            lblHobbies.ForeColor = Color.White;
            lblBasketball.ForeColor = Color.White;
            lblBasketballCount.ForeColor = Color.White;
            lblVolleyball.ForeColor = Color.White;
            lblVolleyballCount.ForeColor = Color.White;
            lblSoccer.ForeColor = Color.White;
            lblSoccerCount.ForeColor = Color.White;

            panelHobbies.Invalidate();
        }

        private void panelHobbies_MouseLeave(object sender, EventArgs e)
        {
            panelHobbies.BackColor = Color.Transparent;

            lblHobbies.ForeColor = SystemColors.ControlText;
            lblBasketball.ForeColor = SystemColors.ControlText;
            lblBasketballCount.ForeColor = SystemColors.ControlText;
            lblVolleyball.ForeColor = SystemColors.ControlText;
            lblVolleyballCount.ForeColor = SystemColors.ControlText;
            lblSoccer.ForeColor = SystemColors.ControlText;
            lblSoccerCount.ForeColor = SystemColors.ControlText;

            panelHobbies.Invalidate();
        }


        private void panelGender_MouseEnter(object sender, EventArgs e)
        {
            Color highlight = Color.FromArgb(21, 112, 239);
            panelGender.BackColor = highlight;

            lblGender.ForeColor = Color.White;
            lblMale.ForeColor = Color.White;
            lblFemale.ForeColor = Color.White;
            lblMaleCount.ForeColor = Color.White;
            lblFemaleCount.ForeColor = Color.White;

            lblGender.Invalidate();
            lblMale.Invalidate();
            lblFemale.Invalidate();
            lblMaleCount.Invalidate();
            lblFemaleCount.Invalidate();
        }

        private void panelGender_MouseLeave(object sender, EventArgs e)
        {
            panelGender.BackColor = Color.Transparent;

            lblGender.ForeColor = SystemColors.ControlText;
            lblMale.ForeColor = SystemColors.ControlText;
            lblFemale.ForeColor = SystemColors.ControlText;
            lblMaleCount.ForeColor = SystemColors.ControlText;
            lblFemaleCount.ForeColor = SystemColors.ControlText;

            lblGender.Invalidate();
            lblMale.Invalidate();
            lblFemale.Invalidate();
            lblMaleCount.Invalidate();
            lblFemaleCount.Invalidate();
        }

        private void panelActive_Paint(object sender, PaintEventArgs e)
        {
            Color borderColor = Color.FromArgb(21, 112, 239);
            int borderWidth = 1;

            using (Pen borderPen = new Pen(borderColor, borderWidth))
            {
                // Draw rectangle border 1px inside to avoid clipping
                Rectangle rect = new Rectangle(0, 0, panelActive.Width - 1, panelActive.Height - 1);
                e.Graphics.DrawRectangle(borderPen, rect);
            }
        }

        private void panelGender_Paint(object sender, PaintEventArgs e)
        {
            Color borderColor = Color.FromArgb(21, 112, 239);
            int borderWidth = 1;

            using (Pen borderPen = new Pen(borderColor, borderWidth))
            {
                // Draw rectangle border 1px inside to avoid clipping
                Rectangle rect = new Rectangle(0, 0, panelActive.Width - 1, panelActive.Height - 1);
                e.Graphics.DrawRectangle(borderPen, rect);
            }
        }

        private void panelActive_MouseEnter(object sender, EventArgs e)
        {
            Color highlight = Color.FromArgb(21, 112, 239);
            panelActive.BackColor = highlight;

            lblActiveStud.ForeColor = Color.White;
            lblActiveCount.ForeColor = Color.White;

            lblActiveStud.Invalidate();   // Force repaint
            lblActiveCount.Invalidate();
        }

        private void panelActive_MouseLeave(object sender, EventArgs e)
        {
            panelActive.BackColor = Color.Transparent;

            lblActiveStud.ForeColor = SystemColors.ControlText;
            lblActiveCount.ForeColor = SystemColors.ControlText;

            lblActiveStud.Invalidate();   // Force repaint
            lblActiveCount.Invalidate();
        }

        private void panelInactive_Paint(object sender, PaintEventArgs e)
        {
            Color borderColor = Color.FromArgb(21, 112, 239);
            int borderWidth = 1;

            using (Pen borderPen = new Pen(borderColor, borderWidth))
            {
                // Draw rectangle border 1px inside to avoid clipping
                Rectangle rect = new Rectangle(0, 0, panelActive.Width - 1, panelActive.Height - 1);
                e.Graphics.DrawRectangle(borderPen, rect);
            }
        }

        private void panelInactive_MouseEnter(object sender, EventArgs e)
        {
            Color highlight = Color.FromArgb(21, 112, 239);
            panelInactive.BackColor = highlight;

            lblInactiveStud.ForeColor = Color.White;
            lblInactiveCount.ForeColor = Color.White;

            lblInactiveStud.Invalidate();   // Force repaint
            lblInactiveCount.Invalidate();
        }

        private void panelInactive_MouseLeave(object sender, EventArgs e)
        {
            panelInactive.BackColor = Color.Transparent;

            lblInactiveStud.ForeColor = SystemColors.ControlText;
            lblInactiveCount.ForeColor = SystemColors.ControlText;

            lblInactiveStud.Invalidate();   // Force repaint
            lblInactiveCount.Invalidate();
        }


        private void btnActive_MouseEnter(object sender, EventArgs e)
        {
            btnActive.BackColor = Color.FromArgb(21, 112, 239);
            btnActive.FlatAppearance.BorderColor = Color.FromArgb(21, 112, 239);
        }

        private void btnActive_MouseLeave(object sender, EventArgs e)
        {
            btnActive.BackColor = Color.Transparent;
            btnActive.FlatAppearance.BorderColor = Color.White;
        }

        private void btnInactive_MouseEnter(object sender, EventArgs e)
        {
            btnInactive.BackColor = Color.FromArgb(21, 112, 239);
            btnInactive.FlatAppearance.BorderColor = Color.FromArgb(21, 112, 239);
        }

        private void btnInactive_MouseLeave(object sender, EventArgs e)
        {
            btnInactive.BackColor = Color.Transparent;
            btnInactive.FlatAppearance.BorderColor = Color.White;
        }

        private void btnLogs_MouseEnter(object sender, EventArgs e)
        {
            btnLogs.BackColor = Color.FromArgb(21, 112, 239);
            btnLogs.FlatAppearance.BorderColor = Color.FromArgb(21, 112, 239);
        }

        private void btnLogs_MouseLeave(object sender, EventArgs e)
        {
            btnLogs.BackColor = Color.Transparent;
            btnLogs.FlatAppearance.BorderColor = Color.White;
        }

        private void btnLogout_MouseEnter(object sender, EventArgs e)
        {
            btnLogout.BackColor = Color.White;
            btnLogout.FlatAppearance.BorderColor = Color.White;
            btnLogout.ForeColor = Color.Red;
        }

        private void btnLogout_MouseLeave(object sender, EventArgs e)
        {
            btnLogout.BackColor = Color.Transparent;
            btnLogout.FlatAppearance.BorderColor = Color.White;
            btnLogout.ForeColor = Color.White;
        }

        private void btnAdd_MouseEnter(object sender, EventArgs e)
        {
            btnAdd.BackColor = Color.FromArgb(21, 112, 239);
            btnAdd.FlatAppearance.BorderColor = Color.FromArgb(21, 112, 239);
        }

        private void btnAdd_MouseLeave(object sender, EventArgs e)
        {
            btnAdd.BackColor = Color.FromArgb(26, 26, 26);
            btnAdd.FlatAppearance.BorderColor = Color.FromArgb(26, 26, 26);
        }

        private void btnViewAll_MouseEnter(object sender, EventArgs e)
        {
            btnViewAll.BackColor = Color.FromArgb(21, 112, 239);
            btnViewAll.FlatAppearance.BorderColor = Color.FromArgb(21, 112, 239);
        }

        private void btnViewAll_MouseLeave(object sender, EventArgs e)
        {
            btnViewAll.BackColor = Color.FromArgb(26, 26, 26);
            btnViewAll.FlatAppearance.BorderColor = Color.FromArgb(26, 26, 26);
        }


        public void showLogs(DataGridView d)
        {
            string fileLocation = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;

            string folder = "assets";
            string file = "Book1.xlsx";
            string path = Path.Combine(fileLocation, folder, file);

            Workbook book = new Workbook();

            book.LoadFromFile(path);

            Worksheet sh = book.Worksheets[0];

            DataTable dt = sh.ExportDataTable();

            d.DataSource = dt;
        }

        public int showCount(int c, string val)
        {
            string fileLocation = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;

            string folder = "assets";
            string file = "Book1.xlsx";
            string path = Path.Combine(fileLocation, folder, file);

            Workbook book = new Workbook();

            book.LoadFromFile(path);

            Worksheet sh = book.Worksheets[0];

            int r = sh.Rows.Length;

            int ctr = 0;

            for (int i = 2; i <= r; i++)
            {
                string cellValue = sh.Range[i, c].Value;

                if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains(val))
                {
                    ctr++;
                }
            }
            return ctr;
        }

        private void Dashboard_Load(object sender, EventArgs e)
        {
            lblName.Text = "Hello " + _userName + "!";
            lblDate.Text = "Date: " + DateTime.Now.ToString("MM/dd/yyyy");

            string fileLocation = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            string folder = "assets";
            string file = "Book1.xlsx";
            string path = Path.Combine(fileLocation, folder, file);

            Workbook book = new Workbook();
            book.LoadFromFile(path);
            Worksheet sheet = book.Worksheets[0];

            int row = sheet.Rows.Length;

            for (int i = 1; i <= row; i++)
            {
                string sheetUsername = sheet.Range[i, 11].Value;
                if (sheetUsername == _userName)
                {
                    string profilePicPath = sheet.Range[i, 14].Value;
                    if (!string.IsNullOrEmpty(profilePicPath) && File.Exists(profilePicPath))
                    {
                        pbProfile.SizeMode = PictureBoxSizeMode.StretchImage;
                        pbProfile.Image = Image.FromFile(profilePicPath);
                    }
                    break;
                }
            }
        }

        private void btnActive_Click_1(object sender, EventArgs e)
        {
            string fileLocation = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;

            string folder = "assets";
            string file = "Book1.xlsx";
            string path = Path.Combine(fileLocation, folder, file);

            Workbook book = new Workbook();

            book.LoadFromFile(path);

            Worksheet sh = book.Worksheets[0];

            DataTable dt = sh.ExportDataTable();
            DataTable filtered = dt.Clone();

            foreach (DataRow row in dt.Rows)
            {
                if (row[12].ToString() == "Active")
                {
                    filtered.ImportRow(row);
                }
            }

            f2 = new Form2(_userName, _form1, false, "Active", true);
            f2.dataGridView1.DataSource = filtered;
            f2.Show();
        }

        private void btnInactive_Click_1(object sender, EventArgs e)
        {
            string fileLocation = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;

            string folder = "assets";
            string file = "Book1.xlsx";
            string path = Path.Combine(fileLocation, folder, file);

            Workbook book = new Workbook();

            book.LoadFromFile(path);

            Worksheet sh = book.Worksheets[0];

            DataTable dt = sh.ExportDataTable();
            DataTable filtered = dt.Clone();

            foreach (DataRow row in dt.Rows)
            {
                if (row[12].ToString() == "Inactive")
                {
                    filtered.ImportRow(row);
                }
            }

            f2 = new Form2(_userName, _form1, true, "Inactive", false);
            f2.dataGridView1.DataSource = filtered;
            f2.Show();
        }

        private void btnLogs_Click_1(object sender, EventArgs e)
        {
            Logs_Page logsPage = new Logs_Page(_userName, _form1);

            logsPage.Show();
        }

        private void btnLogout_Click_1(object sender, EventArgs e)
        {
            Login lg = new Login();

            DialogResult confirmation = MessageBox.Show("Are you sure you want to log out?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (confirmation == DialogResult.Yes)
            {
                lg.Show();

                this.Hide();

                logs.insertLogs(_userName, "Successfully Logged out!");
            }
        }

        private void btnAdd_Click_1(object sender, EventArgs e)
        {
            DialogResult add_confirmation = MessageBox.Show("Are you sure you want to add new record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (add_confirmation == DialogResult.Yes)
            {
                _form1.Show();
            }
        }

        private void btnViewAll_Click_1(object sender, EventArgs e)
        {
            f2 = new Form2(_userName, _form1, false, "All", true); // Pass `true` to hide Reactivate button
            f2.Show();
        }
    }
}
