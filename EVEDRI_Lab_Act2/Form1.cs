using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Text.RegularExpressions;

namespace EVEDRI_Lab_Act2
{
    public partial class Form1 : Form
    {
        private Form2 form2;
        Logs logs = new Logs();

        public Form1()
        {
            InitializeComponent();

            btnBack.MouseEnter += btnBack_MouseEnter;
            btnBack.MouseLeave += btnBack_MouseLeave;
            btnAdd.MouseEnter += btnAdd_MouseEnter;
            btnAdd.MouseLeave += btnAdd_MouseLeave;
            btnUpdate.MouseEnter += btnUpdate_MouseEnter;
            btnUpdate.MouseLeave += btnUpdate_MouseLeave;
            btnBrowse.MouseEnter += btnBrowse_MouseEnter;
            btnBrowse.MouseLeave += btnBrowse_MouseLeave;
        }

        private void btnBack_MouseEnter(object sender, EventArgs e)
        {
            btnBack.BackColor = Color.FromArgb(21, 112, 239);
            btnBack.FlatAppearance.BorderColor = Color.FromArgb(21, 112, 239);
        }

        private void btnBack_MouseLeave(object sender, EventArgs e)
        {
            btnBack.BackColor = Color.FromArgb(26, 26, 26);
            btnBack.FlatAppearance.BorderColor = Color.FromArgb(26, 26, 26);
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

        private void btnUpdate_MouseEnter(object sender, EventArgs e)
        {
            btnUpdate.BackColor = Color.FromArgb(21, 112, 239);
            btnUpdate.FlatAppearance.BorderColor = Color.FromArgb(21, 112, 239);
        }

        private void btnUpdate_MouseLeave(object sender, EventArgs e)
        {
            btnUpdate.BackColor = Color.FromArgb(26, 26, 26);
            btnUpdate.FlatAppearance.BorderColor = Color.FromArgb(26, 26, 26);
        }

        private void btnBrowse_MouseEnter(object sender, EventArgs e)
        {
            btnBrowse.BackColor = Color.FromArgb(21, 112, 239);
            btnBrowse.FlatAppearance.BorderColor = Color.FromArgb(21, 112, 239);
        }

        private void btnBrowse_MouseLeave(object sender, EventArgs e)
        {
            btnBrowse.BackColor = Color.FromArgb(26, 26, 26);
            btnBrowse.FlatAppearance.BorderColor = Color.FromArgb(26, 26, 26);
        }

        private void btnDisplay_Click(object sender, EventArgs e)
        {
            if (form2 == null)
            {
                MessageBox.Show("No records to display. Please add a record first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            form2.Show();
        }

        private void dtpBirthday_ValueChanged(object sender, EventArgs e)
        {
            string[] date = dtpBirthday.Text.Split(',');

            int age = DateTime.Now.Year - Convert.ToInt32(date[2]);

            txtAge.Text = age.ToString();
        }

        public bool CheckEmpty()
        {
            StringBuilder error = new StringBuilder();

            if (string.IsNullOrWhiteSpace(txtName.Text)) error.AppendLine("Name is required.");
            if (!rdoMale.Checked && !rdoFemale.Checked) error.AppendLine("Gender is required.");
            if (!chkBasketball.Checked && !chkVolleyball.Checked && !chkSoccer.Checked) error.AppendLine("At least one hobby must be selected.");
            if (string.IsNullOrWhiteSpace(txtAddress.Text)) error.AppendLine("Address is required.");
            if (string.IsNullOrWhiteSpace(txtEmail.Text)) error.AppendLine("Email is required.");

            int birthYear = dtpBirthday.Value.Year;
            int currentYear = DateTime.Now.Year;
            int calculatedAge = currentYear - birthYear;

            if (birthYear >= currentYear)
                error.AppendLine("Birth year cannot be in the future.");
            else if (calculatedAge < 8)
                error.AppendLine("User must be at least 8 years old.");

            if (string.IsNullOrWhiteSpace(txtAge.Text)) error.AppendLine("Age is required.");
            if (cmbFavColor.SelectedIndex == -1) error.AppendLine("Favorite color must be selected.");
            if (string.IsNullOrWhiteSpace(txtSaying.Text)) error.AppendLine("Saying is required.");
            if (cmbProgram.SelectedIndex == -1) error.AppendLine("Program must be selected.");
            if (string.IsNullOrWhiteSpace(txtUsername.Text)) error.AppendLine("Username is required.");
            if (string.IsNullOrWhiteSpace(txtPassword.Text)) error.AppendLine("Password is required.");
            if (string.IsNullOrWhiteSpace(txtProfilePic.Text)) error.AppendLine("Profile picture is required.");


            if (error.Length > 0)
            {
                MessageBox.Show(error.ToString(), "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        // Method to check if username already exists in the spreadsheet
        private bool IsUsernameExists(string username)
        {
            string fileLocation = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            string folder = "assets";
            string file = "Book1.xlsx";
            string path = Path.Combine(fileLocation, folder, file);

            Workbook book = new Workbook(); 
            book.LoadFromFile(path);
            Worksheet sheet = book.Worksheets[0];

            int rows = sheet.Rows.Length + 1;

            foreach (var row in sheet.Rows)
            {
                string existingUsername = row.Columns[10].Value; // Column 11 is index 10
                if (existingUsername.Equals(username, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false; // Username does not exist
        }

        private void btnBack_Click_1(object sender, EventArgs e)
        {
            DialogResult confirmation = MessageBox.Show("Are you sure you want to go back to the Dashboard?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (confirmation == DialogResult.Yes)
            {
                this.Hide();
            }
        }

        private void btnBrowse_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();

            if (file.ShowDialog() == DialogResult.OK)
            {
                txtProfilePic.Text = file.FileName;
            }
        }

        private void btnAdd_Click_1(object sender, EventArgs e)
        {
            if (!CheckEmpty())
                return;


            if (IsUsernameExists(txtUsername.Text))
            {
                MessageBox.Show("Username already exists!", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!Regex.IsMatch(txtName.Text, @"^[a-zA-Z\s]+$"))
            {
                MessageBox.Show("Name must only contain letters.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!Regex.IsMatch(txtEmail.Text, @"^[^@\s]+@[^@\s]+\.[^@\s]+$"))
            {
                MessageBox.Show("Invalid email format.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string gender = "";

            if (rdoMale.Checked)
            {
                gender = rdoMale.Text;
            }

            if (rdoFemale.Checked)
            {
                gender = rdoFemale.Text;
            }

            string hobby = "";

            if (chkBasketball.Checked)
            {
                hobby += chkBasketball.Text + ", ";
            }

            if (chkVolleyball.Checked)
            {
                hobby += chkVolleyball.Text + ", ";
            }

            if (chkSoccer.Checked)
            {
                hobby += chkSoccer.Text + ", ";
            }

            string fileLocation = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;

            string folder = "assets";
            string file = "Book1.xlsx";
            string path = Path.Combine(fileLocation, folder, file);

            Workbook book = new Workbook();

            book.LoadFromFile(path);

            Worksheet sheet = book.Worksheets[0];

            int row = sheet.Rows.Length + 1;

            sheet.Range[row, 1].Value = txtName.Text;
            sheet.Range[row, 2].Value = gender;
            sheet.Range[row, 3].Value = hobby;
            sheet.Range[row, 4].Value = txtAddress.Text;
            sheet.Range[row, 5].Value = txtEmail.Text;
            sheet.Range[row, 6].Value = dtpBirthday.Text;
            sheet.Range[row, 7].Value = txtAge.Text;
            sheet.Range[row, 8].Value = cmbFavColor.Text;
            sheet.Range[row, 9].Value = txtSaying.Text;
            sheet.Range[row, 10].Value = cmbProgram.Text;
            sheet.Range[row, 11].Value = txtUsername.Text;
            sheet.Range[row, 12].Value = txtPassword.Text;
            sheet.Range[row, 13].Value = "Active";
            sheet.Range[row, 14].Value = txtProfilePic.Text;

            book.SaveToFile(path, ExcelVersion.Version2016);

            DataTable dt = sheet.ExportDataTable();

            form2 = new Form2(txtUsername.Text, this);

            form2.dataGridView1.DataSource = dt;

            MessageBox.Show("Successfully added!");

            logs.insertLogs(txtUsername.Text, "New record has added successfully!");

            //reset fields
            txtName.Clear();
            rdoMale.Checked = false;
            rdoFemale.Checked = false;
            chkBasketball.Checked = false;
            chkVolleyball.Checked = false;
            chkSoccer.Checked = false;
            txtAddress.Clear();
            txtEmail.Clear();
            txtAge.Clear();
            cmbFavColor.SelectedIndex = -1;
            txtSaying.Clear();
            cmbProgram.SelectedIndex = -1;
            txtUsername.Clear();
            txtPassword.Clear();
            txtProfilePic.Clear();

            //focus to name field
            txtName.Focus();
        }

        private void btnUpdate_Click_1(object sender, EventArgs e)
        {
            if (!CheckEmpty())
                return;

            if (!Regex.IsMatch(txtName.Text, @"^[a-zA-Z\s]+$"))
            {
                MessageBox.Show("Name must only contain letters.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!Regex.IsMatch(txtEmail.Text, @"^[^@\s]+@[^@\s]+\.[^@\s]+$"))
            {
                MessageBox.Show("Invalid email format.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {

                string name = txtName.Text;

                string gender = "";

                if (rdoMale.Checked)
                {
                    gender = rdoMale.Text;
                }

                if (rdoFemale.Checked)
                {
                    gender = rdoFemale.Text;
                }

                string hobby = "";

                if (chkBasketball.Checked)
                {
                    hobby += chkBasketball.Text + ", ";
                }

                if (chkVolleyball.Checked)
                {
                    hobby += chkVolleyball.Text + ", ";
                }

                if (chkSoccer.Checked)
                {
                    hobby += chkSoccer.Text + ", ";
                }

                string favColor = cmbFavColor.Text;

                string saying = txtSaying.Text;

                Workbook book = new Workbook();

                string fileLocation = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;

                string folder = "assets";
                string file = "Book1.xlsx";
                string path = Path.Combine(fileLocation, folder, file);

                book.LoadFromFile(path);

                Worksheet sheet = book.Worksheets[0];

                if (!int.TryParse(lblID.Text, out int rowIndex))
                {
                    MessageBox.Show("Invalid ID Format", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                int row = rowIndex + 2;

                sheet.Range[row, 1].Value = txtName.Text;
                sheet.Range[row, 2].Value = gender;
                sheet.Range[row, 3].Value = hobby;
                sheet.Range[row, 4].Value = txtAddress.Text;
                sheet.Range[row, 5].Value = txtEmail.Text;
                sheet.Range[row, 6].Value = dtpBirthday.Text;
                sheet.Range[row, 7].Value = txtAge.Text;
                sheet.Range[row, 8].Value = cmbFavColor.Text;
                sheet.Range[row, 9].Value = txtSaying.Text;
                sheet.Range[row, 10].Value = cmbProgram.Text;
                sheet.Range[row, 11].Value = txtUsername.Text;
                sheet.Range[row, 12].Value = txtPassword.Text;
                sheet.Range[row, 13].Value = "Active";
                sheet.Range[row, 14].Value = txtProfilePic.Text;

                book.SaveToFile(path, ExcelVersion.Version2016);

                DataTable dt = sheet.ExportDataTable();

                form2 = new Form2(txtUsername.Text, this);

                form2.dataGridView1.DataSource = dt;

                MessageBox.Show("Successfully updated!");

                logs.insertLogs(txtUsername.Text, "Record has been updated.");

            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }
    }
}
