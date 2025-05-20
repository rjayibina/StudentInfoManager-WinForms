using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace EVEDRI_Lab_Act2
{
    public partial class Login : Form
    {
        Logs logs = new Logs();

        public Login()
        {
            InitializeComponent();

            btnLogin.MouseEnter += btnLogin_MouseEnter;
            btnLogin.MouseLeave += btnLogin_MouseLeave;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            // Display a confirmation message box
            DialogResult result = MessageBox.Show("Are you sure you want to exit?", "Exit Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            // Check the user's response
            if (result == DialogResult.Yes)
            {
                // Close the form
                this.Close();
            }
        }

        private void btnLogin_Click_1(object sender, EventArgs e)
        {
            string username = txtUsername.Text.Trim();
            string password = textBox1.Text.Trim();

            // Validation: both fields empty
            if (string.IsNullOrEmpty(username) && string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Please enter both username and password.", "Login Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Validation: only username empty
            if (string.IsNullOrEmpty(username))
            {
                MessageBox.Show("Please enter your username.", "Login Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Validation: only password empty
            if (string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Please enter your password.", "Login Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string fileLocation = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            string folder = "assets";
            string file = "Book1.xlsx";
            string path = Path.Combine(fileLocation, folder, file);

            Workbook book = new Workbook();
            book.LoadFromFile(path);

            Worksheet sheet = book.Worksheets[0];
            int row = sheet.Rows.Length;

            bool usernameExists = false;
            bool isAuthenticated = false;
            string userName = "";
            string profilePicPath = "";
            Form1 f1 = new Form1();

            for (int i = 1; i <= row; i++)
            {
                string sheetUsername = sheet.Range[i, 11].Value;
                string sheetPassword = sheet.Range[i, 12].Value;
                string status = sheet.Range[i, 13].Value;

                if (sheetUsername == username)
                {
                    usernameExists = true;

                    if (sheetPassword == password)
                    {
                        if (status == "Inactive")
                        {
                            MessageBox.Show("Your account is not active.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            logs.insertLogs(username, "Attempted login but user is inactive.");
                            return;
                        }

                        userName = username;
                        profilePicPath = sheet.Range[i, 14].Value;

                        logs.insertLogs(userName, "Successfully Logged in!");
                        isAuthenticated = true;
                        break;
                    }
                }
            }

            if (!usernameExists)
            {
                MessageBox.Show("Username does not exist. Please check and try again.", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (!isAuthenticated)
            {
                MessageBox.Show("Incorrect password. Please try again.", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                Dashboard dashboard = new Dashboard(userName, f1);
                MessageBox.Show("Login Success", "Welcome", MessageBoxButtons.OK, MessageBoxIcon.Information);
                dashboard.Show();
                this.Hide();
            }
        }

        private void btnLogin_MouseEnter(object sender, EventArgs e)
        {
            btnLogin.BackColor = Color.FromArgb(21, 112, 239);
            btnLogin.FlatAppearance.BorderColor = Color.FromArgb(21, 112, 239);
        }

        private void btnLogin_MouseLeave(object sender, EventArgs e)
        {
            btnLogin.BackColor = Color.FromArgb(26, 26, 26);
            btnLogin.FlatAppearance.BorderColor = Color.FromArgb(26, 26, 26);
        }
    }
}
