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

namespace EVEDRI_Lab_Act2
{
    public partial class Logs_Page: Form
    {
        Logs logs = new Logs();

        private string _userName;

        public Logs_Page(string userName, Form1 form1, bool hideDeleteButtons = false)
        {
            InitializeComponent();

            _userName = userName;

            LoadExcelFile();

            btnSearch.MouseEnter += btnSearch_MouseEnter;
            btnSearch.MouseLeave += btnSearch_MouseLeave;
            btnClose.MouseEnter += btnClose_MouseEnter;
            btnClose.MouseLeave += btnClose_MouseLeave;
        }

        private void btnClose_MouseEnter(object sender, EventArgs e)
        {
            btnClose.BackColor = Color.FromArgb(21, 112, 239);
            btnClose.FlatAppearance.BorderColor = Color.FromArgb(21, 112, 239);
        }

        private void btnClose_MouseLeave(object sender, EventArgs e)
        {
            btnClose.BackColor = Color.FromArgb(26, 26, 26);
            btnClose.FlatAppearance.BorderColor = Color.FromArgb(26, 26, 26);
        }

        private void btnSearch_MouseEnter(object sender, EventArgs e)
        {
            btnSearch.BackColor = Color.FromArgb(21, 112, 239);
            btnSearch.FlatAppearance.BorderColor = Color.FromArgb(21, 112, 239);
        }

        private void btnSearch_MouseLeave(object sender, EventArgs e)
        {
            btnSearch.BackColor = Color.FromArgb(26, 26, 26);
            btnSearch.FlatAppearance.BorderColor = Color.FromArgb(26, 26, 26);
        }

        public void LoadExcelFile()
        {
            string fileLocation = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;

            string folder = "assets";
            string file = "Book1.xlsx";
            string path = Path.Combine(fileLocation, folder, file);

            Workbook book = new Workbook();

            book.LoadFromFile(path);

            Worksheet sheet = book.Worksheets[1];

            DataTable dt = sheet.ExportDataTable();

            dataGridView1.DataSource = dt;
        }

        private void btnClose_Click_1(object sender, EventArgs e)
        {
            DialogResult close_confimation = MessageBox.Show("Are you sure you want to go back to the Dashboard?", "Confirmation", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

            // Check the user's response
            if (close_confimation == DialogResult.OK)
            {
                //Dashboard ds = new Dashboard(_userName, _form1);

                //ds.Show();

                this.Close();
            }
        }

        private void btnSearch_Click_1(object sender, EventArgs e)
        {
            dataGridView1.ClearSelection();

            try
            {

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[0].Value.ToString().Equals(txtSearchBox.Text))
                    {
                        row.Selected = true;

                        //break;
                    }
                }

            }
            catch (Exception err)
            {

                MessageBox.Show(err.Message);
            }

            logs.insertLogs(_userName, "Searched for " + txtSearchBox.Text + " in Logs page.");
        }
    }
}
