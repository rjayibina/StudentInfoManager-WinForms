using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.AutoFilter;
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

namespace EVEDRI_Lab_Act2
{
    public partial class Form2 : Form
    {
        private string _userName;

        private Form1 _form1;

        private string _filterType;

        Logs logs = new Logs();

        public Form2(string userName, Form1 form1, bool hideDeleteButtons = false, string filterType = "All", bool hideReactivateButton = false)
        {
            InitializeComponent();

            LoadExcelFile();

            _userName = userName;
            _filterType = filterType;
            _form1 = form1;

            // Set label title based on the filter type
            lblTitle.Text = _filterType + " Records";

            // Center it horizontally around X = 504
            Size textSize = TextRenderer.MeasureText(lblTitle.Text, lblTitle.Font);
            int centeredX = 504 - (textSize.Width / 2);
            lblTitle.Location = new Point(centeredX, 111);

            if (hideDeleteButtons)
            {
                btnDelete.Visible = false;
                btnDeleteAll.Visible = false;
            }

            if (hideReactivateButton)
            {
                btnSetActive.Visible = false;
            }

            btnClose.MouseEnter += btnClose_MouseEnter;
            btnClose.MouseLeave += btnClose_MouseLeave;
            btnSearch.MouseEnter += btnSearch_MouseEnter;
            btnSearch.MouseLeave += btnSearch_MouseLeave;
            btnSetActive.MouseEnter += btnSetActive_MouseEnter;
            btnSetActive.MouseLeave += btnSetActive_MouseLeave;
            btnDelete.MouseEnter += btnDelete_MouseEnter;
            btnDelete.MouseLeave += btnDelete_MouseLeave;            
            btnDeleteAll.MouseEnter += btnDeleteAll_MouseEnter;
            btnDeleteAll.MouseLeave += btnDeleteAll_MouseLeave;
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

        private void btnSetActive_MouseEnter(object sender, EventArgs e)
        {
            btnSetActive.BackColor = Color.FromArgb(21, 112, 239);
            btnSetActive.FlatAppearance.BorderColor = Color.FromArgb(21, 112, 239);
        }

        private void btnSetActive_MouseLeave(object sender, EventArgs e)
        {
            btnSetActive.BackColor = Color.FromArgb(26, 26, 26);
            btnSetActive.FlatAppearance.BorderColor = Color.FromArgb(26, 26, 26);
        }

        private void btnDelete_MouseEnter(object sender, EventArgs e)
        {
            btnDelete.BackColor = Color.FromArgb(21, 112, 239);
            btnDelete.FlatAppearance.BorderColor = Color.FromArgb(21, 112, 239);
        }

        private void btnDelete_MouseLeave(object sender, EventArgs e)
        {
            btnDelete.BackColor = Color.FromArgb(26, 26, 26);
            btnDelete.FlatAppearance.BorderColor = Color.FromArgb(26, 26, 26);
        }

        private void btnDeleteAll_MouseEnter(object sender, EventArgs e)
        {
            btnDeleteAll.BackColor = Color.FromArgb(21, 112, 239);
            btnDeleteAll.FlatAppearance.BorderColor = Color.FromArgb(21, 112, 239);
        }

        private void btnDeleteAll_MouseLeave(object sender, EventArgs e)
        {
            btnDeleteAll.BackColor = Color.FromArgb(26, 26, 26);
            btnDeleteAll.FlatAppearance.BorderColor = Color.FromArgb(26, 26, 26);
        }


        public void LoadExcelFile()
        {
            string fileLocation = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;

            string folder = "assets";
            string file = "Book1.xlsx";
            string path = Path.Combine(fileLocation, folder, file);

            Workbook book = new Workbook();

            book.LoadFromFile(path);

            Worksheet sheet = book.Worksheets[0];

            DataTable dt = sheet.ExportDataTable();

            dataGridView1.DataSource = dt;
        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (_form1 == null)
                {
                    MessageBox.Show("Form1 reference is missing.");
                    return;
                }

                int r = dataGridView1.CurrentCell?.RowIndex ?? -1;

                if (r < 0 || r >= dataGridView1.Rows.Count || dataGridView1.Rows[r].IsNewRow)
                {
                    MessageBox.Show("Invalid or empty row selected.");
                    return;
                }

                var cells = dataGridView1.Rows[r].Cells;
                if (cells.Count < 14)
                {
                    MessageBox.Show("Row data is incomplete or corrupted.");
                    return;
                }

                _form1.lblID.Text = r.ToString();
                _form1.txtName.Text = cells[0]?.Value?.ToString() ?? "";

                string gender = cells[1]?.Value?.ToString() ?? "";
                _form1.rdoMale.Checked = gender == "Male";
                _form1.rdoFemale.Checked = gender == "Female";

                string hobbies = cells[2]?.Value?.ToString() ?? "";
                _form1.chkBasketball.Checked = hobbies.Contains("Basketball");
                _form1.chkVolleyball.Checked = hobbies.Contains("Volleyball");
                _form1.chkSoccer.Checked = hobbies.Contains("Soccer");

                _form1.txtAddress.Text = cells[3]?.Value?.ToString() ?? "";
                _form1.txtEmail.Text = cells[4]?.Value?.ToString() ?? "";

                string birthday = cells[5]?.Value?.ToString();
                _form1.dtpBirthday.Value = DateTime.TryParse(birthday, out DateTime dob) ? dob : DateTime.Now;

                _form1.txtAge.Text = cells[6]?.Value?.ToString() ?? "";
                _form1.cmbFavColor.SelectedItem = cells[7]?.Value?.ToString() ?? "";
                _form1.txtSaying.Text = cells[8]?.Value?.ToString() ?? "";
                _form1.cmbProgram.SelectedItem = cells[9]?.Value?.ToString() ?? "";
                _form1.txtUsername.Text = cells[10]?.Value?.ToString() ?? "";
                _form1.txtPassword.Text = cells[11]?.Value?.ToString() ?? "";

                string picPath = cells[13]?.Value?.ToString() ?? "";
                _form1.txtProfilePic.Text = System.IO.File.Exists(picPath) ? picPath : "";

                _form1.btnAdd.Visible = false;
                _form1.btnUpdate.Visible = true;

                _form1.lblTitle.Text = "Updating a Record";

                _form1.Show();
                this.Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while selecting a record:\n" + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
                bool found = false;

                // Perform search
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[0].Value.ToString().Equals(txtSearchBox.Text, StringComparison.OrdinalIgnoreCase))
                    {
                        row.Selected = true;
                        found = true;
                        break;
                    }
                }

                // Determine the correct context for the search log
                string searchContext = string.IsNullOrEmpty(_filterType) ? "All Users page" : _filterType + " Users page";

                // Log the search
                Logs logs = new Logs();
                logs.insertLogs(_userName, "Searched for " + txtSearchBox.Text + " in " + searchContext + ".");

                if (!found)
                {
                    MessageBox.Show("Record not found.");
                }
            }

            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }

        private void btnSetActive_Click_1(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select a row to activate.");
                return;
            }

            DialogResult confirmation = MessageBox.Show("Are you sure you want to set this record as Active?",
                                                        "Confirmation",
                                                        MessageBoxButtons.YesNo,
                                                        MessageBoxIcon.Question);

            if (confirmation == DialogResult.Yes)
            {
                DataGridViewRow selectedRow = dataGridView1.SelectedRows[0];
                string selectedName = selectedRow.Cells[0].Value.ToString(); // Assuming name is in column 0

                string fileLocation = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
                string folder = "assets";
                string file = "Book1.xlsx";
                string path = Path.Combine(fileLocation, folder, file);

                Workbook book = new Workbook();
                book.LoadFromFile(path);
                Worksheet sheet = book.Worksheets[0];

                int rowCount = sheet.Rows.Length;

                for (int i = 2; i <= rowCount; i++) // start from row 2 to skip header
                {
                    string nameInSheet = sheet.Range[i, 1].Value; // column 1 = Excel column A

                    if (nameInSheet == selectedName)
                    {
                        sheet.Range[i, 13].Value = "Active"; // column 13 = Status
                        break;
                    }
                }

                book.SaveToFile(path, ExcelVersion.Version2016);

                MessageBox.Show("Record has been set to Active!");

                logs.insertLogs(_userName, "Reactivated a user: " + selectedName);

                // Filter and reload only Inactive records in the DataGridView
                DataTable dt = sheet.ExportDataTable();
                DataTable filtered = dt.Clone();

                foreach (DataRow row in dt.Rows)
                {
                    if (row[12].ToString() == "Inactive") // column 13 (index 12)
                    {
                        filtered.ImportRow(row);
                    }
                }

                dataGridView1.DataSource = filtered;
            }
        }

        private void btnDelete_Click_1(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select a row to delete.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult confirmation = MessageBox.Show("Are you sure you want to delete the selected user(s)?", "Confirmation", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

            if (confirmation != DialogResult.OK) return;

            string fileLocation = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            string folder = "assets";
            string file = "Book1.xlsx";
            string path = Path.Combine(fileLocation, folder, file);

            Workbook book = new Workbook();
            book.LoadFromFile(path);
            Worksheet sheet = book.Worksheets[0];

            // Collect names of users marked as inactive for logging
            List<string> deletedUsers = new List<string>();

            foreach (DataGridViewRow dgvRow in dataGridView1.SelectedRows)
            {
                string nameToDelete = dgvRow.Cells[0].Value.ToString();

                for (int i = 2; i <= sheet.LastRow; i++) // assuming data starts from row 2
                {
                    string nameInSheet = sheet.Range[i, 1].Value;

                    if (nameInSheet == nameToDelete)
                    {
                        sheet.Range[i, 13].Value = "Inactive"; // Mark as inactive
                        deletedUsers.Add(nameToDelete);
                        break;
                    }
                }

                dataGridView1.Rows.Remove(dgvRow);
            }

            // Save changes to the Excel file
            book.SaveToFile(path, ExcelVersion.Version2016);

            // Log all deleted users after processing
            foreach (var name in deletedUsers)
            {
                logs.insertLogs(_userName, $"Marked '{name}' as Inactive.");
            }

            MessageBox.Show("Successfully marked as Inactive and removed from the view.");
        }

        private void btnDeleteAll_Click_1(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                // Confirm deletion
                var confirmResult = MessageBox.Show("Are you sure you want to delete all rows?",
                                                     "Confirm Deletion",
                                                     MessageBoxButtons.YesNo);
                if (confirmResult == DialogResult.Yes)
                {
                    // Clear all rows in the DataGridView
                    dataGridView1.Rows.Clear();
                }
            }
        }
    }
}
