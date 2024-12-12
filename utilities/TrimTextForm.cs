using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities
{
    public partial class TrimTextForm : Form
    {
        public TrimTextForm()
        {
            InitializeComponent();
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            var selectedRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if (selectedRange == null)
            {
                MessageBox.Show("Please select a valid range.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Handle the selected option
            if (optDeleteLeadingTrailingSpaces.Checked)
            {
                DeleteLeadingTrailingSpaces(selectedRange);
            }
            else if (optDeleteLeadingTrailingExcessiveSpaces.Checked)
            {
                DeleteLeadingTrailingExcessiveSpaces(selectedRange);
            }
            else if (optDeleteLeadingCharacters.Checked)
            {
                DeleteLeadingCharacters(selectedRange);
            }
            else if (optDeleteEndingCharacters.Checked)
            {
                DeleteEndingCharacters(selectedRange);
            }
            else
            {
                MessageBox.Show("Please select an option.", "No Option Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            MessageBox.Show("Operation completed successfully.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // Delete leading and trailing spaces
        private void DeleteLeadingTrailingSpaces(Excel.Range selectedRange)
        {
            foreach (Excel.Range cell in selectedRange.Cells)
            {
                if (cell.Value2 != null)
                {
                    if (cell.Value2 is string)
                    {
                        cell.Value2 = cell.Value2.ToString().Trim();
                    }
                }
            }
        }

        // Delete leading, trailing, and excessive spaces
        private void DeleteLeadingTrailingExcessiveSpaces(Excel.Range selectedRange)
        {
            foreach (Excel.Range cell in selectedRange.Cells)
            {
                if (cell.Value2 != null)
                {
                    if (cell.Value2 is string)
                    {
                        string cleanedText = System.Text.RegularExpressions.Regex.Replace(cell.Value2.ToString().Trim(), @"\s+", " ");
                        cell.Value2 = cleanedText;
                    }
                }
            }
        }

        // Delete a specified number of leading characters
        private void DeleteLeadingCharacters(Excel.Range selectedRange)
        {
            if (ValidateNumberInput(out int numberOfChars))
            {
                foreach (Excel.Range cell in selectedRange.Cells)
                {
                    if (cell.Value2 != null)
                    {
                        if (cell.Value2 is string)
                        {
                            string newValue = TruncateString(cell.Value2.ToString(), numberOfChars, isLeading: true);
                            cell.Value2 = newValue;
                        }
                        else if (IsConvertibleToString(cell.Value2))
                        {
                            string newValue = TruncateString(cell.Value2.ToString(), numberOfChars, isLeading: true);
                            cell.Value2 = newValue;
                        }
                    }
                }
            }
        }

        // Delete a specified number of ending characters
        private void DeleteEndingCharacters(Excel.Range selectedRange)
        {
            if (ValidateNumberInput(out int numberOfChars))
            {
                foreach (Excel.Range cell in selectedRange.Cells)
                {
                    if (cell.Value2 != null)
                    {
                        if (cell.Value2 is string)
                        {
                            string newValue = TruncateString(cell.Value2.ToString(), numberOfChars, isLeading: false);
                            cell.Value2 = newValue;
                        }
                        else if (IsConvertibleToString(cell.Value2))
                        {
                            string newValue = TruncateString(cell.Value2.ToString(), numberOfChars, isLeading: false);
                            cell.Value2 = newValue;
                        }
                    }
                }
            }
        }

        // Validate input from the TextBox
        private bool ValidateNumberInput(out int numberOfChars)
        {
            if (int.TryParse(txtNumberOfCharacters.Text, out numberOfChars) && numberOfChars > 0)
            {
                return true;
            }
            else
            {
                MessageBox.Show("Please enter a valid number of characters to remove.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                numberOfChars = 0;
                return false;
            }
        }

        // Truncate a string
        private string TruncateString(string input, int numberOfChars, bool isLeading)
        {
            if (isLeading)
            {
                return input.Length > numberOfChars ? input.Substring(numberOfChars) : string.Empty;
            }
            else
            {
                return input.Length > numberOfChars ? input.Substring(0, input.Length - numberOfChars) : string.Empty;
            }
        }

        // Check if an object is convertible to a string
        private bool IsConvertibleToString(object value)
        {
            return value is DateTime || value is double || value is int || value is decimal;
        }
    }
}
