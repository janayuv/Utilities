using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities
{
    public partial class SequenceForm : Form
    {
        public enum FillOrder
        {
            Vertical,
            Horizontal
        }

        public int StartNumber { get; private set; }
        public int Increment { get; private set; }
        public int NumberOfDigits { get; private set; }
        public string Prefix { get; private set; }
        public string Suffix { get; private set; }
        public FillOrder SequenceFillOrder { get; private set; }

        public SequenceForm()
        {
            InitializeComponent();

            // Wire up the event handlers
            btnGeneratePreview.Click += btnGeneratePreview_Click;
            btnOK.Click += btnOK_Click;
            btnCancel.Click += btnCancel_Click;
        }
        private void btnGeneratePreview_Click(object sender, EventArgs e)
        {
            try
            {
                // Read user inputs
                StartNumber = int.Parse(txtStartNumber.Text);
                Increment = int.Parse(txtIncrement.Text);
                NumberOfDigits = int.Parse(txtDigits.Text);
                Prefix = txtPrefix.Text;
                Suffix = txtSuffix.Text;

                // Determine the fill order based on comboBox selection
                if (comboBoxFillOrder.SelectedItem.ToString() == "Fill Vertically")
                {
                    SequenceFillOrder = FillOrder.Vertical;
                }
                else
                {
                    SequenceFillOrder = FillOrder.Horizontal;
                }

                // Get the number of cells selected by the user
                int numberOfEntries = 10; // Default number of entries in preview
                if (SequenceFillOrder == FillOrder.Vertical)
                {
                    numberOfEntries = 10; // Preview vertically by default (you can change this logic)
                }
                else if (SequenceFillOrder == FillOrder.Horizontal)
                {
                    numberOfEntries = 10; // Preview horizontally by default (you can change this logic)
                }

                // Generate a sequence preview
                listBoxPreview.Items.Clear(); // Clear previous preview
                for (int i = 0; i < numberOfEntries; i++) // Preview based on selected range
                {
                    int currentNumber = StartNumber + (i * Increment);

                    // Apply padding logic: if NumberOfDigits is 0, do not pad
                    string formattedNumber = Prefix + currentNumber.ToString().PadLeft(NumberOfDigits > 0 ? NumberOfDigits : currentNumber.ToString().Length, '0') + Suffix;

                    listBoxPreview.Items.Add(formattedNumber);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Handle OK button click to save inputs and close the form
        private void btnOK_Click(object sender, EventArgs e)
        {
            try
            {
                // Get the Excel application and active worksheet
                Excel.Application excelApp = Globals.ThisAddIn.Application;
                Excel.Worksheet activeSheet = excelApp.ActiveSheet as Excel.Worksheet;

                if (activeSheet == null)
                {
                    MessageBox.Show("No active sheet found.");
                    return;
                }

                // Get the selected range of cells in the active sheet
                Excel.Range selectedRange = excelApp.Selection as Excel.Range;

                if (selectedRange == null || selectedRange.Cells.Count == 0)
                {
                    MessageBox.Show("Please select some cells to insert the sequence.");
                    return;
                }

                // Get the user inputs
                int currentNumber = StartNumber;
                int increment = Increment;

                // Get the number of rows and columns from the selected range
                int numRows = selectedRange.Rows.Count;
                int numCols = selectedRange.Columns.Count;

                // Check if we need to fill vertically or horizontally
                for (int i = 0; i < numRows; i++)  // Loop through rows
                {
                    for (int j = 0; j < numCols; j++)  // Loop through columns
                    {
                        // Apply padding logic if necessary
                        string formattedNumber = Prefix + currentNumber.ToString().PadLeft(NumberOfDigits > 0 ? NumberOfDigits : currentNumber.ToString().Length, '0') + Suffix;

                        // If the user selected cells horizontally
                        if (SequenceFillOrder == FillOrder.Horizontal)
                        {
                            // Fill cells horizontally: Move across columns in the same row
                            selectedRange.Cells[i + 1, j + 1].Value = formattedNumber;
                        }
                        // If the user selected cells vertically
                        else if (SequenceFillOrder == FillOrder.Vertical)
                        {
                            // Fill cells vertically: Move down rows in the same column
                            selectedRange.Cells[j + 1, i + 1].Value = formattedNumber;
                        }

                        // Increment the number for the next cell
                        currentNumber += increment;
                    }
                }

                // Show a confirmation message
                MessageBox.Show("Sequence inserted!");

                // Close the form
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error inserting sequence: " + ex.Message);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            // Handle Cancel button click
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void SequenceForm_Load(object sender, EventArgs e)
        {
            // Set default fill order to "Fill Vertically"
            if (comboBoxFillOrder.Items.Count == 0) // Ensure items are not already added
            {
                comboBoxFillOrder.Items.Add("Fill Vertically");
                comboBoxFillOrder.Items.Add("Fill Horizontally");
            }
            comboBoxFillOrder.SelectedItem = "Fill Horizontally"; // Default value

            // Set default number of digits to 0
            txtDigits.Text = "0"; // Default value is 0
        }
    }
}
