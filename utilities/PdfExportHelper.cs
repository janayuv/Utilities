using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities
{
    public static class PdfExportHelper
    {
        public static void ExportRangeToPdf(Excel.Range selectedRange)
        {
            try
            {
                // Ensure the range is valid
                if (selectedRange == null)
                {
                    throw new ArgumentNullException(nameof(selectedRange), "The selected range is null.");
                }

                Excel.Application app = selectedRange.Application;
                Excel.Worksheet activeSheet = selectedRange.Worksheet;

                // Step 1: Set the print area to the selected range
                activeSheet.PageSetup.PrintArea = selectedRange.Address;

                // Step 2: Set print scaling to fit the range on one page
                activeSheet.PageSetup.Zoom = false; // Disable Zoom
                activeSheet.PageSetup.FitToPagesWide = 1; // Fit to one page wide
                activeSheet.PageSetup.FitToPagesTall = 1; // Fit to one page tall

                // Step 3: Show Print Preview
                DialogResult previewResult = MessageBox.Show(
                    "Would you like to preview the selected range before exporting?",
                    "Print Preview",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question
                );

                if (previewResult == DialogResult.Cancel)
                {
                    return; // Exit if the user cancels
                }

                if (previewResult == DialogResult.Yes)
                {
                    // Show print preview with the selected range as the print area
                    selectedRange.Application.Dialogs[Excel.XlBuiltInDialog.xlDialogPrintPreview].Show();
                }

                // Step 4: Select Save Location
                using (SaveFileDialog saveDialog = new SaveFileDialog())
                {
                    saveDialog.Filter = "PDF Files|*.pdf";
                    saveDialog.Title = "Save PDF File";
                    saveDialog.FileName = $"{activeSheet.Name}_Selection.pdf";

                    if (saveDialog.ShowDialog() != DialogResult.OK)
                    {
                        return; // Exit if the user cancels the save dialog
                    }

                    string savePath = saveDialog.FileName;

                    // Step 5: Export to PDF
                    selectedRange.ExportAsFixedFormat(
                        Excel.XlFixedFormatType.xlTypePDF,
                        savePath,
                        Excel.XlFixedFormatQuality.xlQualityStandard,
                        IncludeDocProperties: true,
                        IgnorePrintAreas: false,
                        From: Type.Missing,
                        To: Type.Missing,
                        OpenAfterPublish: true
                    );

                    MessageBox.Show($"PDF successfully exported to:\n{savePath}", "Export Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while exporting to PDF:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
