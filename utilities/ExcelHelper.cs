using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities
{
    public static class ExcelHelper
    {
        public static void ColorDuplicatesOnlyForDuplicates()
        {
            Excel.Range selectedRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if (selectedRange == null)
            {
                MessageBox.Show("Please select a range of cells first.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var colorList = new List<object>
            {
                Excel.XlRgbColor.rgbSkyBlue,
                Excel.XlRgbColor.rgbLightGreen,
                Excel.XlRgbColor.rgbLightPink,
                Excel.XlRgbColor.rgbLightYellow,
                Excel.XlRgbColor.rgbLightCoral,
                Excel.XlRgbColor.rgbLavender,
                Excel.XlRgbColor.rgbPaleGreen,
                Excel.XlRgbColor.rgbWheat,
                Excel.XlRgbColor.rgbLightSalmon,
                Excel.XlRgbColor.rgbKhaki
            };

            // Step 1: Read all cell values into a dictionary
            object[,] values = selectedRange.Value2;
            var countDict = new Dictionary<string, int>();

            for (int row = 1; row <= values.GetLength(0); row++)
            {
                for (int col = 1; col <= values.GetLength(1); col++)
                {
                    var cellValue = values[row, col]?.ToString();
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        if (countDict.ContainsKey(cellValue))
                            countDict[cellValue]++;
                        else
                            countDict[cellValue] = 1;
                    }
                }
            }

            // Step 2: Assign colors to duplicates
            var colorDict = new Dictionary<string, object>();
            int colorIndex = 0;

            for (int row = 1; row <= values.GetLength(0); row++)
            {
                for (int col = 1; col <= values.GetLength(1); col++)
                {
                    var cellValue = values[row, col]?.ToString();
                    if (!string.IsNullOrEmpty(cellValue) && countDict[cellValue] > 1)
                    {
                        if (!colorDict.ContainsKey(cellValue))
                        {
                            if (colorIndex < colorList.Count)
                            {
                                colorDict[cellValue] = colorList[colorIndex++];
                            }
                            else
                            {
                                colorDict[cellValue] = GenerateUniqueColor(colorIndex++);
                            }
                        }

                        selectedRange.Cells[row, col].Interior.Color = colorDict[cellValue];
                    }
                    else
                    {
                        selectedRange.Cells[row, col].Interior.ColorIndex = Excel.Constants.xlNone;
                    }
                }
            }

            MessageBox.Show("Duplicates have been successfully colored.", "Operation Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static object GenerateUniqueColor(int index)
        {
            Random rnd = new Random(index);
            int red = rnd.Next(50, 255);   // Avoid very dark colors
            int green = rnd.Next(50, 255);
            int blue = rnd.Next(50, 255);

            return (red << 16) | (green << 8) | blue; // Combine RGB values
        }
    }
}
