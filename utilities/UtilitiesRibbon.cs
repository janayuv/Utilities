using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Microsoft.Office.Core;
using System;

namespace utilities
{
    public partial class UtilitiesRibbon : RibbonBase
    {
        private void UtilitiesRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            // Optional: Any load logic for the ribbon
        }
        private void btnColorDuplicates_Click(object sender, RibbonControlEventArgs e)
        {
            ColorDuplicatesOnlyForDuplicates();
        }

        private void btnInsertSequence_Click(object sender, RibbonControlEventArgs e)
        {
            SequenceForm sequenceForm = new SequenceForm(); // Your form name
            sequenceForm.Show(); // This will display the form
        }


        private object GenerateUniqueColor(int index)
        {
            Random rnd = new Random(index);
            int red = rnd.Next(50, 255);   // Avoid very dark colors
            int green = rnd.Next(50, 255);
            int blue = rnd.Next(50, 255);

            return (red << 16) | (green << 8) | blue; // Combine RGB values
        }

        private void ColorDuplicatesOnlyForDuplicates()
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

            var colorDict = new Dictionary<string, object>();
            var countDict = new Dictionary<string, int>();

            foreach (Excel.Range cell in selectedRange)
            {
                if (cell.Value2 != null && !string.IsNullOrEmpty(cell.Value2.ToString()))
                {
                    string value = cell.Value2.ToString();
                    if (countDict.ContainsKey(value))
                        countDict[value]++;
                    else
                        countDict.Add(value, 1);
                }
            }

            int colorIndex = 0;
            foreach (Excel.Range cell in selectedRange)
            {
                if (cell.Value2 != null && !string.IsNullOrEmpty(cell.Value2.ToString()))
                {
                    string value = cell.Value2.ToString();
                    if (countDict[value] > 1)
                    {
                        if (!colorDict.ContainsKey(value))
                        {
                            if (colorIndex < colorList.Count)
                            {
                                colorDict[value] = colorList[colorIndex++];
                            }
                            else
                            {
                                colorDict[value] = GenerateUniqueColor(colorIndex++); // Generate dynamic color
                            }
                        }

                        cell.Interior.Color = colorDict[value];
                    }
                    else
                    {
                        cell.Interior.ColorIndex = Excel.Constants.xlNone;
                    }
                }
            }

            MessageBox.Show("Duplicates have been successfully colored.", "Operation Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

                
    }
}
