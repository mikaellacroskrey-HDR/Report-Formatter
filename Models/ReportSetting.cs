using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.Windows;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic.FileIO;

namespace ReportHelper.Models
{
    public class ReportSetting
    {
        public string FileName = string.Empty;
        public string FilePath = string.Empty;
        public string FolderPath = string.Empty;
        public string SavedFilePath = string.Empty;
        public bool ApplyNameFormatting = false;
        public bool ForceIntervals = false;
        public string StartTime = string.Empty;
        public string EndTime = string.Empty;
        public string EquipmentName = string.Empty;
        public string Date = string.Empty;
        public string DateAltFormat = string.Empty;

        public ReportSetting()
        {
        }

        /// <summary>
        /// Read the file and return a list of all of the Timestamps.
        /// </summary>
        /// <returns></returns>
        public List<string> GetTimeStamps()
        {
            List<string> timeStamps = new List<string>();

            using (TextFieldParser parser = new TextFieldParser(FilePath))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();
                    if (fields.FirstOrDefault() != null && fields.FirstOrDefault().ToUpper() != "TIMESTAMP")
                    {
                        timeStamps.Add(fields.First());
                    }
                }
            }
            return timeStamps;
        }

        /// <summary>
        /// Generate an .xlsx file of the transposed data within the specified timespan.
        /// </summary>
        public bool GenerateReport()
        {
            bool containsStartTime = false;
            bool containsEndTime = false;

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return false;
            }
                        
            object misValue = System.Reflection.Missing.Value;
            Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
            Worksheet xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            bool inInterval = false;
            int colIdx = 1;
            int rowIdx = 3;

            try
            {
                // Initial report formatting
                //xlWorkSheet.Name = EquipmentName; // Commenting out because it is not needed and was causing code to fail in some cases.
                xlWorkSheet.Cells[1, 1] = Date;
                xlWorkSheet.Cells[2, 1] = EquipmentName;

                Range rg = (Range)xlWorkSheet.Cells[3,1];
                rg.EntireRow.NumberFormat = "hh:mm:ss AM/PM";


                // Read the file and transpose data
                using (TextFieldParser parser = new TextFieldParser(FilePath))
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");

                    DateTime lastDateTime = DateTime.MinValue;

                    while (!parser.EndOfData)
                    {
                        string[] fields = parser.ReadFields();

                        // Read line if the first field of the line is a header, or after the start timestamp
                        if (inInterval || fields.FirstOrDefault() != null && fields.FirstOrDefault().ToUpper() == StartTime.ToUpper() || fields.FirstOrDefault().ToUpper() == "TIMESTAMP")
                        {
                            if (fields.FirstOrDefault().ToUpper() != "TIMESTAMP")
                            {
                                inInterval = true;
                                containsStartTime = true;
                            }

                            foreach (string field in fields)
                            {
                                // Insert timestamps to force intervals from 15 min to 5 min if applicable
                                if (ForceIntervals && rowIdx == 3 && colIdx >= 2)
                                {
                                    if (colIdx == 2)
                                    {
                                        lastDateTime = DateTime.Parse(field);
                                    }
                                    else if (DateTime.Parse(field).Minute - lastDateTime.Minute == 15 || (DateTime.Parse(field).Hour - lastDateTime.Hour == 1 && DateTime.Parse(field).Minute - lastDateTime.Minute == -45))
                                    {
                                        xlWorkSheet.Cells[rowIdx, colIdx] = lastDateTime.AddMinutes(5); ;
                                        xlWorkSheet.Cells[rowIdx, colIdx + 1] = lastDateTime.AddMinutes(10);
                                        xlWorkSheet.Cells[rowIdx, colIdx + 2] = DateTime.Parse(field);
                                        lastDateTime = DateTime.Parse(field);
                                        colIdx += 2;
                                        rowIdx++;
                                        continue;
                                    }
                                }

                                xlWorkSheet.Cells[rowIdx, colIdx] = field;
                                rowIdx++;
                            }
                            colIdx++;
                            rowIdx = 3;
                        }

                        // Stop parsing the file after we get to the end of timespan.
                        if (fields.FirstOrDefault() != null && fields.FirstOrDefault().ToUpper() == EndTime.ToUpper())
                        {
                            containsEndTime = true;
                            break;
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("An error occured generating the .xlsx file. \n " + ex.Message);

                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                // Free objects
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
                return false;
            }

            xlWorkSheet.Columns.AutoFit();

            try
            {
                //Save and close the Excel file
                xlWorkBook.SaveAs(SavedFilePath, XlFileFormat.xlOpenXMLWorkbook, misValue,
                misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            // Free objects
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            return (containsStartTime && containsEndTime);
        }
    }
}
