using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using ReportHelper.Models;
using System.IO;
using System.Linq;
using System.Windows.Input;

namespace ReportHelper
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string filePath = string.Empty;
        public ReportSetting primaryReportSetting;
        private const string CSV = ".csv";
        private const string XLSX = ".xlsx";

        public MainWindow()
        {
            InitializeComponent();
            primaryReportSetting = new ReportSetting();
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";

            if (!(bool)ofd.ShowDialog())
            {
                return;
            }

            primaryReportSetting.FilePath = ofd.FileName;
            primaryReportSetting.FileName = ofd.SafeFileName;
            primaryReportSetting.FolderPath = primaryReportSetting.FilePath.Replace(primaryReportSetting.FileName, "");
            primaryReportSetting.EquipmentName = primaryReportSetting.FileName.Remove(primaryReportSetting.FileName.LastIndexOf('_'), primaryReportSetting.FileName.Length - primaryReportSetting.FileName.LastIndexOf('_'));
            txtBoxFilePath.Text = ofd.FileName;
            List<string> timeStamps = primaryReportSetting.GetTimeStamps();
            cbStartTime.ItemsSource = timeStamps;
            cbEndTime.ItemsSource = timeStamps;

        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnRun_Click(object sender, RoutedEventArgs e)
        {
            RefreshSettings();

            if (!VerifySettings())
            {
                return;
            }

            Cursor cursor = this.Cursor;
            this.Cursor = Cursors.Wait;
               
            primaryReportSetting.ApplyNameFormatting = (bool)chkBoxNameTemplate.IsChecked;

            primaryReportSetting.DateAltFormat = "XX.XX.XXXX";

            if (DateTime.TryParse(cbStartTime.SelectedItem.ToString(), out DateTime dateTime))
            {
                primaryReportSetting.DateAltFormat = $"{dateTime.Month}.{dateTime.Day}.{dateTime.Year}";
                primaryReportSetting.Date = $"{dateTime.Month}/{dateTime.Day}/{dateTime.Year}";
            }
            primaryReportSetting.SavedFilePath = GetNewFileName(primaryReportSetting);

            List<ReportSetting> reportSettings = new List<ReportSetting>();

            if ((bool)chkBoxRunFolder.IsChecked)
            {
                // get folder directory
                string folderPath = primaryReportSetting.FolderPath;
                
                // create a list of folders in that folder
                string[] filesPaths = Directory.GetFiles(folderPath).Where(f => f.Contains(CSV)).ToArray<string>();
                
                foreach (string file in filesPaths)
                {
                    if(File.Exists(file))
                    {
                        ReportSetting reportSetting = new ReportSetting();
                        reportSetting.FilePath = file;
                        reportSetting.SavedFilePath = GetNewFileName(reportSetting);
                        reportSetting.StartTime = primaryReportSetting.StartTime;
                        reportSetting.EndTime = primaryReportSetting.EndTime;
                        reportSetting.Date = primaryReportSetting.Date;
                        reportSetting.FileName = Path.GetFileName(reportSetting.FilePath);
                        reportSetting.EquipmentName = reportSetting.FileName.Remove(reportSetting.FileName.LastIndexOf('_'), reportSetting.FileName.Length - reportSetting.FileName.LastIndexOf('_'));

                        reportSettings.Add(reportSetting);
                    }
                }
            }
            else
            {
                reportSettings.Add(primaryReportSetting);
            }

            string successfulFiles = string.Empty;
            string failedFiles = string.Empty;

            
            // process all files
            foreach (ReportSetting reportSetting in reportSettings)
            {
                if (ProcessFile(reportSetting))
                {
                    successfulFiles += reportSetting.FileName + "\n";
                }
                else
                {
                    failedFiles += reportSetting.FileName + "\n";
                }
            }

            this.Cursor = cursor;

            if (failedFiles.Length > 0)
            {
                MessageBox.Show($"Files processed successfully: \n" +
                    $"{successfulFiles} \n" +
                    $"Failed files: \n" +
                    $"{failedFiles}");
            }
            else
            {
                MessageBox.Show($"Files processed successfully:  \n" +
                    $"{successfulFiles}");
            }
           

        }

        private bool ProcessFile(ReportSetting reportSetting)
        {
            if (reportSetting.GenerateReport())
            {
                //MessageBox.Show($"{sfd.FileName} has been created!", "Report Helper");
                return true;
                
            }
            else
            {
                //MessageBox.Show("An error occured and the file was not created.", "Report Helper");
                return false;
            }
        }

        private string GetNewFileName(ReportSetting reportSetting)
        {
            if (primaryReportSetting.ApplyNameFormatting)
            {
                return Path.Combine(primaryReportSetting.FolderPath, $"{ primaryReportSetting.DateAltFormat} {reportSetting.EquipmentName} TMS Raw Data_Not Reviewed");
            }
            return reportSetting.FilePath.Replace(".csv", "") + "_Edited" + XLSX;

        }

        /// <summary>
        /// Read updated values from user inputs.
        /// </summary>
        private void RefreshSettings()
        {
            primaryReportSetting.ApplyNameFormatting = (bool)chkBoxNameTemplate.IsChecked;
            primaryReportSetting.StartTime = (string)cbStartTime.SelectedItem;
            primaryReportSetting.EndTime = (string)cbEndTime.SelectedItem;
            primaryReportSetting.FilePath = txtBoxFilePath.Text;
            primaryReportSetting.ForceIntervals = (bool)chkBoxForceIntervals.IsChecked;
        }

        /// <summary>
        /// Make sure the user has input valid settings.
        /// </summary>
        /// <returns></returns>
        private bool VerifySettings()
        {
            string errors = string.Empty;

            if (!File.Exists(primaryReportSetting.FilePath))
            {
                errors += $"File {primaryReportSetting.FilePath} not found.\n";
            }

            if (string.IsNullOrEmpty(primaryReportSetting.StartTime))
            {
                errors += "No start time specified.\n";
            }

            if (string.IsNullOrEmpty(primaryReportSetting.EndTime))
            {
                errors += "No end time specified.\n";
            }

            if (errors != string.Empty)
            {
                MessageBox.Show($"The following error(s) were found with the specified settings.\n" +
                    $"{errors}" +
                    $"Please fix and try again.", "Setting Error");

                return false;
            }

            return true;
        }

        private void cbStartTime_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RefreshSettings();
        }

        private void cbEndTime_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RefreshSettings();
        }

        private void txtBoxFilePath_TextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshSettings();
        }

        private void chkBoxNameTemplate_Click(object sender, RoutedEventArgs e)
        {
            RefreshSettings();
        }

        private void chkBoxForceIntervals_Click(object sender, RoutedEventArgs e)
        {
            RefreshSettings();
        }

        private void chkBoxRunFolder_Click(object sender, RoutedEventArgs e)
        {
            RefreshSettings();
        }
    }
}
