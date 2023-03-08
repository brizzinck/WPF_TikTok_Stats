using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using HtmlAgilityPack;
using OfficeOpenXml;
using System.IO;
using Microsoft.Win32;

namespace TIkTokStats
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Background.Play();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ViewInfoOfUser();
        }

        private void ViewInfoOfUser()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            bool? result = openFileDialog.ShowDialog();
            if (result == false || result == null) return;
            List<string> userUrl = ParseUrlUserFromExcel(openFileDialog);
            List<string>[] userInfo = new List<string>[userUrl.Count];
            for (int i = 0; i < userUrl.Count; i++)
            {
                var web = new HtmlWeb();
                var doc = web.Load(userUrl[i]);
                userInfo[i] = new List<string>
                {
                    doc.DocumentNode.SelectSingleNode("//strong[@title='Likes']").InnerHtml,
                    doc.DocumentNode.SelectSingleNode("//strong[@title='Followers']").InnerHtml
                };
            }
            SaveInExcel(openFileDialog, userInfo);
        }

        private void SaveInExcel(OpenFileDialog openFileDialog, List<string>[] userInfo)
        {
            Stream stream;
            if ((stream = openFileDialog.OpenFile()) != null)
            {
                var newFile = new FileInfo(openFileDialog.FileName);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(newFile))
                {
                    stream.Close();
                    ExcelWorksheet worksheet = CheckExcelPage(package);
                    worksheet.Cells[1, 1].Value = "User Url";
                    worksheet.Cells[1, 2].Value = "User Likes";
                    worksheet.Cells[1, 3].Value = "User Followers";
                    for (int i = 0; i < userInfo.Length; i++)
                    {
                        for (int j = 0; j < userInfo[i].Count; j++)
                        {
                            worksheet.Cells[i + 2, j + 2].Value = userInfo[i][j];
                        }
                    }
                    package.Save();
                    ShowMessageBox();
                }
            }
        }
        private List<string> ParseUrlUserFromExcel(OpenFileDialog openFileDialog)
        {
            Stream stream;
            if ((stream = openFileDialog.OpenFile()) != null)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var newFile = new FileInfo(openFileDialog.FileName);
                using (var package = new ExcelPackage(newFile))
                {
                    stream.Close();
                    ExcelWorksheet worksheet = CheckExcelPage(package);
                    string userUrl = String.Empty;
                    List<string> usersUrl = new List<string>();
                    int index = 2;
                    do
                    {
                        var cell = worksheet.Cells[index, 1];
                        userUrl = cell.Value?.ToString() ?? string.Empty;
                        if (string.IsNullOrEmpty(userUrl)) break;
                        usersUrl.Add(userUrl);
                        index++;
                    } while (!string.IsNullOrEmpty(userUrl));
                    return usersUrl;
                }
            }
            return null;
        }
        private static ExcelWorksheet CheckExcelPage(ExcelPackage package)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            if (worksheet == null)
                package.Workbook.Worksheets.Add("WorkSheet1");
            worksheet = package.Workbook.Worksheets[0];
            return worksheet;
        }
        private static void ShowMessageBox()
        {
            MessageBox.Show("Дані оновлено", "Виконано", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        private void Background_MediaEnded(object sender, RoutedEventArgs e)
        {
            Background.Position = TimeSpan.MinValue;
        }
    }
}