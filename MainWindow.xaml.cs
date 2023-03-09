using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using HtmlAgilityPack;
using OfficeOpenXml;
using System.IO;
using Microsoft.Win32;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Reflection;
using VisioForge.Core.VideoEdit.Timeline.Timeline;

namespace TIkTokStats
{
    public partial class MainWindow : Window
    {
        private int index = 0;
        public MainWindow()
        {
            InitializeComponent();
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Background.Play();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ViewInfoOfUser(true, null);
        }

        private async void ViewInfoOfUser(bool isNew, OpenFileDialog openFileDialog)
        {
            try
            {
                if (isNew)
                {
                    index = 0;
                    OpenFileDialog openFileDialogTemp = new OpenFileDialog();
                    bool? result = openFileDialogTemp.ShowDialog();
                    if (result == false || result == null) return;
                    openFileDialog = openFileDialogTemp;
                }
                if (openFileDialog != null)
                {
                    Stream stream = stream = openFileDialog.OpenFile();
                    List<string> userUrl = await ParseUrlUserFromExcelAsync(openFileDialog);
                    List<string>[] userInfo = new List<string>[userUrl.Count];
                    for (int i = index; i < userUrl.Count; i++)
                    {
                        var web = new HtmlWeb();
                        var doc = await web.LoadFromWebAsync(userUrl[i]);
                        userInfo[i] = new List<string>
                        {
                            doc.DocumentNode.SelectSingleNode("//strong[@title='Likes']").InnerHtml,
                            doc.DocumentNode.SelectSingleNode("//strong[@title='Followers']").InnerHtml
                        };
                        for (int j = 0; j < userInfo[i].Count; j++)
                        {
                            await SaveInExcelAsync(stream, openFileDialog.FileName, i, j, userInfo[i][j]);
                        }
                        Progress.Text = ((float)(i + 1) / (float)userUrl.Count * 100).ToString("0") + "%";
                        index = i;
                    }
                    stream.Close();
                    ShowMessageBox();
                }
            }
            catch (System.NullReferenceException)
            {
                ViewInfoOfUser(false, openFileDialog);
                MessageBox.Show("Помилка в запусу одного файлу на " + (index + 2) + " рядку", 
                    "Помилка", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private async Task<List<string>> ParseUrlUserFromExcelAsync(OpenFileDialog openFileDialog)
        {
            Stream stream;
            try
            {
                if ((stream = openFileDialog.OpenFile()) != null)
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    var newFile = new FileInfo(openFileDialog.FileName);
                    using (var package = new ExcelPackage(newFile))
                    {
                        await Task.Run(() => stream.Close());
                        ExcelWorksheet worksheet = await CheckExcelPageAsync(package); 
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
            }
            catch (System.NullReferenceException)
            {
                MessageBox.Show("Помилка в парсенгу url", "Помилка", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            return null;
        }

        private async Task SaveInExcelAsync(Stream stream, string path, int indexCellsX, int indexCellsY, string value)
        {
            var newFile = new FileInfo(path);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(newFile))
            {
                await Task.Run(() => stream.Close());
                ExcelWorksheet worksheet = await CheckExcelPageAsync(package);
                SetBaseInfo(worksheet);
                worksheet.Cells[indexCellsX + 2, indexCellsY + 2].Value = value;
                await Task.Run(() => package.Save());
            }
        }
        private static async Task<ExcelWorksheet> CheckExcelPageAsync(ExcelPackage package)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            if (worksheet == null)
            {
                await Task.Run(() => package.Workbook.Worksheets.Add("WorkSheet1"));
                worksheet = package.Workbook.Worksheets[0];
            }
            return worksheet;
        }
        private static void SetBaseInfo(ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1].Value = "User Url";
            worksheet.Cells[1, 2].Value = "User Likes";
            worksheet.Cells[1, 3].Value = "User Followers";
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