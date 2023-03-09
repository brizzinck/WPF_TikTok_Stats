using System;
using System.Collections.Generic;
using System.Windows;
using HtmlAgilityPack;
using OfficeOpenXml;
using System.IO;
using Microsoft.Win32;
using System.Threading.Tasks;

namespace TIkTokStats
{
    public partial class MainWindow : Window
    {
        private int _index = 0;
        private ExcelPackage _package;
        ExcelWorksheet _worksheet;

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
                    _index = 0;
                    OpenFileDialog openFileDialogTemp = new OpenFileDialog();
                    bool? result = openFileDialogTemp.ShowDialog();
                    if (result == false || result == null) return;
                    openFileDialog = openFileDialogTemp;
                }
                if (openFileDialog != null)
                {
                    Stream stream  = openFileDialog.OpenFile();
                    _package = GetExcel(new FileInfo(openFileDialog.FileName));
                    if (_package != null)
                    {
                        _worksheet = await CheckExcelPageAsync(_package);
                        using (_package)
                        {
                            await Task.Run(() => stream.Close());
                            string userUrl = String.Empty;
                            List<string> usersUrl = new List<string>();
                            do
                            {
                                userUrl = GetUrl(_worksheet);
                                if (string.IsNullOrEmpty(userUrl)) break;
                                await WriteAllInfo(openFileDialog, stream, userUrl);
                                Progress.Text = "Пройшло " + (_index + 1) + " елементів";
                                _index++;

                            } while (!string.IsNullOrEmpty(userUrl));
                            SetBaseInfo();
                        }
                        stream.Close();
                        ShowMessageBox();
                        Progress.Text = string.Empty;
                    }
                }
            }
            catch (System.NullReferenceException)
            {
                if (_package == null)
                {
                    _index = 0;
                    return;
                }
                _index += 1;
                ViewInfoOfUser(false, openFileDialog);
                await ErrorSaveAsync();
            }
        }

        private async Task<bool> WriteAllInfo(OpenFileDialog openFileDialog, Stream stream, string userUrl)
        {
            var web = new HtmlWeb();
            if (IsUrlValid(userUrl) == true)
            {
                var doc = await web.LoadFromWebAsync(userUrl);
                if (doc != null)
                {
                    List<string> userInfo = new List<string>
                    {
                        doc.DocumentNode.SelectSingleNode("//strong[@title='Likes']").InnerHtml,
                        doc.DocumentNode.SelectSingleNode("//strong[@title='Followers']").InnerHtml
                    };
                    for (int j = 0; j < userInfo.Count; j++)
                    {
                        await SaveInExcelAsync(stream, openFileDialog.FileName, _index, j, userInfo[j]);
                    }
                }
                else
                    await ErrorSaveAsync();
                return true;
            }
            return false;
        }
        private static ExcelPackage GetExcel(FileInfo newFile)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                return new ExcelPackage(newFile);
            }
            catch (System.IO.InvalidDataException)
            {
                ShowErrorInvalidDataException();
            }
            return null;
        }
        private string GetUrl(ExcelWorksheet worksheet)
        {
            string userUrl;
            var cell = worksheet.Cells[_index + 2, 1];
            userUrl = cell.Value?.ToString() ?? string.Empty;
            return userUrl;
        }
        public static bool IsUrlValid(string url)
        {
            if (Uri.TryCreate(url, UriKind.Absolute, out Uri uriResult))
            {
                return (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
            }
            return false;
        }
        private async Task ErrorSaveAsync()
        {
            await Task.Run(() =>
            {
                MessageBox.Show("Помилка в запису одного файлу на " + (_index + 1) + 
                    " рядку!\nМожливо силка не працює!", "Помилка", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
            });
        }
        private async Task SaveInExcelAsync(Stream stream, string path, int indexCellsX, int indexCellsY, string value)
        {
            var newFile = new FileInfo(path);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(newFile))
            {
                await Task.Run(() => stream.Close());
                ExcelWorksheet worksheet = await CheckExcelPageAsync(package);
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
        private void SetBaseInfo()
        {
            _worksheet.Cells[1, 1].Value = "User Url";
            _worksheet.Cells[1, 2].Value = "User Likes";
            _worksheet.Cells[1, 3].Value = "User Followers";
        }
        private static void ShowMessageBox()
        {
            MessageBox.Show("Дані оновлено!", "Виконано", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        private static void ShowErrorInvalidDataException()
        {
            MessageBox.Show("Здається, ти обрав не файл Excel", "Помилка", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        private void Background_MediaEnded(object sender, RoutedEventArgs e)
        {
            Background.Position = TimeSpan.MinValue;
        }
    }
}