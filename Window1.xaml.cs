using Microsoft.Win32;
using Spire.Xls;
using System;
using System.Data;
using System.Net;
using System.Net.Mail;
using System.Windows;
using System.Windows.Controls;

namespace WordLekcia
{
    public partial class Window1 : Window
    {
        public Window1()
        {
            InitializeComponent();
        }

        private void AddColumnButton_Click(object sender, RoutedEventArgs e)
        {
            string columnName = ColumnNameTextBox.Text.Trim();

            if (!string.IsNullOrEmpty(columnName))
            {
                if (griiiiiid.ItemsSource != null)
                {
                    DataTable table = (griiiiiid.ItemsSource as DataView).Table;
                    if (!table.Columns.Contains(columnName))
                    {
                        table.Columns.Add(columnName);

                        // Обновление DataGrid
                        griiiiiid.ItemsSource = null;
                        griiiiiid.ItemsSource = table.DefaultView;
                    }
                    else
                    {
                        MessageBox.Show("Столбец с таким именем уже существует.");
                    }
                }
                else
                {
                    MessageBox.Show("Нет данных для добавления столбца.");
                }
            }
            else
            {
                MessageBox.Show("Введите название столбца.");
            }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            if (griiiiiid.ItemsSource != null)
            {
                Workbook book = new Workbook();
                book.Worksheets.Clear();
                Worksheet sheet = book.Worksheets.Add("это из проги");

                sheet.InsertDataView(griiiiiid.ItemsSource as DataView, true, 1, 1);

                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    FileName = "это из проги.xlsx",
                    Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    book.SaveToFile(saveFileDialog.FileName);
                }
            }
            else
            {
                MessageBox.Show("Нет данных для экспорта.");
            }
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                Workbook book = new Workbook();
                book.LoadFromFile(openFileDialog.FileName);

                Worksheet worksheet = book.Worksheets[0];
                CellRange range = worksheet.AllocatedRange;
                var table = worksheet.ExportDataTable(range, true);
                griiiiiid.ItemsSource = table.DefaultView;
            }
        }

        private void SendButton_Click(object sender, RoutedEventArgs e)
        {
            Window2 window2 = new Window2();
            window2.Show();
            Close();
        }
    }
}
