using System;
using System.Linq;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Collections.Generic;

namespace _4333Project
{
    public partial class _4333_Mavrin : System.Windows.Window // Используем полное имя класса System.Windows.Window
    {
        public _4333_Mavrin()
        {
            InitializeComponent();
        }

        // Обработчик для кнопки "Импорт данных"
        private void ImportClientsData_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog()
                {
                    DefaultExt = "*.xls;*.xlsx",
                    Filter = "файл Excel (Clients.xlsx)|*.xlsx",
                    Title = "Выберите файл базы данных"
                };

                if (!(ofd.ShowDialog() == true))
                    return;

                string[,] list;
                Microsoft.Office.Interop.Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application(); // Явно указываем Excel.Application
                Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
                Worksheet ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];

                var lastCell = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
                int _columns = (int)lastCell.Column;
                int _rows = (int)lastCell.Row;

                list = new string[_rows, _columns];

                for (int j = 0; j < _columns; j++)
                {
                    for (int i = 0; i < _rows; i++)
                    {
                        list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                    }
                }

                int lastRow = 0;
                for (int i = 0; i < _rows; i++)
                {
                    if (list[i, 0] != string.Empty)
                    {
                        lastRow = i;
                    }
                }

                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
                GC.Collect();

                // Добавление данных в базу данных
                using (ArleEntities dbContext = new ArleEntities())
                {
                    for (int i = 1; i <= lastRow; i++)
                    {
                        var client = new Clients
                        {
                            FullName = list[i, 0],
                            ClientCode = list[i, 1],
                            BirthDate = DateTime.ParseExact(list[i, 2], "dd.MM.yyyy", null),
                            PostalCode = list[i, 3],
                            City = list[i, 4],
                            Street = list[i, 5],
                            House = list[i, 6],
                            Apartment = list[i, 7],
                            Email = list[i, 8]
                        };

                        dbContext.Clients.Add(client);
                    }

                    dbContext.SaveChanges();
                }

                MessageBox.Show("Успешное импортирование данных", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Обработчик для кнопки "Экспорт в Excel"
        private void ExportClientsToExcel_Click(object sender, RoutedEventArgs e)
        {

            try

            {

                var allClients = new List<Clients>();



                using (ArleEntities dbContext = new
      ArleEntities())

                {

                    allClients = dbContext.Clients.ToList();

                }



                var groupedClients =
      allClients.GroupBy(c => c.Street);



                var app = new Microsoft.Office.Interop.Excel.Application();

                app.SheetsInNewWorkbook =
      groupedClients.Count();

                Microsoft.Office.Interop.Excel.Workbook workbook =
       app.Workbooks.Add(Type.Missing);



                int sheetIndex = 1;



                foreach (var group in groupedClients)

                {

                    Microsoft.Office.Interop.Excel.Worksheet worksheet =
       app.Worksheets.Item[sheetIndex];

                    worksheet.Name = group.Key;



                    worksheet.Cells[1, 1] = "ФИО";

                    worksheet.Cells[1, 2] = "Код клиента";

                    worksheet.Cells[1, 3] = "Дата рождения";

                    worksheet.Cells[1, 4] = "Индекс";

                    worksheet.Cells[1, 5] = "Город";

                    worksheet.Cells[1, 6] = "Улица";

                    worksheet.Cells[1, 7] = "Дом";

                    worksheet.Cells[1, 8] = "Квартира";

                    worksheet.Cells[1, 9] = "E-mail";



                    int startRowIndex = 2;



                    foreach (var client in group)

                    {


                        worksheet.Cells[startRowIndex, 1] = client.FullName;

                        worksheet.Cells[startRowIndex, 2] = client.ClientCode;


                        worksheet.Cells[startRowIndex, 3] =
                        client.BirthDate.ToString("dd.MM.yyyy");

                        worksheet.Cells[startRowIndex, 4] = client.PostalCode;


                        worksheet.Cells[startRowIndex, 5] = client.City;

                        worksheet.Cells[startRowIndex, 6] = client.Street;


                        worksheet.Cells[startRowIndex, 7] = client.House;

                        worksheet.Cells[startRowIndex, 8] = client.Apartment;


                        worksheet.Cells[startRowIndex, 9] = client.Email;



                        startRowIndex++;

                    }



                    worksheet.Columns.AutoFit();



                    sheetIndex++;

                }



                app.Visible = true;



                MessageBox.Show("Успешное экспортирование данных", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);

            }

            catch (Exception ex)

            {

                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }

        }
    }
}
