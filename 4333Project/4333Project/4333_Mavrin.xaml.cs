using System;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Globalization;

namespace _4333Project
{
    public partial class _4333_Mavrin : System.Windows.Window
    {
        public _4333_Mavrin()
        {
            InitializeComponent();
        }

        // Обработчик для импорта данных из JSON
        private void ImportJsonData_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Окно для выбора файла
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "JSON Files|*.json";
                if (openFileDialog.ShowDialog() == true)
                {
                    string filePath = openFileDialog.FileName;
                    string jsonData = File.ReadAllText(filePath);
                    var clients = JsonConvert.DeserializeObject<List<Client>>(jsonData);

                    using (ArleEntities dbContext = new ArleEntities())
                    {
                        foreach (var client in clients)
                        {
                            // Обрабатываем дату
                            DateTime? birthDate = null;
                            if (!string.IsNullOrEmpty(client.BirthDate))
                            {
                                birthDate = DateTime.ParseExact(client.BirthDate, "dd.MM.yyyy", CultureInfo.InvariantCulture);
                            }

                            // Создаем новый объект для БД
                            var newClient = new Clients
                            {
                                ID = client.Id,
                                FullName = client.FullName,
                                ClientCode = client.CodeClient,
                                BirthDate = client.BirthDate == null ? DateTime.MinValue : DateTime.ParseExact(client.BirthDate, "dd.MM.yyyy", CultureInfo.InvariantCulture),

                                PostalCode = client.Index,
                                City = client.City,
                                Street = client.Street,
                                House = client.Home.ToString(),
                                Apartment = client.Kvartira.ToString(),
                                Email = client.E_mail
                            };

                            dbContext.Clients.Add(newClient);
                        }
                        dbContext.SaveChanges();
                    }

                    MessageBox.Show("Данные успешно импортированы!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        // Обработчик для экспорта данных в Word
        private void ExportDataToWord_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (ArleEntities dbContext = new ArleEntities())
                {
                    var allClients = dbContext.Clients.ToList();
                    var groupedClients = allClients.GroupBy(c => c.Street); // Группировка по улице

                    // Окно для выбора места назначения
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Word Document (*.docx)|*.docx";
                    if (saveFileDialog.ShowDialog() == true)
                    {
                        string savePath = saveFileDialog.FileName;

                        // Создаем новый документ Word
                        Word.Application wordApp = new Word.Application();
                        Word.Document doc = wordApp.Documents.Add();

                        foreach (var group in groupedClients)
                        {
                            // Добавляем заголовок для группы (например, улицы)
                            Word.Paragraph para = doc.Content.Paragraphs.Add();
                            para.Range.Text = $"Улица: {group.Key}";
                            para.Range.InsertParagraphAfter();

                            // Добавляем данные клиентов в текущую группу
                            foreach (var client in group)
                            {
                                Word.Paragraph clientPara = doc.Content.Paragraphs.Add();
                                clientPara.Range.Text = $"{client.FullName}, Код клиента: {client.ClientCode}, Email: {client.Email}";
                                clientPara.Range.InsertParagraphAfter();
                            }
                        }

                        // Сохраняем файл в указанном месте
                        doc.SaveAs2(savePath);
                        doc.Close();
                        wordApp.Quit();

                        MessageBox.Show("Данные успешно экспортированы в Word!");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }
    }
}

// Класс, соответствующий структуре JSON данных
public class Client
    {
        public int Id { get; set; }
        public string FullName { get; set; }
        public string CodeClient { get; set; }
        public string BirthDate { get; set; } // Дата в формате строки
        public string Index { get; set; }
        public string City { get; set; }
        public string Street { get; set; }
        public int Home { get; set; }
        public int Kvartira { get; set; }
        public string E_mail { get; set; }
    }

