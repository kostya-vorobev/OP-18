using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using Var18.Classes.ModelData;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Var18.Classes
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private DocumentData _documentData = new DocumentData();

        // Коллекции для ComboBox
        public ObservableCollection<string> HandedOverPositions { get; } = new ObservableCollection<string>
        {
            "Кладовщик",
            "Старший кладовщик",
            "Заведующий складом"
        };

        public ObservableCollection<string> AcceptedPositions { get; } = new ObservableCollection<string>
        {
            "Менеджер склада",
            "Заместитель директора",
            "Директор"
        };

        public ObservableCollection<string> AdminPositions { get; } = new ObservableCollection<string>
        {
            "Начальник отдела",
            "Заместитель директора",
            "Директор"
        };

        public ObservableCollection<string> HandedOverNames { get; } = new ObservableCollection<string>
        {
            "Иванов Иван Иванович",
            "Петров Петр Петрович",
            "Сидорова Анна Владимировна"
        };

        public ObservableCollection<string> AcceptedNames { get; } = new ObservableCollection<string>
        {
            "Смирнов Александр Васильевич",
            "Кузнецова Елена Дмитриевна",
            "Федоров Михаил Сергеевич"
        };

        public ObservableCollection<string> AdminNames { get; } = new ObservableCollection<string>
        {
            "Николаева Ольга Игоревна",
            "Волков Денис Александрович",
            "Павлова Татьяна Викторовна"
        };

        public MainViewModel()
        {
            // Инициализация данных
            GoodsItems.CollectionChanged += (s, e) => DocumentData.UpdateTotals();

            // Установка значений по умолчанию для подписей
            DocumentData.HandedOverPosition = HandedOverPositions[0];
            DocumentData.HandedOverName = HandedOverNames[0];
            DocumentData.AcceptedPosition = AcceptedPositions[1];
            DocumentData.AcceptedName = AcceptedNames[1];
            DocumentData.AdminPosition = AdminPositions[2];
            DocumentData.AdminName = AdminNames[2];
        }

        public DocumentData DocumentData
        {
            get => _documentData;
            set
            {
                _documentData = value;
                OnPropertyChanged(nameof(DocumentData));
            }
        }

        public ObservableCollection<GoodsItem> GoodsItems => DocumentData.GoodsItems;

        public decimal GoodsTotal => DocumentData.GoodsTotal;
        public decimal TotalAmount => DocumentData.TotalAmount;


        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public bool Validate()
        {
            if (string.IsNullOrWhiteSpace(DocumentData.OrganizationName))
            {
                MessageBox.Show("Укажите название организации", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(DocumentData.DocumentNumber))
            {
                MessageBox.Show("Укажите номер документа", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (DocumentData.DocumentDate == default)
            {
                MessageBox.Show("Укажите дату документа", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (GoodsItems.Count == 0)
            {
                MessageBox.Show("Добавьте хотя бы один товар", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(DocumentData.HandedOverName) ||
                string.IsNullOrWhiteSpace(DocumentData.AcceptedName) ||
                string.IsNullOrWhiteSpace(DocumentData.AdminName))
            {
                MessageBox.Show("Заполните все подписи", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            return true;
        }

        public void ExportToExcel()
        {
            string debugLogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "excel_debug.log");
            File.WriteAllText(debugLogPath, $"Начало экспорта {DateTime.Now}\n");

            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                // Инициализация Excel
                excelApp = new Application { Visible = true, DisplayAlerts = false };
                File.AppendAllText(debugLogPath, "Excel инициализирован\n");

                // Открытие шаблона
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LAW_26677.attach_LAW_26677_18.xlsx");
                if (!File.Exists(templatePath))
                {
                    string error = $"Файл шаблона не найден: {templatePath}";
                    File.AppendAllText(debugLogPath, error + "\n");
                    MessageBox.Show(error);
                    return;
                }

                workbook = excelApp.Workbooks.Open(templatePath);
                worksheet = (Worksheet)workbook.Sheets[1];
                File.AppendAllText(debugLogPath, $"Шаблон открыт: {templatePath}\n");

                // 1. Заполнение шапки документа
                worksheet.Range["AA13"].Value = DocumentData.DocumentNumber;      // Номер документа
                worksheet.Range["AI13"].Value = DocumentData.DocumentDate.ToString("dd.MM.yyyy"); // Дата
                worksheet.Range["A6"].Value = DocumentData.OrganizationName;    // Название организации
                worksheet.Range["A8"].Value = DocumentData.Department;      // Структурное подразделение
                worksheet.Range["AO6"].Value = DocumentData.OKPOCode;      // Структурное подразделение
                worksheet.Range["AO9"].Value = DocumentData.OKDPCode;      // Структурное подразделение

                // 2. Заполнение таблицы товаров (начинается с 20 строки)
                int startRow = 20;
                foreach (var item in GoodsItems)
                {
                    if (string.IsNullOrEmpty(item.Name)) continue;

                    // Основные колонки товаров
                    worksheet.Cells[startRow, 1].Value = item.Number;           // A - № п/п
                    worksheet.Cells[startRow, 4].Value = item.Name;             // D - Наименование
                    worksheet.Cells[startRow, 19].Value = item.Code;            // S - Код товара
                    worksheet.Cells[startRow, 22].Value = item.Unit;            // V - Единица измерения
                    worksheet.Cells[startRow, 25].Value = item.OKEICode;        // Y - Код ОКЕИ
                    worksheet.Cells[startRow, 33].Value = item.Quantity;        // AG - Количество
                    worksheet.Cells[startRow, 40].Value = item.Price;           // AN - Цена
                    worksheet.Cells[startRow, 45].Value = item.Price * item.Quantity; // AS - Сумма

                    File.AppendAllText(debugLogPath, $"Товар {item.Name} записан в строку {startRow}\n");
                    startRow++;
                }

                // 3. Заполнение подписей
                // Сдал
                worksheet.Range["G52"].Value = DocumentData.HandedOverPosition;  // Должность
                worksheet.Range["AB52"].Value = DocumentData.HandedOverName;    // ФИО

                // Принял
                worksheet.Range["G54"].Value = DocumentData.AcceptedPosition;
                worksheet.Range["AB54"].Value = DocumentData.AcceptedName;

                // Представитель администрации
                worksheet.Range["R56"].Value = DocumentData.AdminPosition;
                worksheet.Range["AL56"].Value = DocumentData.AdminName;

                // 4. Сохранение файла
                string savePath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    $"Акт_передачи_№{DocumentData.DocumentNumber}_{DateTime.Now:yyyyMMddHHmmss}.xlsx");

                workbook.SaveAs(savePath, XlFileFormat.xlOpenXMLWorkbook);
                File.AppendAllText(debugLogPath, $"Файл сохранен: {savePath}\n");

                if (File.Exists(savePath))
                {
                    MessageBox.Show($"Акт успешно сохранён:\n{savePath}");
                }
                else
                {
                    MessageBox.Show("Ошибка при сохранении файла. Проверьте лог: " + debugLogPath);
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(debugLogPath, $"ОШИБКА: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"Ошибка при экспорте: {ex.Message}\nПодробности в логе: {debugLogPath}");
            }
            finally
            {
                // Корректное закрытие Excel
                try
                {
                    if (workbook != null)
                    {
                        workbook.Close(false);
                        Marshal.ReleaseComObject(workbook);
                    }
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                catch (Exception ex)
                {
                    File.AppendAllText(debugLogPath, $"Ошибка при закрытии Excel: {ex.Message}\n");
                }
            }
        }
    }
}