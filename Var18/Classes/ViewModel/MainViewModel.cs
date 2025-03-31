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

        public ObservableCollection<string> OrganizationNames { get; } = new ObservableCollection<string>
        {
            "ООО \"РОСТИКС\"",
            "ООО \"Мария - Ра\""
        };

        public MainViewModel()
        {

            GoodsItems.CollectionChanged += (s, e) => DocumentData.UpdateTotals();


            DocumentData.HandedOverPosition = HandedOverPositions[0];
            DocumentData.HandedOverName = HandedOverNames[0];
            DocumentData.AcceptedPosition = AcceptedPositions[1];
            DocumentData.AcceptedName = AcceptedNames[1];
            DocumentData.AdminPosition = AdminPositions[2];
            DocumentData.AdminName = AdminNames[2];
            DocumentData.OrganizationName = OrganizationNames[1];
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

        public ObservableCollection<GoodsItem> GoodsItems { get; } = new ObservableCollection<GoodsItem>();
        public decimal GoodsTotal => DocumentData.GoodsTotal;
        public decimal TotalAmount => DocumentData.TotalAmount;


        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

      
        public void LoadDataFromTemplate()
        {
            try
            {
                GoodsItems.Clear();

                Application excelApp = null;
                Workbook workbook = null;

                try
                {
                    excelApp = new Application { Visible = false };
                    workbook = excelApp.Workbooks.Open(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LAW_26677.attach_LAW_26677_18.xlsx"));
                    var worksheet = (Worksheet)workbook.Sheets[1];


                    int startRow = 24;

                    int endRow = 51;
                    GoodsItems.Clear();

                    for (int row = startRow; row <= endRow; row++)
                    {

                        var lineNumber = GetIntValue(worksheet.Cells[row, 1]);


                        var item = new GoodsItem
                        {
                            Number = lineNumber++,
                            Name = GetStringValue(worksheet.Cells[row, 5]), 
                            Code = GetIntValue(worksheet.Cells[row, 19]), 
                            Unit = GetStringValue(worksheet.Cells[row, 22]), 
                            OKEICode = GetIntValue(worksheet.Cells[row, 26]), 
                            Weight = GetDecimalValue(worksheet.Cells[row, 30]), 
                            Quantity = GetDecimalValue(worksheet.Cells[row, 34]), 
                            Price = GetDecimalValue(worksheet.Cells[row, 39]), 
                        };

                        GoodsItems.Add(item);

                    }
                }
                finally
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

                OnPropertyChanged(nameof(GoodsItems));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}");
            }
        }

        private string GetStringValue(Range range)
        {
            return range.Value?.ToString() ?? string.Empty;
        }

        private int GetIntValue(Range range)
        {
            return int.TryParse(range.Value?.ToString(), out int result) ? result : 0;
        }

        private decimal GetDecimalValue(Range range)
        {
            return decimal.TryParse(range.Value?.ToString(), out decimal result) ? result : 0m;
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

                excelApp = new Application { Visible = true, DisplayAlerts = false };
                File.AppendAllText(debugLogPath, "Excel инициализирован\n");


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


                worksheet.Range["AA13"].Value = DocumentData.DocumentNumber;
                worksheet.Range["AI13"].Value = DocumentData.DocumentDate.ToString("dd.MM.yyyy");
                worksheet.Range["A6"].Value = DocumentData.OrganizationName;
                worksheet.Range["A8"].Value = DocumentData.Department;
                worksheet.Range["AO6"].Value = DocumentData.OKPOCode;
                worksheet.Range["AO9"].Value = DocumentData.OKDPCode;


                ClearGoodsRows(worksheet, 24, 51);


                worksheet.Range["G52"].Value = DocumentData.HandedOverPosition;
                worksheet.Range["AB52"].Value = DocumentData.HandedOverName;
                worksheet.Range["G54"].Value = DocumentData.AcceptedPosition;
                worksheet.Range["AB54"].Value = DocumentData.AcceptedName;
                worksheet.Range["R56"].Value = DocumentData.AdminPosition;
                worksheet.Range["AL56"].Value = DocumentData.AdminName;


                int currentRow = 24;
                int numbRow = 1;
                decimal goodsTotalQuantity = 0;
                decimal goodsTotalWeight = 0;
                decimal goodsTotalSum = 0;
                decimal containerTotalQuantity = 0;
                decimal containerTotalWeight = 0;
                decimal containerTotalSum = 0;
                string isName = "Товары";

                foreach (var item in GoodsItems)
                {
                    if (string.IsNullOrEmpty(item.Name)) continue;

                    if (item.Name != "Товары" && item.Name != "Тара" &&
                        item.Name != "Итого" && item.Name != "Всего по акту" &&
                        item.Quantity == 0 && item.Weight == 0 && item.Price == 0)
                    {
                        continue;
                    }

                    worksheet.Cells[currentRow, 1].Value = numbRow++;
                    worksheet.Cells[currentRow, 5].Value = item.Name;
                    if (item.Name == "Тара") isName = item.Name;

                    worksheet.Cells[currentRow, 19].Value = item.Code == 0 ? null : (object)item.Code;
                    worksheet.Cells[currentRow, 22].Value = string.IsNullOrEmpty(item.Unit) ? null : item.Unit;
                    worksheet.Cells[currentRow, 26].Value = item.OKEICode == 0 ? null : (object)item.OKEICode;
                    worksheet.Cells[currentRow, 30].Value = item.Weight == 0 ? null : (object)item.Weight;
                    worksheet.Cells[currentRow, 34].Value = item.Quantity == 0 ? null : (object)item.Quantity;
                    worksheet.Cells[currentRow, 39].Value = item.Price == 0 ? null : (object)item.Price;


                    decimal sum = item.Price * item.Quantity;
                    worksheet.Cells[currentRow, 44].Value = sum == 0 ? null : (object)sum;

                    if (isName == "Товары")
                    {
                        goodsTotalQuantity += item.Quantity;
                        goodsTotalWeight += item.Weight;
                        goodsTotalSum += sum;

                        if (item.Name == "Итого")
                        {
                            worksheet.Cells[currentRow, 19].Value = "Х";
                            worksheet.Cells[currentRow, 22].Value = "Х";
                            worksheet.Cells[currentRow, 26].Value = "Х";
                            worksheet.Cells[currentRow, 30].Value = goodsTotalWeight == 0 ? null : (object)goodsTotalWeight;
                            worksheet.Cells[currentRow, 34].Value = goodsTotalQuantity == 0 ? null : (object)goodsTotalQuantity;
                            worksheet.Cells[currentRow, 39].Value = "Х";
                            worksheet.Cells[currentRow, 44].Value = goodsTotalSum == 0 ? null : (object)goodsTotalSum;
                        }
                    }
                    else
                    {
                        containerTotalQuantity += item.Quantity;
                        containerTotalWeight += item.Weight;
                        containerTotalSum += sum;

                        if (item.Name == "Итого")
                        {
                            worksheet.Cells[currentRow, 19].Value = "Х";
                            worksheet.Cells[currentRow, 22].Value = "Х";
                            worksheet.Cells[currentRow, 26].Value = "Х";
                            worksheet.Cells[currentRow, 30].Value = containerTotalWeight == 0 ? null : (object)containerTotalWeight;
                            worksheet.Cells[currentRow, 34].Value = containerTotalQuantity == 0 ? null : (object)containerTotalQuantity;
                            worksheet.Cells[currentRow, 39].Value = "Х";
                            worksheet.Cells[currentRow, 44].Value = containerTotalSum == 0 ? null : (object)containerTotalSum;
                        }
                    }

                    if (item.Name == "Всего по акту")
                    {
                        decimal overallTotalSum = goodsTotalSum + containerTotalSum;
                        worksheet.Cells[currentRow, 19].Value = "Х";
                        worksheet.Cells[currentRow, 22].Value = "Х";
                        worksheet.Cells[currentRow, 26].Value = "Х";
                        worksheet.Cells[currentRow, 30].Value = "Х";
                        worksheet.Cells[currentRow, 34].Value = "Х";
                        worksheet.Cells[currentRow, 39].Value = "Х";
                        worksheet.Cells[currentRow, 44].Value = overallTotalSum == 0 ? null : (object)overallTotalSum;
                    }

                    File.AppendAllText(debugLogPath, $"Товар {item.Name} записан в строку {currentRow}\n");
                    currentRow++;
                }

                for (int row = 51; row >= 24; row--)
                {
                    Range range = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 44]];
                    bool isEmpty = true;

                    foreach (Range cell in range.Cells)
                    {
                        if (cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                        {
                            isEmpty = false;
                            break;
                        }
                    }

                    if (isEmpty)
                    {
                        ((Range)worksheet.Rows[row]).Delete(XlDeleteShiftDirection.xlShiftUp);
                        File.AppendAllText(debugLogPath, $"Удалена пустая строка {row}\n");
                    }
                }


                string savePath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    $"Акт_передачи_№{DocumentData.DocumentNumber}_{DateTime.Now:yyyyMMddHHmmss}.xlsx");

                workbook.SaveAs(savePath, XlFileFormat.xlOpenXMLWorkbook);
                File.AppendAllText(debugLogPath, $"Файл сохранен: {savePath}\n");

                MessageBox.Show(File.Exists(savePath)
                    ? $"Акт успешно сохранён:\n{savePath}"
                    : "Ошибка при сохранении файла. Проверьте лог: " + debugLogPath);
            }
            catch (Exception ex)
            {
                File.AppendAllText(debugLogPath, $"ОШИБКА: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"Ошибка при экспорте: {ex.Message}\nПодробности в логе: {debugLogPath}");
            }
            finally
            {

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
        // Метод для очистки строк с товарами
        private void ClearGoodsRows(Worksheet worksheet, int startRow, int endRow)
        {
            try
            {
                for (int row = startRow; row <= endRow; row++)
                {
                    worksheet.Cells[row, 1].Value = "";   
                    worksheet.Cells[row, 5].Value = "";   
                    worksheet.Cells[row, 19].Value = "";  
                    worksheet.Cells[row, 22].Value = "";   
                    worksheet.Cells[row, 26].Value = "";  
                    worksheet.Cells[row, 30].Value = "";  
                    worksheet.Cells[row, 34].Value = "";   
                    worksheet.Cells[row, 39].Value = "";  
                    worksheet.Cells[row, 44].Value = "";   
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "excel_debug.log"),
                    $"Ошибка при очистке строк: {ex.Message}\n");
            }
        }
        private void CopyRow(Worksheet worksheet, int sourceRow, int destRow)
        {
            try
            {
                Range sourceRange = worksheet.Rows[sourceRow];
                Range destRange = worksheet.Rows[destRow];
                sourceRange.Copy(destRange);
            }
            catch (Exception ex)
            {
                File.AppendAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "excel_debug.log"),
                    $"Ошибка при копировании строки: {ex.Message}\n");
            }
        }
    }
}