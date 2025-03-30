using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Var18.Classes.ModelData;
using Var18.Classes;

namespace Var18
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private MainViewModel _viewModel;
        public MainWindow()
        {
            InitializeComponent();
            _viewModel = new MainViewModel();
            DataContext = _viewModel;
            _viewModel.LoadDataFromTemplate(); // Загружаем данные при старте
            

            // Инициализация тестовых данных
            InitializeTestData();

            // Подписка на события
            GoodsGrid.AutoGeneratingColumn += GoodsGrid_AutoGeneratingColumn;
            GoodsGrid.RowEditEnding += GoodsGrid_RowEditEnding;

            // Устанавливаем источник данных
            GoodsGrid.ItemsSource = _viewModel.GoodsItems;
        

    }
        private void InitializeTestData()
        {

            // Устанавливаем значения по умолчанию
            _viewModel.DocumentData.HandedOverPosition = _viewModel.HandedOverPositions[0];
            _viewModel.DocumentData.HandedOverName = _viewModel.HandedOverNames[0];

            _viewModel.DocumentData.AcceptedPosition = _viewModel.AcceptedPositions[1];
            _viewModel.DocumentData.AcceptedName = _viewModel.AcceptedNames[1];

            _viewModel.DocumentData.AdminPosition = _viewModel.AdminPositions[2];
            _viewModel.DocumentData.AdminName = _viewModel.AdminNames[2];
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Обновляем номера строк при загрузке
            UpdateRowNumbers();
        }

        private void ExportCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {

        }

        private void GoodsGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            // Можно настроить автоматически генерируемые колонки
        }

        private void GoodsGrid_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            // Обновляем номера строк после редактирования
            UpdateRowNumbers();
        }

        private void UpdateRowNumbers()
        {
            for (int i = 0; i < _viewModel.GoodsItems.Count; i++)
            {
                _viewModel.GoodsItems[i].Number = i + 1;
            }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            // Сбрасываем все подсветки
            ResetAllValidation();

            // Проверяем обязательные поля
            bool hasErrors = false;

            // Проверка названия организации
            if (string.IsNullOrWhiteSpace(_viewModel.DocumentData.OrganizationName))
            {
                SetErrorStyle(OrganizationName);
                hasErrors = true;
            }

            // Проверка названия организации
            if (string.IsNullOrWhiteSpace(_viewModel.DocumentData.Department))
            {
                SetErrorStyle(Department);
                hasErrors = true;
            }

            // Проверка названия организации
            if (string.IsNullOrWhiteSpace(_viewModel.DocumentData.OKPOCode))
            {
                SetErrorStyle(OKPOCode);
                hasErrors = true;
            }

            // Проверка названия организации
            if (string.IsNullOrWhiteSpace(_viewModel.DocumentData.OKDPCode))
            {
                SetErrorStyle(OKDPCode);
                hasErrors = true;
            }

            // Проверка номера документа
            if (string.IsNullOrWhiteSpace(_viewModel.DocumentData.DocumentNumber))
            {
                SetErrorStyle(DocumentNumber);
                hasErrors = true;
            }

            // Проверка даты документа
            if (_viewModel.DocumentData.DocumentDate == default)
            {
                SetErrorStyle(DocDate);
                hasErrors = true;
            }

            // Проверка товаров
            if (_viewModel.GoodsItems.Count == 0)
            {
                MessageBox.Show("Добавьте хотя бы один товар", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                hasErrors = true;
            }

            // Проверка подписей
            if (string.IsNullOrWhiteSpace(_viewModel.DocumentData.HandedOverName))
            {
                SetErrorStyle(HandedOverName);
                hasErrors = true;
            }

            if (string.IsNullOrWhiteSpace(_viewModel.DocumentData.AcceptedName))
            {
                SetErrorStyle(AcceptedName);
                hasErrors = true;
            }

            if (string.IsNullOrWhiteSpace(_viewModel.DocumentData.AdminName))
            {
                SetErrorStyle(AdminName);
                hasErrors = true;
            }

            // Если есть ошибки - не продолжаем
            if (hasErrors)
            {
                MessageBox.Show("Заполните все обязательные поля", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            _viewModel.ExportToExcel();
        }

        private void SetErrorStyle(Control control)
        {
            if (control is TextBox textBox)
            {
                textBox.BorderBrush = Brushes.Red;
                textBox.BorderThickness = new Thickness(0, 0, 0, 2);
            }
            else if (control is DatePicker datePicker)
            {
                datePicker.BorderBrush = Brushes.Red;
                datePicker.BorderThickness = new Thickness(0, 0, 0, 2);
            }
            else if (control is ComboBox comboBox)
            {
                comboBox.BorderBrush = Brushes.Red;
                comboBox.BorderThickness = new Thickness(0, 0, 0, 2);
            }
        }

        private void ResetAllValidation()
        {
            // Сбрасываем стили для всех полей
            var controls = new Control[]
            {
        OrganizationName, DocumentNumber, DocDate,
        HandedOverName, AcceptedName, AdminName
            };

            foreach (var control in controls)
            {
                if (control is TextBox textBox)
                {
                    textBox.BorderBrush = (SolidColorBrush)new BrushConverter().ConvertFrom("#FFBDBDBD");
                    textBox.ToolTip = null;
                }
                else if (control is DatePicker datePicker)
                {
                    datePicker.BorderBrush = (SolidColorBrush)new BrushConverter().ConvertFrom("#FFBDBDBD");
                    datePicker.ToolTip = null;
                }
                else if (control is ComboBox comboBox)
                {
                    comboBox.BorderBrush = (SolidColorBrush)new BrushConverter().ConvertFrom("#FFBDBDBD");
                    comboBox.ToolTip = null;
                }
            }
        }

        private void ToggleComboBoxDropDown(object sender, RoutedEventArgs e)
        {
            if (sender is Button button &&
                button.TemplatedParent is ComboBox comboBox)
            {
                comboBox.IsDropDownOpen = !comboBox.IsDropDownOpen;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }
    }


}
