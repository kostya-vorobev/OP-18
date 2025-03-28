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

            // Инициализация тестовых данных
            InitializeTestData();

            // Подписка на события
            GoodsGrid.AutoGeneratingColumn += GoodsGrid_AutoGeneratingColumn;
            GoodsGrid.RowEditEnding += GoodsGrid_RowEditEnding;

        }
        private void InitializeTestData()
        {
            // Пример товара для демонстрации
            _viewModel.GoodsItems.Add(new GoodsItem
            {
                Number = 1,
                Name = "Пример товара",
                Code = "001",
                Unit = "шт",
                OKEICode = "796",
                Quantity = 10,
                Price = 100.50m
            });

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
            _viewModel.DocumentData.UpdateTotals(); // Обновляем итоги перед экспортом
            _viewModel.ExportToExcel();
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
            _viewModel.ExportToExcel();
        }

        private void ToggleComboBoxDropDown(object sender, RoutedEventArgs e)
        {
            if (sender is Button button &&
                button.TemplatedParent is ComboBox comboBox)
            {
                comboBox.IsDropDownOpen = !comboBox.IsDropDownOpen;
            }
        }
    }


}
