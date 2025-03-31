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
    public partial class MainWindow : Window
    {

        private readonly MainViewModel _viewModel;
        public MainWindow()
        {
            InitializeComponent();
            _viewModel = new MainViewModel();
            DataContext = _viewModel;
            _viewModel.LoadDataFromTemplate();


            InitializeTestData();


            GoodsGrid.AutoGeneratingColumn += GoodsGrid_AutoGeneratingColumn;
            GoodsGrid.RowEditEnding += GoodsGrid_RowEditEnding;


            GoodsGrid.ItemsSource = _viewModel.GoodsItems;


        }
        private void InitializeTestData()
        {
            _viewModel.DocumentData.HandedOverPosition = _viewModel.HandedOverPositions[0];

            _viewModel.DocumentData.AcceptedPosition = _viewModel.AcceptedPositions[1];

            _viewModel.DocumentData.AdminPosition = _viewModel.AdminPositions[2];

            _viewModel.DocumentData.OrganizationName = _viewModel.OrganizationNames[1];
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            UpdateRowNumbers();
        }

        private void ExportCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {

        }

        private void GoodsGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {

        }

        private void GoodsGrid_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {

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

            ResetAllValidation();


            bool hasErrors = false;


            if (string.IsNullOrWhiteSpace(_viewModel.DocumentData.OrganizationName))
            {
                SetErrorStyle(OrganizationName);
                hasErrors = true;
            }


            if (string.IsNullOrWhiteSpace(_viewModel.DocumentData.Department))
            {
                SetErrorStyle(Department);
                hasErrors = true;
            }


            if (string.IsNullOrWhiteSpace(_viewModel.DocumentData.OKPOCode))
            {
                SetErrorStyle(OKPOCode);
                hasErrors = true;
            }

            if (string.IsNullOrWhiteSpace(_viewModel.DocumentData.OKDPCode))
            {
                SetErrorStyle(OKDPCode);
                hasErrors = true;
            }


            if (string.IsNullOrWhiteSpace(_viewModel.DocumentData.DocumentNumber))
            {
                SetErrorStyle(DocumentNumber);
                hasErrors = true;
            }


            if (_viewModel.DocumentData.DocumentDate == default)
            {
                SetErrorStyle(DocDate);
                hasErrors = true;
            }


            if (_viewModel.GoodsItems.Count == 0)
            {
                MessageBox.Show("Добавьте хотя бы один товар", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                hasErrors = true;
            }

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

        private void OrganizationName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void OrganizationName_LostFocus(object sender, RoutedEventArgs e)
        {
            UpdateCode(OrganizationName.Text);
        }

        private void OrganizationName_Loaded(object sender, RoutedEventArgs e)
        {
            UpdateCode(OrganizationName.Text);
        }
        private void UpdateCode(string OrgName)
        {
            switch (OrgName)
            {
                case "ООО \"Мария - Ра\"":
                    OKPOCode.Text = "10036039";
                    OKPOCode.IsReadOnly = true;
                    break;
                case "ООО \"РОСТИКС\"":
                    OKPOCode.Text = "46737022";
                    OKPOCode.IsReadOnly = true;
                    break;
                default:
                    OKPOCode.IsReadOnly = false;
                    OKDPCode.IsReadOnly = false;
                    break;
            }
        }

        private void HandedOverPosition_Loaded(object sender, RoutedEventArgs e)
        {
            var index = _viewModel.HandedOverPositions.IndexOf(HandedOverPosition.Text);
            if (index != -1)
                HandedOverName.SelectedIndex = HandedOverPosition.SelectedIndex;

        }

        private void HandedOverName_Loaded(object sender, RoutedEventArgs e)
        {
            var index = _viewModel.HandedOverNames.IndexOf(HandedOverName.Text);
            if (index != -1)
                HandedOverPosition.SelectedIndex = HandedOverName.SelectedIndex;
        }

        private void AcceptedPosition_Loaded(object sender, RoutedEventArgs e)
        {
            var index = _viewModel.AcceptedPositions.IndexOf(AcceptedPosition.Text);
            if (index != -1)
                AcceptedName.SelectedIndex = AcceptedPosition.SelectedIndex;
        }

        private void AcceptedName_Loaded(object sender, RoutedEventArgs e)
        {
            var index = _viewModel.AcceptedNames.IndexOf(AcceptedName.Text);
            if (index != -1)
                AcceptedPosition.SelectedIndex = AcceptedName.SelectedIndex;
        }

        private void AdminPosition_Loaded(object sender, RoutedEventArgs e)
        {
            var index = _viewModel.AdminPositions.IndexOf(AdminPosition.Text);
            if (index != -1)
                AdminName.SelectedIndex = AdminPosition.SelectedIndex;
        }

        private void AdminName_Loaded(object sender, RoutedEventArgs e)
        {
            var index = _viewModel.AdminNames.IndexOf(AdminName.Text);
            if (index != -1)
                AdminPosition.SelectedIndex = AdminName.SelectedIndex;

        }

        private void AdminPosition_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var index = _viewModel.AdminPositions.IndexOf(AdminPosition.Text);
            if (index != -1)
                AdminName.SelectedIndex = AdminPosition.SelectedIndex;
        }

        private void AdminName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var index = _viewModel.AdminNames.IndexOf(AdminName.Text);
            if (index != -1)
                AdminPosition.SelectedIndex = AdminName.SelectedIndex;
        }

        private void AcceptedName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var index = _viewModel.AcceptedNames.IndexOf(AcceptedName.Text);
            if (index != -1)
                AcceptedPosition.SelectedIndex = AcceptedName.SelectedIndex;
        }

        private void AcceptedPosition_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var index = _viewModel.AcceptedPositions.IndexOf(AcceptedPosition.Text);
            if (index != -1)
                AcceptedName.SelectedIndex = AcceptedPosition.SelectedIndex;
        }

        private void HandedOverName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var index = _viewModel.HandedOverNames.IndexOf(HandedOverName.Text);
            if (index != -1)
                HandedOverPosition.SelectedIndex = HandedOverName.SelectedIndex;
        }

        private void HandedOverPosition_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var index = _viewModel.HandedOverPositions.IndexOf(HandedOverPosition.Text);
            if (index != -1)
                HandedOverName.SelectedIndex = HandedOverPosition.SelectedIndex;
        }

        private void GoodsGrid_CurrentCellChanged(object sender, EventArgs e)
        {
            if (GoodsGrid.CurrentColumn?.DisplayIndex == 7)
            {
                decimal goodsTotalQuantity = 0;
                decimal goodsTotalWeight = 0;
                decimal goodsTotalSum = 0;
                decimal containerTotalQuantity = 0;
                decimal containerTotalWeight = 0;
                decimal containerTotalSum = 0;
                string currentSection = "Товары";
                foreach (var item in _viewModel.GoodsItems)
                {


                    decimal sum = item.Price * item.Quantity;
                    if (item.Name == "Тара")
                    {
                        currentSection = "Тара";
                        sum = 0;
                    }

                    if (currentSection == "Товары")
                    {

                        if (item.Name != "Итого" && item.Name != "Всего по акту")
                        {
                            goodsTotalQuantity += item.Quantity;
                            goodsTotalWeight += item.Weight;
                            goodsTotalSum += sum;
                            item.Amount = item.Price * item.Quantity;
                        }

                        else if (item.Name == "Итого")
                        {
                            item.Quantity = goodsTotalQuantity;
                            item.Weight = goodsTotalWeight;
                            item.Price = 0;
                            item.Amount = goodsTotalSum;
                                            
                        }
                    }
                    else 
                    {
                        if (item.Name != "Итого" && item.Name != "Всего по акту")
                        {
                            containerTotalQuantity += item.Quantity;
                            containerTotalWeight += item.Weight;
                            containerTotalSum += sum;
                            item.Amount = item.Price * item.Quantity;
                        }
                        else if (item.Name == "Итого")
                        {
                            item.Quantity = containerTotalQuantity;
                            item.Weight = containerTotalWeight;
                            item.Price = 0;
                            item.Amount = containerTotalSum;
                        }
                    }

                    if (item.Name == "Всего по акту")
                    {
                        item.Quantity = 0;
                        item.Weight = 0;
                        item.Price = 0;
                        item.Amount = goodsTotalSum + containerTotalSum;

                    }

                }


            }
        }

        private void GoodsGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            GoodsGrid.ItemsSource = _viewModel.GoodsItems;
        }
    }
}
