﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Var18.Classes.ModelData
{
    public class UnitInfo
    {
        public string Code { get; set; }  // Код ОКЕИ
        public string Name { get; set; }  // Наименование единицы
        public string ShortName { get; set; }  // Краткое наименование
    }

    public class GoodsItem : INotifyPropertyChanged
    {
        private int _number;
        private string _name;
        private int _code;
        private string _unit;
        private int _okeiCode;
        private decimal _weight;
        private decimal _quantity;
        private decimal _price;

        public int Number
        {
            get => _number;
            set
            {
                if (_number != value)
                {
                    _number = value;
                    OnPropertyChanged();
                }
            }
        }

        public string Name
        {
            get => _name;
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChanged();
                }
            }
        }

        public int Code
        {
            get => _code;
            set
            {
                if (_code != value)
                {
                    _code = value;
                    OnPropertyChanged();
                }
            }
        }

        public string Unit
        {
            get => _unit;
            set
            {
                _unit = value;
                OnPropertyChanged();
            }
        }

        public int OKEICode
        {
            get => _okeiCode;
            set
            {
                if (_okeiCode != value)
                {
                    _okeiCode = value;
                    OnPropertyChanged();
                }
            }
        }

        public decimal Weight
        {
            get => _weight;
            set
            {
                if (_weight != value)
                {
                    _weight = value;
                    OnPropertyChanged();
                }
            }
        }

        public decimal Quantity
        {
            get => _quantity;
            set
            {
                if (_quantity != value)
                {
                    _quantity = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(Amount));
                }
            }
        }

        public decimal Price
        {
            get => _price;
            set
            {
                if (_price != value)
                {
                    _price = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(Amount));
                }
            }
        }
        //public decimal Amount => Quantity * Price;

        // Изменяемое свойство Amount
        private decimal _amount;
        public decimal Amount
        {
            get => _amount;
            set
            {
                _amount = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private UnitInfo _selectedUnit;

        public UnitInfo SelectedUnit
        {
            get => _selectedUnit;
            set
            {
                _selectedUnit = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(UnitCode)); // Обновляем код при изменении единицы
                OnPropertyChanged(nameof(UnitName)); // Обновляем название
            }
        }

        // Код ОКЕИ (только для чтения)
        public string UnitCode => SelectedUnit?.Code;

        // Название единицы (только для чтения)
        public string UnitName => SelectedUnit?.ShortName;


        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public GoodsItem Clone()
        {
            return new GoodsItem
            {
                Number = this.Number,
                Name = this.Name,
                Code = this.Code,
                Unit = this.Unit,
                OKEICode = this.OKEICode,
                Weight = this.Weight,
                Quantity = this.Quantity,
                Price = this.Price,
                Amount = this.Amount
            };
        }

        public bool Validate()
        {
            if (string.IsNullOrWhiteSpace(Name))
            {
                return false;
            }

            if (Quantity <= 0 || Price < 0)
            {
                return false;
            }

            return true;
        }
    }
}
