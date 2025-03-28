using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Var18.Classes.ModelData
{
    public class DocumentData : INotifyPropertyChanged
    {
        private decimal _goodsTotal;

        public string OKUDCode { get; set; } = "0330518";
        public string OKPOCode { get; set; }
        public string OKDPCode { get; set; }
        public string OrganizationName { get; set; }
        public string Department { get; set; }
        public string DocumentNumber { get; set; }
        public DateTime DocumentDate { get; set; } = DateTime.Now;
        public ObservableCollection<GoodsItem> GoodsItems { get; set; } = new ObservableCollection<GoodsItem>();

        public decimal GoodsTotal
        {
            get => _goodsTotal;
            set
            {
                if (_goodsTotal != value)
                {
                    _goodsTotal = value;
                    OnPropertyChanged();
                }
            }
        }

        public decimal TotalAmount => GoodsTotal;

        public string HandedOverPosition { get; set; }
        public string HandedOverName { get; set; }
        public string AcceptedPosition { get; set; }
        public string AcceptedName { get; set; }
        public string AdminPosition { get; set; }
        public string AdminName { get; set; }

        public void UpdateTotals()
        {
            GoodsTotal = GoodsItems.Sum(item => item.Amount);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
