using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeWhitener
{
    class MainWindowVM : ViewModelBase
    {
        public MainWindowVM()
        {
            Items = new ObservableCollection<OfficeFileItemVM>();
        }

        public double OverlayOpacity { get { return _OverlayOpacity; } set { _OverlayOpacity = value; RaisePropertyChanged("OverlayOpacity"); } } private double _OverlayOpacity = 0.6;
        
        private ObservableCollection<OfficeFileItemVM> _Items;

        public ObservableCollection<OfficeFileItemVM> Items
        {
            get { return _Items; }
            set { _Items = value; _Items.CollectionChanged += _Items_CollectionChanged; }
        }

        void _Items_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            OverlayOpacity = (Items.Count > 0) ? 0.3 : 0.6;
        }

    }
}
