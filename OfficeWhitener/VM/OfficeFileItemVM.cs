using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Media;

namespace OfficeWhitener
{
    class OfficeFileItemVM : ViewModelBase
    {
        public OfficeFileItemVM()
        {

        }

        public OfficeFileItemVM(string filePath)
        {
            FilePath = filePath;
            Filename = Path.GetFileName(filePath);
        }

        public string FilePath { get { return _FilePath; } set { _FilePath = value; RaisePropertyChanged("FilePath"); } } private string _FilePath;
        public string Filename { get { return _Filename; } set { _Filename = value; RaisePropertyChanged("Filename"); } } private string _Filename;
        public string ErrorMessage { get { return _ErrorMessage; } set { _ErrorMessage = value; RaisePropertyChanged("ErrorMessage"); } } private string _ErrorMessage;
        public string State { get { return _State; } set { _State = value; RaisePropertyChanged("State"); RaisePropertyChanged("StateSymbol"); RaisePropertyChanged("StateBrush"); } } private string _State;
        public string StateSymbol
        {
            get
            {
                switch (State)
                {
                    case "OK":
                        return "\uE10B";
                    case "NG":
                        return "\uE10A";
                    default:
                        return "\uE11B";
                }
            }
        }
        public SolidColorBrush StateBrush
        {
            get
            {
                switch (State)
                {
                    case "OK":
                        return Brushes.LimeGreen;
                    case "NG":
                        return Brushes.Red;
                    default:
                        return Brushes.Silver;
                }
            }
        }

    }
}
