using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDeserializer.Models;

namespace ExcelDeserializer.ViewModels {
    public class DocumentViewModel : INotifyPropertyChanged {

        public DocumentViewModel() {

        }

        public Row DocumentGrid { get; set; }

        public decimal Cell1 {
            get {
                return DocumentGrid.Cell1;
            }
            set {
                DocumentGrid.Cell1 = value;
                RaisePropertyChanged("Cell1");
            }
        }

        public decimal Cell2 {
            get {
                return DocumentGrid.Cell2;
            }
            set {
                DocumentGrid.Cell2 = value;
                RaisePropertyChanged("Cell2");
            }
        }

        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        private void RaisePropertyChanged(string propertyName) {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        #endregion
    }
}
