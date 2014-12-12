using System;
using System.Collections.Generic;
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
using Microsoft.Win32;
using System.Collections.ObjectModel;
using ExcelDeserializer.Models;
using ClosedXML.Excel;
using System.IO;

namespace ExcelDeserializer {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {

        public ObservableCollection<XLRowCustom> Worksheet { get; set; }
        public XLWorkbook Workbook { get; set; }
        public IXLWorksheet SelectedWorksheet {
            get {
                //return Workbook.Worksheets.ElementAt(cbxWorksheets.SelectedIndex);
                return Workbook.Worksheets.FirstOrDefault();
            }
        }
        public OpenFileDialog OpenDialog { get; set; }

        public string FilePath {
            get {
                return OpenDialog.FileName;
            }
        }

        public MainWindow() {
            Worksheet = new ObservableCollection<XLRowCustom>();
            InitializeComponent();
            DataContext = this;
        }

        private void btnOpen_Click(object sender, RoutedEventArgs e) {
            #region OpenDialog initialization
            OpenDialog = new OpenFileDialog();
            OpenDialog.Title = "Choose File";
            OpenDialog.Filter = "Excel 2013|*.xlsx|Text file|*.txt";
            OpenDialog.Multiselect = false;
            #endregion

            if (OpenDialog.ShowDialog().Value) {
                tbxFilePath.Text = OpenDialog.SafeFileName;
            }

            GetDataFromExcel();
            BindData();
        }

        #region Get data from Excel and bind
        private XLWorkbook GetDataFromExcel() {
            return Workbook = new XLWorkbook(FilePath);
        }

        private void BindData() {
            if (Workbook != null) {
                Worksheet.Clear();
                foreach (var row in SelectedWorksheet.Rows()) {
                    Worksheet.Add(new XLRowCustom(row));
                }

                var t = SelectedWorksheet.Rows().Select(r => r.Cell(1));
                var t2 = SelectedWorksheet.Rows().Select(r => r.Cell(2));

                cbxWorksheets.ItemsSource = Workbook.Worksheets; 
            }
        }
        #endregion

        private void cbxWorksheets_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            BindData();
        }

        private void btnRemoveLastRow_Click(object sender, RoutedEventArgs e) {
            Worksheet.Remove(Worksheet.Last());
        }

        private void btnAddRow_Click(object sender, RoutedEventArgs e) {
            SelectedWorksheet.Rows().Last().InsertRowsBelow(1);
            SelectedWorksheet.Rows().Last().Cell(1).Value = 0;
            SelectedWorksheet.Rows().Last().Cell(2).Value = 0;
            BindData();
        }

        private void btnSave_Click(object sender, RoutedEventArgs e) {
            #region SaveDialog initialization
            SaveDialog = new SaveFileDialog();
            SaveDialog.Title = "Choose...";
            SaveDialog.Filter = "Excel 2013|*.xlsx|Text file|*.txt";
            #endregion

            if (SaveDialog.ShowDialog().Value) {
                Workbook.SaveAs(SaveDialog.FileName); 
            }
        }

        public SaveFileDialog SaveDialog { get; set; }
    }
}
