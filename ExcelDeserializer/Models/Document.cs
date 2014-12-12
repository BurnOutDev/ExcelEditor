using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDeserializer.Models {
    public class Row {
        public Row(decimal cell1Value, decimal cell2Value) {
            Cell1 = cell1Value;
            Cell2 = cell1Value;
        }

        public Row() {

        }

        public decimal Cell1 { get; set; }
        public decimal Cell2 { get; set; }
    }

    public class XLRowCustom {
        public XLRowCustom(IXLRow row)
        {
            Cell1 = row.Cell(1);
            Cell2 = row.Cell(2);
        }
        public IXLCell Cell1 { get; set; }
        public IXLCell Cell2 { get; set; }
    }
}
