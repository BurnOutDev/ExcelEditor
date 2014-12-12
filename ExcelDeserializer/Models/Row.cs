using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDeserializer.Models {
    class Row {

        #region Constructors
        public Row(Column column) {
            
        }
        #endregion

        public Cell Cell1 { get; set; }
        public Cell Cell2 { get; set; }
    }
}
