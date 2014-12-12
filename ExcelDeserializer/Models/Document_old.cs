using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDeserializer.Models {
    class Document {

        #region Constructors
        public Document(params Column[] columns) {
            Columns = columns;
        }

        public Document() {
            Columns = new List<Column>();
        }
        #endregion

        #region Properties
        public ICollection<Column> Columns { get; set; }
        public ICollection<Row> Rows { get; set; }
        #endregion
    }
}
