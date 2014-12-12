using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDeserializer.Models {
    class Column {
        public ICollection<Row> Rows { get; set; }
    }
}
