using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;

namespace ExcelDeserializer.Converters
{
    [ValueConversion(typeof(DateTime), typeof(String))]
    class ObjectToDecimalConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            try
            {
                return decimal.Parse(value.ToString());
            }
            catch
            {

                throw new ApplicationException("Enter numbers only!");
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }
    }
}
