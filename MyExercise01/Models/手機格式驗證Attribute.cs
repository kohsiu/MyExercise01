using System;
using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;

namespace MyExercise01.Models
{
    internal class 手機格式驗證Attribute : DataTypeAttribute
    {
        public 手機格式驗證Attribute() : base(DataType.Text)
        {

        }

        public override bool IsValid(object value)
        {
            if (value != null)
            {
                string pattern = @"(\d{4}(-)\d{6})$";
                var isMatch = Regex.IsMatch(value.ToString(), pattern);
                return isMatch;
            }

            return true;
        }
    }
}