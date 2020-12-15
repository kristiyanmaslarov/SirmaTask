using System;
using System.Reflection;
using System.Text;

namespace Infrastructure
{
    public class Employee : EmployeeBase
    {
        public Employee()
        {
            //2) DateTo приема и стойност „NULL“ (което е еквивалент на „днес“)
            DateTo = DateTime.Now.Date;
        }

        private PropertyInfo[] _PropertyInfos = null;

        public override string ToString()
        {
            if (_PropertyInfos == null)
                _PropertyInfos = this.GetType().GetProperties();

            var sb = new StringBuilder();

            foreach (var info in _PropertyInfos)
            {
                var value = info.GetValue(this, null) ?? "(null)";
                sb.Append(info.Name + ": " + value + " ");
            }

            return sb.ToString();
        }
    }
}
