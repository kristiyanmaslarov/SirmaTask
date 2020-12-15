using System;

namespace Infrastructure
{
    public class EmployeeBase
    {
        public int ID { get; set; }

        public int ProjectID { get; set; }

        public DateTime DateFrom { get; set; }

        public DateTime DateTo { get; set; }
    }
}
