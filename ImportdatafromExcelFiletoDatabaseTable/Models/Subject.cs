using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ImportdatafromExcelFiletoDatabaseTable.Models
{
    public class Subject
    {
        public string RefId { get; set; }

        public string Name { get; set; }
        public string Code { get; set; }
        public string Description { get; set; }
    }
}
