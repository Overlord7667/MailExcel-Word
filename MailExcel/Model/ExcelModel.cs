using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailExcel.Model
{
    class ExcelModel
    {
        public string ListName {get;set;}
        public string TableName { get; set; }
        public int CellsCount { get; set; }
        public int RandomMin { get; set; }
        public int RandomMax { get; set; }
        public ExcelModel()
        {
            ListName = "Лист1";
            TableName = "Потребление газировки по месецам";
            CellsCount = 12;
            RandomMin = 2;
            RandomMax = 50;
        }
    }
}
