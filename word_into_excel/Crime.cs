using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace word_into_excel
{
    public class Crime
    {
        public string _date;
        public string _title;
        public string _department;
        public string _kusp;
        public string _description;
        public string _infaboutinitiation;
        public Crime(string date, string title, string department, string kusp, string description, string infaboutinitiation)
        {
            _date = date;
            _title = title;
            _department = department;
            _kusp = kusp;
            _description = description;
            _infaboutinitiation = infaboutinitiation;
        }
    }
}
