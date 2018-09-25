using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;

namespace Inheritance
{
    class Program
    {
        static void Main(string[] args)
        {          
            AppConfig app = new AppConfig();
            Reports_base un = new Reports_base(app.GetConnectionPrefix());
            // to run the report
            un.RunReports();                     
        }
    }
}

