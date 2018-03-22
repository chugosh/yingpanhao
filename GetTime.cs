using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication1
{
    class GetTime
    {
        public string getDateToday() {
            return DateTime.Now.ToString("yyyy-MM-dd");
        }
        public string getDateYestoday() {
            return DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
        }
    }
}
