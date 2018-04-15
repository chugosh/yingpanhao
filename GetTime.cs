using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication1
{
    class GetTime
    {
    	string time;
    	public void setTime(string time){
    		this.time = time;
    	}
        public string getDateToday() {
            return DateTime.Now.ToString("yyyy-MM-dd");
        }
        public string getDateYestoday() {
            return DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
        }
    }
}
