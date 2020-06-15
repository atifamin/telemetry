using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AzureRepoStatistics
{
    public class User
    {
        public string login { get; set; }
        public int id { get; set; }
        public string name { get; set; }
        public string url { get; set; }
        public string email { get; set; }
        public string company { get; set; }
    }
}
