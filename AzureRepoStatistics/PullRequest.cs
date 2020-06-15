using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AzureRepoStatistics
{
    public class PullRequest
    {
        public int id { get; set; }
        public string url { get; set; }
        public string diff_url { get; set; }
    }
}
