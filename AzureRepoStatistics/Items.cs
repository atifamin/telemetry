using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AzureRepoStatistics
{
    public class Items
    {
        public string url { get; set; }
        public int id { get; set; }
        public string node_id { get; set; }
        public User user { get; set; }
        public PullRequest pull_request { get; set; }
        public DateTime created_at { get; set; }
    }
}
