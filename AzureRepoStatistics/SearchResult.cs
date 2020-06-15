using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AzureRepoStatistics
{
    class SearchResult
    {
        public int total_count { get; set; }
        public bool incomplete_results { get; set; }
        public List<Items> items { get; set; }
    }
}
