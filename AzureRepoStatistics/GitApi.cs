using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace AzureRepoStatistics
{

    public class GitApi
    {
        private string _apiUrl;
        private string _owner;
        private string _repo;
        private string _token;
        private HttpClient _client;
        DataTable _dt = new DataTable();

        public GitApi()
        {
            _apiUrl = ConfigurationManager.AppSettings["api_url"];
            _owner = ConfigurationManager.AppSettings["api_owner"];
            _repo = ConfigurationManager.AppSettings["repo_name"];
            _token = ConfigurationManager.AppSettings["api_token"];

        }

        private void SetupDataTable()
        {
            Console.WriteLine("Setup Data Table");
            
            _dt.Columns.Add("ActivityDate", typeof(DateTime));
            _dt.Columns.Add("GitUser", typeof(string));
            _dt.Columns.Add("AccountType", typeof(string));
            _dt.Columns.Add("Status", typeof(string));
            _dt.Columns.Add("TotalContribution", typeof(Int32));

            //Add foldername as columns
            _dt.Columns.Add("DataConnectors", typeof(Int32));
            _dt.Columns.Add("Workbooks", typeof(Int32));
            _dt.Columns.Add("Playbooks", typeof(Int32));
            _dt.Columns.Add("Exploration Queries", typeof(Int32));
            _dt.Columns.Add("Hunting Queries", typeof(Int32));
            _dt.Columns.Add("Sample Data", typeof(Int32));
            _dt.Columns.Add("Tools", typeof(Int32));
            _dt.Columns.Add("Detections", typeof(Int32));
            _dt.Columns.Add("Notebooks @ efbace2", typeof(Int32));

            _dt.Columns["TotalContribution"].DefaultValue = 0;
            _dt.Columns["DataConnectors"].DefaultValue = 0;
            _dt.Columns["Workbooks"].DefaultValue = 0;
            _dt.Columns["Playbooks"].DefaultValue = 0;
            _dt.Columns["Exploration Queries"].DefaultValue = 0;
            _dt.Columns["Hunting Queries"].DefaultValue = 0;            
            _dt.Columns["Sample Data"].DefaultValue = 0;
            _dt.Columns["Tools"].DefaultValue = 0;
            _dt.Columns["Detections"].DefaultValue = 0;
            _dt.Columns["Notebooks @ efbace2"].DefaultValue = 0;

        }
        public DataTable ProcessRepo(DateTime startDate, DateTime endDate)
        {
            SetupDataTable();
            SearchRepo(startDate);
            SearchRepo(endDate);
            return _dt; 
        }
        public DataTable SearchRepo(DateTime date)
        {
            Console.WriteLine(string.Format("Fetching {0} from {1} Repo", date.ToShortDateString(), _repo));

            string query = string.Format("search/issues?q=repo:{0}/{1}+is:pr+is:merged+sort:author-date-asc+merged:%3E={2}&sort=merged", _owner, _repo, date.ToString("yyyy-MM-dd"));
            _client = new HttpClient();
            _client.BaseAddress = new Uri(_apiUrl);
            _client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue(_repo, "1.0"));
            _client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            _client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Token", _token);
            var getTasks = _client.GetAsync(query);
            getTasks.Wait();
            var result = getTasks.Result;
            if (getTasks.IsCompleted)
            {
                var readTask = result.Content.ReadAsStringAsync();
                SearchResult response = JsonConvert.DeserializeObject<SearchResult>(readTask.Result);
                foreach (var item in response.items)
                {


                    //Add count on all folders which has been changed.
                    List<PullFile> files = GetPullFiles(item.pull_request.url);
                    int total = 0;
                    User user = GetUser(item.user.url);
                    foreach (var file in files)
                    {
                        string folder = GetParentFolder(file.filename);
                        if (_dt.Columns.Contains(folder))
                        {
                            if (file.status == "added" || file.status == "modified")
                            {
                                DataRow dr = _dt.NewRow();
                                dr["ActivityDate"] = item.closed_at.ToShortDateString(); //only date 
                                dr["GitUser"] = item.user.login;
                                dr["Status"] = file.status;
                                dr[folder] = 1;
                                dr["TotalContribution"] = 1;
                                //Get User info

                                string email = (user.email == null ? "" : user.email.ToLower());
                                string company = (user.company == null ? "" : user.company.ToLower());
                                //Check if External or MSFT user
                                if (email.Contains("microsoft") || company.Contains("microsoft"))
                                    dr["AccountType"] = "MSFT";
                                else
                                    dr["AccountType"] = "External";

                                total++;
                                _dt.Rows.Add(dr);
                            }
                        }
                    }

                }
            }

            return _dt;
        }


        List<PullFile> GetPullFiles(string url)
        {
            _client = new HttpClient();
            _client.BaseAddress = new Uri(_apiUrl);
            _client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue(_repo, "1.0"));
            _client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            _client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Token", _token);

            string pullUrl = string.Concat(url, "/files");
            var pullRquest = _client.GetAsync(pullUrl);
            pullRquest.Wait();
            var result = pullRquest.Result;
            if (pullRquest.IsCompleted)
            {
                var readTask = result.Content.ReadAsStringAsync();

                List<PullFile> response = JsonConvert.DeserializeObject<List<PullFile>>(readTask.Result);

                return response;
            }
            else
                return null;
        }

        string GetParentFolder(string directory)
        {
            //Don't Count if it is PNG or SVG file.
            if (directory.ToLower().Contains(".png") || directory.ToLower().Contains(".svg"))
                return "";

            string[] directories = directory.Split('/');
            if (directories.Count() > 0)
                return directories[0];
            else
                return directory;
            

        }

        User GetUser(string url)
        {
            _client = new HttpClient();
            _client.BaseAddress = new Uri(_apiUrl);
            _client.DefaultRequestHeaders.Add("User-Agent", "git-hub");
            var pullRquest = _client.GetAsync(url);
            pullRquest.Wait();
            var result = pullRquest.Result;
            if (pullRquest.IsCompleted)
            {
                var readTask = result.Content.ReadAsStringAsync();

                User response = JsonConvert.DeserializeObject<User>(readTask.Result);

                return response;
            }
            else
                return null;
        }
    }


}
