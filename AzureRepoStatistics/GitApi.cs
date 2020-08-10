using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace AzureRepoStatistics
{

    public class GitApi
    {
        private string _apiUrl, _owner, _repo, _notesRepo, _token;
        private HttpClient _client;
        DataTable _dt = new DataTable();
        List<string> _dataArray= new List<string>();
        private decimal _pageSize = 100;
        public GitApi()
        {
            _apiUrl = ConfigurationManager.AppSettings["api_url"];
            _owner = ConfigurationManager.AppSettings["api_owner"];
            _repo = ConfigurationManager.AppSettings["repo_name"];
            _notesRepo = ConfigurationManager.AppSettings["notes_repo_name"];
            _token = ConfigurationManager.AppSettings["api_token"];

        }

        private void SetupDataTable()
        {
            Console.WriteLine("Setup Data Table");
            
            _dt.Columns.Add("StartDate", typeof(DateTime));
            _dt.Columns.Add("EndDate", typeof(DateTime));
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

            DateTime date = startDate;
            _dataArray = GetTxtFileLines();
            SearchRepo(startDate, startDate, endDate);
            SearchNotesRepo(startDate,startDate,endDate);


            //while (date <= endDate)
            //{
            //    SearchRepo(date,startDate,endDate);
            //    SearchNotesRepo(date,startDate,endDate);
            //    date = date.AddDays(1);

            //}

            //SearchRepo(startDate);
            //SearchRepo(endDate);

            //SearchNotesRepo(startDate);
            //SearchNotesRepo(endDate);

            return _dt; 
        }
        public void SearchRepo(DateTime date,DateTime startDate, DateTime endDate)
        {
            Console.WriteLine(string.Format("Fetching contributions from {1} Repo", date.ToShortDateString(), _repo));

            string query = string.Format("search/issues?q=repo:{0}/{1}+is:pr+is:merged+sort:author-date-asc+merged:>={2}&sort=merged&per_page={3}", _owner, _repo, date.ToString("yyyy-MM-dd"), _pageSize);
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
                
                decimal pageCount = Math.Ceiling(Convert.ToDecimal(response.total_count) / _pageSize); 
                for (int pageNo = 1; pageNo <= pageCount; pageNo++)
                {
                    query = string.Format("search/issues?q=repo:{0}/{1}+is:pr+is:merged+sort:author-date-asc+merged:>={2}&sort=merged&per_page={3}&page={4}", _owner, _repo, date.ToString("yyyy-MM-dd"),_pageSize, pageNo);
                    var apiTasks = _client.GetAsync(query);
                    apiTasks.Wait();
                    var apiResult = apiTasks.Result;
                    if (apiTasks.IsCompleted)
                    {
                        var apiReadTask = apiResult.Content.ReadAsStringAsync();
                        SearchResult apiResponse = JsonConvert.DeserializeObject<SearchResult>(apiReadTask.Result);

                        foreach (var item in apiResponse.items)
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
                                        dr["StartDate"] = startDate.ToShortDateString(); //only date 
                                        dr["EndDate"] = endDate.ToShortDateString(); //only date 
                                        dr["GitUser"] = item.user.login;
                                        dr["Status"] = file.status;
                                        dr[folder] = 1;
                                        dr["TotalContribution"] = 1;
                                        
                                        string email = (user.email == null ? "" : user.email.ToLower());
                                        string company = (user.company == null ? "" : user.company.ToLower());
                                        //Check if External or MSFT user
                                        // if user.name is in _dataArray then account = MSFT else External
                                        if (_dataArray.Contains(item.user.login))
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

                }
            }
        }
        public void SearchNotesRepo(DateTime date,DateTime startDate,DateTime endDate)
        {
            Console.WriteLine(string.Format("Fetching contributions from {1} Repo", date.ToShortDateString(), _notesRepo));
            string query = string.Format("search/issues?q=repo:{0}/{1}+is:pr+is:merged+sort:author-date-asc+merged:>={2}&sort=merged&page_size={3}", _owner, _notesRepo, date.ToString("yyyy-MM-dd"),_pageSize);
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
                //int pageSize = 100;
                decimal pageCount = Math.Ceiling(Convert.ToDecimal(response.total_count) / _pageSize); 
                for (int pageNo = 1; pageNo <= pageCount; pageNo++)
                {
                    query = string.Format("search/issues?q=repo:{0}/{1}+is:pr+is:merged+sort:author-date-asc+merged:>={2}&sort=merged&per_page={3}&page={4}", _owner, _repo, date.ToString("yyyy-MM-dd"), _pageSize, pageNo);
                    var apiTasks = _client.GetAsync(query);
                    apiTasks.Wait();
                    result = apiTasks.Result;
                    if (apiTasks.IsCompleted)
                    {
                        readTask = result.Content.ReadAsStringAsync();
                        response = JsonConvert.DeserializeObject<SearchResult>(readTask.Result);
                        foreach (var item in response.items)
                        {
                            //Add count on all folders which has been changed.
                            List<PullFile> files = GetPullFiles(item.pull_request.url);
                            int total = 0;
                            User user = GetUser(item.user.url);
                            foreach (var file in files)
                            {
                                //string folder = GetParentFolder(file.filename);
                                if (!file.filename.ToLower().Contains(".png") && !file.filename.ToLower().Contains(".svg"))
                                {
                                    if (file.status == "added" || file.status == "modified")
                                    {
                                        DataRow dr = _dt.NewRow();
                                        dr["StartDate"] = startDate.ToShortDateString(); //only date 
                                        dr["EndDate"] = endDate.ToShortDateString(); //only date 
                                        dr["GitUser"] = item.user.login;
                                        dr["Status"] = file.status;
                                        dr["Notebooks @ efbace2"] = 1;
                                        dr["TotalContribution"] = 1;

                                        string email = (user.email == null ? "" : user.email.ToLower());
                                        string company = (user.company == null ? "" : user.company.ToLower());
                                        //Check if External or MSFT user
                                       
                                        // if user.name is in list then account = MSFT else External
                                        if (_dataArray.Contains(item.user.login))
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

                }
                    
            }

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
        public static List<string> GetTxtFileLines()
        {
            string directory = System.IO.Directory.GetCurrentDirectory();
            for (int counter_slash = 0; counter_slash < 2; counter_slash++)
            {
                directory = directory.Substring(0, directory.LastIndexOf(@"\"));
            }
            string filePath = string.Concat(directory, "\\names.txt");

            List<string> result = new List<string>(); // A list of strings 
            // Create a stream reader object to read txt file.
            using (StreamReader reader = new StreamReader(filePath))
            {
                string line = string.Empty; // Contains a single line returned by the stream reader object.
                // While there are lines in the file, read a line into the line variable.
                while ((line = reader.ReadLine()) != null)
                {
                    // If the line is not empty, add it to the list.
                    if (line != string.Empty)
                    {
                        result.Add(line);
                    }
                }
            }
            return result;
        }
    }


}
