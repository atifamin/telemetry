using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using System.IO;

namespace AzureRepoStatistics
{
    class Program
    {
        private static Microsoft.Office.Interop.Excel.Workbook mWorkBook;
        private static Microsoft.Office.Interop.Excel.Sheets mWorkSheets;
        private static Microsoft.Office.Interop.Excel.Worksheet mWSheet1;
        private static Microsoft.Office.Interop.Excel.Application oXL;
        static void Main(string[] args)
        {
            ConnectRepo();
            UploadToAzureBlob();
        }
        static void ConnectRepo()
        {


            DateTime endDate = new DateTime(2020, 6, 8);//DateTime.Now;
            DateTime startDate = endDate.AddDays(-1);

            GitApi api = new GitApi();

            DataTable data = api.ProcessRepo(startDate,endDate);



            //added files datatable 
            DataTable dataFilter = data.Select("Status ='added'").CopyToDataTable();
            DataTable dataAdded = dataFilter.AsEnumerable()
                          .GroupBy(g => new { ActivityDate = g["ActivityDate"], GitUser = g["GitUser"], AccountType = g["AccountType"] })
                          .OrderBy(o => o.Key.ActivityDate)
                          .Select(s =>
                          {
                              var row = data.NewRow();

                              row["ActivityDate"] = s.Key.ActivityDate;
                              row["GitUser"] = s.Key.GitUser;
                              row["AccountType"] = s.Key.AccountType;
                              row["TotalContribution"] = s.Sum(r => r.Field<int>("TotalContribution"));
                              row["DataConnectors"] = s.Sum(r => r.Field<int>("DataConnectors"));
                              row["Workbooks"] = s.Sum(r => r.Field<int>("Workbooks"));
                              row["Playbooks"] = s.Sum(r => r.Field<int>("Playbooks"));
                              row["Exploration Queries"] = s.Sum(r => r.Field<int>("Exploration Queries"));
                              row["Hunting Queries"] = s.Sum(r => r.Field<int>("Hunting Queries"));
                              row["Sample Data"] = s.Sum(r => r.Field<int>("Sample Data"));
                              row["Tools"] = s.Sum(r => r.Field<int>("Tools"));
                              row["Detections"] = s.Sum(r => r.Field<int>("Detections"));
                              row["Notebooks @ efbace2"] = s.Sum(r => r.Field<int>("Notebooks @ efbace2"));
                              return row;
                          })
                          .CopyToDataTable();

            //modified files datatable 
            dataFilter = data.Select("Status ='modified'").CopyToDataTable();
            DataTable dataModified = dataFilter.AsEnumerable()
                          .GroupBy(g => new { ActivityDate = g["ActivityDate"], GitUser = g["GitUser"], AccountType = g["AccountType"] })
                          .OrderBy(o => o.Key.ActivityDate)
                          .Select(s =>
                          {
                              var row = data.NewRow();

                              row["ActivityDate"] = s.Key.ActivityDate;
                              row["GitUser"] = s.Key.GitUser;
                              row["AccountType"] = s.Key.AccountType;
                              row["TotalContribution"] = s.Sum(r => r.Field<int>("TotalContribution"));
                              row["DataConnectors"] = s.Sum(r => r.Field<int>("DataConnectors"));
                              row["Workbooks"] = s.Sum(r => r.Field<int>("Workbooks"));
                              row["Playbooks"] = s.Sum(r => r.Field<int>("Playbooks"));
                              row["Exploration Queries"] = s.Sum(r => r.Field<int>("Exploration Queries"));
                              row["Hunting Queries"] = s.Sum(r => r.Field<int>("Hunting Queries"));
                              row["Sample Data"] = s.Sum(r => r.Field<int>("Sample Data"));
                              row["Tools"] = s.Sum(r => r.Field<int>("Tools"));
                              row["Detections"] = s.Sum(r => r.Field<int>("Detections"));
                              row["Notebooks @ efbace2"] = s.Sum(r => r.Field<int>("Notebooks @ efbace2"));

                              return row;
                          })
                          .CopyToDataTable();

            ExportToExcel(dataAdded, dataModified);


        }

        static void ExportToExcel(DataTable dataAdded, DataTable dataModified)
        {
            Console.WriteLine("Writing to Excel file");
            string fileName = ConfigurationManager.AppSettings["template_filename"];
            string path = string.Concat(System.IO.Directory.GetCurrentDirectory(), fileName);
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = false;
            oXL.DisplayAlerts = false;
            mWorkBook = oXL.Workbooks.Open(path, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Get all the sheets in the workbook
            mWorkSheets = mWorkBook.Worksheets;

            //Write to New Contribution
            mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("GitHub New Contribution");
            Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;
            int colCount = range.Columns.Count;
            int rowCount = range.Rows.Count;
            int index = 1;
            foreach (DataRow item in dataAdded.Rows)
            {
                mWSheet1.Cells[rowCount+ index, 1] = item["ActivityDate"];
                mWSheet1.Cells[rowCount + index, 2] = item["GitUser"];
                mWSheet1.Cells[rowCount + index, 3] = item["AccountType"];
                mWSheet1.Cells[rowCount + index, 4] = item["TotalContribution"];
                mWSheet1.Cells[rowCount + index, 5] = item["DataConnectors"];
                mWSheet1.Cells[rowCount + index, 6] = item["Workbooks"];
                mWSheet1.Cells[rowCount + index, 7] = item["Playbooks"];
                mWSheet1.Cells[rowCount + index, 8] = item["Exploration Queries"];
                mWSheet1.Cells[rowCount + index, 9] = item["Hunting Queries"];
                mWSheet1.Cells[rowCount + index, 10] = item["Sample Data"];
                mWSheet1.Cells[rowCount + index, 11] = item["Tools"];
                mWSheet1.Cells[rowCount + index, 12] = item["Detections"];
                mWSheet1.Cells[rowCount + index, 13] = item["Notebooks @ efbace2"];
                index++;
            }

            //Write to Update Contribution
            mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("GitHub Update Contribution");
            index = 1;
            foreach (DataRow item in dataModified.Rows)
            {
                mWSheet1.Cells[rowCount + index, 1] = item["ActivityDate"];
                mWSheet1.Cells[rowCount + index, 2] = item["GitUser"];
                mWSheet1.Cells[rowCount + index, 3] = item["AccountType"];
                mWSheet1.Cells[rowCount + index, 4] = item["TotalContribution"];
                mWSheet1.Cells[rowCount + index, 5] = item["DataConnectors"];
                mWSheet1.Cells[rowCount + index, 6] = item["Workbooks"];
                mWSheet1.Cells[rowCount + index, 7] = item["Playbooks"];
                mWSheet1.Cells[rowCount + index, 8] = item["Exploration Queries"];
                mWSheet1.Cells[rowCount + index, 9] = item["Hunting Queries"];
                mWSheet1.Cells[rowCount + index, 10] = item["Sample Data"];
                mWSheet1.Cells[rowCount + index, 11] = item["Tools"];
                mWSheet1.Cells[rowCount + index, 12] = item["Detections"];
                mWSheet1.Cells[rowCount + index, 13] = item["Notebooks @ efbace2"];
                index++;
            }
            mWorkBook.Save();
            mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
            mWSheet1 = null;
            mWorkBook = null;
            oXL.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            Console.WriteLine("Excel file has been save at location: " + Environment.NewLine + path);
            
        }

        static void UploadToAzureBlob()
        {
            Console.Write("Uploading file to Azure cloud storage");

            string storageConnection = ConfigurationManager.AppSettings["blobstorage_connectionstring"];
            string storageContainer = ConfigurationManager.AppSettings["blobstorage_container"];
            string fileName = ConfigurationManager.AppSettings["template_filename"];
            string filePath = string.Concat(System.IO.Directory.GetCurrentDirectory(), fileName);

            CloudStorageAccount cloudStorageAccount = CloudStorageAccount.Parse(storageConnection);
            //create a block blob 
            CloudBlobClient cloudBlobClient = cloudStorageAccount.CreateCloudBlobClient();

            //create a container 
            CloudBlobContainer cloudBlobContainer = cloudBlobClient.GetContainerReference(storageContainer);

            //create a container if it is not already exists
            if (cloudBlobContainer.CreateIfNotExists())
            {
                cloudBlobContainer.SetPermissionsAsync(new BlobContainerPermissions { PublicAccess = BlobContainerPublicAccessType.Blob });
            }


            var imageToUpload = System.IO.File.OpenRead(filePath);


            //get Blob reference
            CloudBlockBlob cloudBlockBlob = cloudBlobContainer.GetBlockBlobReference(storageContainer);
            var ext = Path.GetExtension(imageToUpload.Name).Split('.');
            cloudBlockBlob.Properties.ContentType = ext[1];

            // Upload using the UploadFromStream method.
            using (var stream = System.IO.File.OpenRead(filePath))
                cloudBlockBlob.UploadFromStream(stream, stream.Length, null, null, null);

            Console.WriteLine("");
            Console.Write("Uploading file to Azure cloud storage has been completed");
            
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("Process has been completed, press any key to exit.");
            Console.ReadKey();
        }

    }

}
