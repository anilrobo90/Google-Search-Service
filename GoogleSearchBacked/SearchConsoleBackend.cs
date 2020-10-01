using ClosedXML.Excel;
using Google.Apis.Customsearch.v1;
using Google.Apis.Customsearch.v1.Data;
using Google.Apis.Services;
using Sentry;
using System;
using System.Collections.Generic;
using System.ServiceProcess;
using System.Threading;
using System.Windows.Threading;

namespace GoogleSearchBacked
{
    public partial class SearchConsoleBackend : ServiceBase
    {
        private Timer timer_thread = null;
        private String searchquery = String.Empty;
        private int noOfSearches = 0;
        private int timeout = 0;


        public SearchConsoleBackend()
        {
            InitializeComponent();
            SentrySdk.Init("https://361f3203cce243259bd6750d7becc03b@o335353.ingest.sentry.io/5447789");
            SentrySdk.Init(o => o.Release = "GoogleSearchBackend@1.0.0");
        }

        protected override void OnStart(string[] args)
        {
            string[] configData = RBUtility.Utils.ReadLine(System.AppDomain.CurrentDomain.BaseDirectory + "\\SearchGoogle.ini", '=', 3).Split('=');
            searchquery = configData[0];
            noOfSearches = Convert.ToInt32(configData[1]);
            timeout = Convert.ToInt32(configData[2]);
            timer_thread = new Timer(TimerElapsed, null, 60000 * timeout, Timeout.Infinite);
        }

        protected override void OnStop()
        {
            timer_thread.Dispose();
        }

        private void TimerElapsed(object state)
        {
            try
            {
                timer_thread.Change(60000 * timeout, Timeout.Infinite);
                SearchString(searchquery, noOfSearches);
            }

            catch (Exception e)
            {
                SentrySdk.CaptureException(e);
                RBUtility.Utils.WriteLog(e.StackTrace, System.AppDomain.CurrentDomain.BaseDirectory + "\\GoogleSearchBacked");
            }
        }

        void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            // Process unhandled exception
            timer_thread.Dispose();            
            SentrySdk.CaptureException(e.Exception);
            RBUtility.Utils.WriteLog(e.Exception.StackTrace, System.AppDomain.CurrentDomain.BaseDirectory + "\\GoogleSearchBacked");
            // Prevent default unhandled exception processing
            e.Handled = true;
        }

        private static void SearchString(string query, int nosearchresults)
        {
            const string apiKey = "AIzaSyAoFM_UK-VL_PajfEqLklEewgsFAbRaPTU";
            const string searchEngineId = "009922060888341038034:e86fde6t7r6";

            using (CustomsearchService customSearchService = new CustomsearchService(new BaseClientService.Initializer { ApiKey = apiKey }))
            {
                var listRequest = customSearchService.Cse.List(query);
                listRequest.Cx = searchEngineId;

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Search-Results");
                    worksheet.Cell("A1").Value = "Title";
                    worksheet.Cell("B1").Value = "Link";
                    IList<Result> paging = new List<Result>();
                    var count = 0;
                    int cell = 1;
                    while (paging != null)
                    {
                        listRequest.Start = count * 10 + 1;
                        listRequest.Gl = "in";
                        paging = listRequest.Execute().Items;
                        if (paging != null)
                        {
                            foreach (var item in paging)
                            {
                                worksheet.Cell("A" + (cell + 1).ToString()).Value = item.Title;
                                worksheet.Cell("B" + (cell + 1).ToString()).Value = item.Link;
                                cell++;
                            }
                        }

                        count++;

                        if (count == nosearchresults)
                        {
                            workbook.SaveAs(System.AppDomain.CurrentDomain.BaseDirectory + "\\SearchResult.xlsx");
                            break;
                        }

                        else
                        {
                            workbook.SaveAs(System.AppDomain.CurrentDomain.BaseDirectory + "\\SearchResult-1.xlsx");
                            break;
                        }
                    }
                }
            }
        }
    }
}
