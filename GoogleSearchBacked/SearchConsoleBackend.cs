using ClosedXML.Excel;
using Google.Apis.Customsearch.v1;
using Google.Apis.Customsearch.v1.Data;
using Google.Apis.Services;
using NLog;
using Sentry;
using System;
using System.Collections.Generic;
using System.Globalization;
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
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();


        public SearchConsoleBackend()
        {
            InitializeComponent();
            SentrySdk.Init(System.Configuration.ConfigurationManager.AppSettings["sentry_api"]);
            SentrySdk.Init(o => o.Release = "GoogleSearchBackend@1.0.0");
        }

        protected override void OnStart(string[] args)
        {
            string test1 = System.Configuration.ConfigurationManager.AppSettings["sentry_api"];
            Logger.Info("Search Service Service Started");            
            searchquery = System.Configuration.ConfigurationManager.AppSettings["searchquery"];
            noOfSearches = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["noOfSearches"],CultureInfo.InvariantCulture);
            timeout = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["timeout"],CultureInfo.InvariantCulture);
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

            catch (ObjectDisposedException e)
            {
                SentrySdk.CaptureException(e);
                Logger.Error(e,"Error Occured While in Time Elapsed");
            }

            catch (ArgumentOutOfRangeException e)
            {
                SentrySdk.CaptureException(e);
                Logger.Error(e, "Error Occured While in Time Elapsed");
            }
        }

        private static void SearchString(string query, int nosearchresults)
        {
            string apiKey = System.Configuration.ConfigurationManager.AppSettings["apiKey"];
            string searchEngineId = System.Configuration.ConfigurationManager.AppSettings["searchEngineId"];

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
                                worksheet.Cell("A" + (cell + 1).ToString(CultureInfo.InvariantCulture)).Value = item.Title;
                                worksheet.Cell("B" + (cell + 1).ToString(CultureInfo.InvariantCulture)).Value = item.Link;
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
