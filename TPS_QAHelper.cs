using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using System.Threading.Tasks;

[assembly: AssemblyVersion("0.1.0.2")]

namespace VMS.TPS
{
    public class Script
    {
        public Script()
        {
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)
        {
            // QATrack+ API config
            string baseaddress = "";                                                // web address for qatrack API (a.e. https://qatrack.contoso.org/api)
            string token = "";                                                      // API token
            string uname = "Eclipse";                                               // unit name
            string utcName1 = "TPS QA Dosisberechnung";                             // names of the QATrack+ Unit Test Collections (testlist name)
            string utcName2 = "TPS QA DVH Data";

            // Eclipse config
            string user = context.CurrentUser.Id;                                   // Eclipse user
            string courseName = "QS Eclipse QAT";                                   // course with all QA plans
            string psumName = "DVH-Summe";                                          // plan sum to extract DVH data from
            string structString = "geo";                                            // starting string of structure ids to evaluate

            // try to get the QATrack+ unit test collection to post data to
            string utc_url;
            Task<string> utcTask;
            try
            {
                utcTask = GetUnitTestCollectionAsync(utcName1, uname, baseaddress, token);
            }
            catch (Exception e)
            {
                MessageBox.Show("GetUnitTestCollection Error:" + e.Message);
                return;
            }

            // Extract reference point doses and monitor units from all plans
            Dictionary<string, TestData> tdd; // = new Dictionary<string, TestData>();
            try
            {
                tdd = ExtractPlanData(courseName, context);
            }
            catch (Exception e)
            {
                MessageBox.Show("ExtractPlanData Error: " + e.Message);
                return;
            }

            // wait for QATrack+ anwser
            utcTask.Wait();
            utc_url = utcTask.Result;

            // post TestData for Referenze Points and Monitor Units
            Task<int> res1Task = PostTestDataAsync(tdd, user, utc_url, baseaddress, token);
            res1Task.Wait();

            // try to get the QATrack+ unit test collection to post data to
            string utc_url2;
            Task<string> utcTask2;
            try
            {
                utcTask2 = GetUnitTestCollectionAsync(utcName2, uname, baseaddress, token);
            }
            catch (Exception e)
            {
                MessageBox.Show("GetUnitTestCollection Error:" + e.Message);
                return;
            }

            // Extract DVH Data from geometric structures
            Dictionary<string, TestData> tdd2; // = new Dictionary<string, TestData>();
            try
            {
                tdd2 = ExtractDVHData(psumName, structString, context);
            }
            catch (Exception e)
            {
                MessageBox.Show("ExtractDVHData Error: " + e.Message);
                return;
            }

            // wait for QATrack+ anwser
            utcTask2.Wait();
            utc_url2 = utcTask2.Result;

            // post TestData for Referenze Points and Monitor Units
            Task<int> res2Task = PostTestDataAsync(tdd2, user, utc_url2, baseaddress, token);
            res2Task.Wait();

        }

        /// <summary>
        /// Search for the matching Unit Test Collection URL to post TestData to.
        /// </summary>
        /// <param name="tlname">QATrack+ test list name</param>
        /// <param name="uname">QATrack+ unit name</param>
        /// <param name="baseaddress">QATrack+ API URI (e.g. "https://qatrack.contoso.org/api/")</param>
        /// <param name="token">QATrack+ API Token</param>
        /// <exception cref="Exception">Unit test collection not unique</exception>
        /// <returns>URL of Unit Test Collection from given unit and testlist</returns>
        public async Task<string> GetUnitTestCollectionAsync(string tlname, string uname, string baseaddress, string token)
        {
            UnitTestCollection_Results jres;

            // example from https://gist.github.com/acamino/51ae7fa45708bc1e8bcda5657374aa48
            using (HttpClient client = new HttpClient())
            {
                // client configuration
                client.BaseAddress = new Uri(baseaddress);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Token", token);
                client.DefaultRequestHeaders.Add("User-Agent", "Anything");
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                // search unit test collection to post data to
                HttpResponseMessage response = await client.GetAsync(string.Format("qc/unittestcollections/?unit__name__icontains={0}&test_list__name__icontains={1}", uname, tlname));
                response.EnsureSuccessStatusCode();
                string resString = await response.Content.ReadAsStringAsync(); //JSON string which can be converted. ReadAsAsync from example is not available!
                jres = JsonConvert.DeserializeObject<UnitTestCollection_Results>(resString);

                if (jres.count != 1)
                {
                    throw new Exception("More than one matching unit test list found!");
                }
            }
            return jres.results[0].url;
        }

        /// <summary>
        /// Post TestData to QATrack+ API
        /// </summary>
        /// <param name="tests">Dictionary with test macro name and TestData</param>
        /// <param name="utc_url">QATrack+ Unit Test Collection URL to post to</param>
        /// <param name="baseaddress"></param>
        /// <param name="token">QATrack+ API Token</param>
        /// <returns>0 if ok, 1 if error occured</returns>
        public async Task<int> PostTestDataAsync(Dictionary<string, TestData> tests, string user, string utc_url, string baseaddress, string token)
        {
            //get the full location of the assembly with DaoTests in it
            string dllLocation = Assembly.GetExecutingAssembly().Location;
            //get the folder that's in
            string dllDirectory = Path.GetDirectoryName(dllLocation);

            using (HttpClient client = new HttpClient())
            {
                client.BaseAddress = new Uri(baseaddress);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Token", token);
                client.DefaultRequestHeaders.Add("User-Agent", "Anything");
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                // DateTime without last digit from miliseconds
                DateTime dt = DateTime.Now.ToUniversalTime();
                DateTime dts = dt.AddTicks(-(dt.Ticks % 10));

                // get id from utc_url
                string[] subs = utc_url.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                string utcid = subs.Last();

                // UTC data to post
                TestList_Post TestList = new TestList_Post
                {
                    unit_test_collection = utc_url,
                    work_started = dts.AddHours(1), //Timezone is not getting used in QAT+?
                    work_completed = dts.AddSeconds(3610),
                    user_key = "Id" + utcid + "_" + DateTime.Now.ToShortDateString(),
                    tests = tests,
                    comment = "Files uploaded by Eclipse user " + user
                };

                //Create HTTP payload from JSON String. https://stackoverflow.com/questions/23585919/send-json-via-post-in-c-sharp-and-receive-the-json-returned
                string stringPayload = JsonConvert.SerializeObject(TestList);
                StringContent httpContent = new StringContent(stringPayload, Encoding.UTF8, "application/json");
                File.AppendAllText(dllDirectory + @"\httpContent.json", httpContent.ReadAsStringAsync().Result + "\n"); //for debugging Testlists, can be disabled if all works as expected

                // Post to Server
                HttpResponseMessage httpResponse = await client.PostAsync("qa/testlistinstances/", httpContent);
                if (httpResponse.Content != null)
                {
                    string responseContent = await httpResponse.Content.ReadAsStringAsync();
                    //do stuff with response here
                    if (responseContent.Contains("error") || responseContent.Contains("already exists"))
                    {
                        MessageBox.Show("QATrack+ response contains errors, please check " + dllDirectory + @"\response.txt");
                        return 1;
                    }
                    File.AppendAllText(dllDirectory + @"\response.txt", responseContent + "\n"); //for debugging Testlists, can be disabled if all works as expected
                }
                return 0;
            }
        }

        /// <summary>
        /// Extracts Reference Point doses and monitor units from all plans in given course.
        /// Creates TestData Object with macro name string in the form "p{Plan.Name}_{RefPt.Id}"
        /// </summary>
        /// <param name="qaCourseName">Id of the course</param>
        /// <param name="context">Eclipse Context</param>
        /// <param name="renormElectron">Renormalize electron plan dose values to 100 MU</param>
        /// <returns>Dictionary of QATrack+ test macro name and TestData values</returns>
        /// <exception cref="Exception">Exception no plans found</exception>
        public Dictionary<string, TestData> ExtractPlanData(string qaCourseName , ScriptContext context, bool renormElectron = true)
        {
            // dict for test macro names and test data
            Dictionary<string, TestData> tdd = new Dictionary<string, TestData>();

            List<ExternalPlanSetup> ps = new List<ExternalPlanSetup>();
            ps = context.ExternalPlansInScope.Where(x => x.Course.Id == qaCourseName).ToList();
            if (ps.Count == 0)
            {
                throw new Exception("Please open series " + qaCourseName);
            }

            // extract dose from Reference Points
            foreach (ExternalPlanSetup plan in ps.OrderBy(x => x.Name)) // Plan names: 01, 02, 03, ...
            {
                // create list of beams and sum up the MUs for testdata
                string tname2 = "p" + plan.Name + "_MU";
                List<Beam> beams = plan.Beams.Where(y => !(y.IsSetupField)).ToList();
                double mu_sum = double.NaN;
                mu_sum = beams.Select(x => x.Meterset.Value).Sum();
                if (double.IsNaN(mu_sum))
                {
                    string beam_values = "";
                    foreach (Beam nanbeam in beams)
                    {
                        beam_values += nanbeam.Id + ": " + nanbeam.Meterset.Value.ToString() + " MU\n";
                    }
                    MessageBox.Show("Error: MU sum is NaN in Plan " + plan.Name + ". Value skipped." + "\n" + beam_values);
                    tdd.Add(tname2, new TestData { skipped = true });
                }
                else { tdd.Add(tname2, new TestData { value = mu_sum }); }

                // renormalize dose value from electron plans to 100 MU, since Eclipse won't keep the MU between recalculations.
                double renorm = 1.0;
                if (renormElectron)
                {
                    if (plan.Beams.Count(x => x.EnergyModeDisplayName.Contains("E")) == 1)
                    {
                        renorm = 100.0 / mu_sum;
                    }
                }

                // get dose from reference points for testdata
                foreach (ReferencePoint rp in plan.ReferencePoints.OrderBy(x => x.Id))
                {
                    string tname1 = "p" + plan.Name + "_" + rp.Id;
                    tdd.Add(tname1, new TestData { value = DoseAtRefPt(rp, plan) * renorm });
                }

            }
            return tdd;
        }

        /// <summary>
        /// Extracts DVH data (min, max, mean, media) from plan sum for all structures with id
        /// starting with given string. Creates TestData Object with macro name string in the form "{str.Id}[0:5]_{[min/max/mean/median]}"
        /// </summary>
        /// <param name="planSumId">Id of the plan sum</param>
        /// <param name="structs">Start of string for Structures to evaluate</param>
        /// <param name="context">Eclipse Context</param>
        /// <returns>Dictionary of QATrack+ test macro name and TestData values</returns>
        /// <exception cref="Exception">Plan sum not found</exception>
        public Dictionary<string, TestData> ExtractDVHData(string planSumId, string structs, ScriptContext context)
        {
            // dict for test macro names and test data
            Dictionary<string, TestData> tdd2 = new Dictionary<string, TestData>();
            PlanSum Sum1;
            try
            {
                Sum1 = context.Course.PlanSums.Single(x => x.Id == planSumId); //"DVH-Summe"
            }
            catch
            {
                throw new Exception("Please open plan sum " + planSumId);
            }

            // get DVH Data from Plansum
            foreach (Structure str in Sum1.StructureSet.Structures.Where(x => x.Id.StartsWith(structs))) //"geo"
            {
                string sub_id = str.Id.Substring(0, 5);
                DVHData tup = Sum1.GetDVHCumulativeData(str, DoseValuePresentation.Absolute, VolumePresentation.Relative, 1);
                tdd2.Add(sub_id + "_min", new TestData { value = tup.MinDose.Dose });
                tdd2.Add(sub_id + "_max", new TestData { value = tup.MaxDose.Dose });
                tdd2.Add(sub_id + "_mean", new TestData { value = tup.MeanDose.Dose });
                tdd2.Add(sub_id + "_median", new TestData { value = tup.MedianDose.Dose });
            }
            return tdd2;
        }

        /// <summary>
        /// Returns the absolute dose value at a given reference point in External Beam Plan.
        /// </summary>
        /// <param name="rpt">Reference Point</param>
        /// <param name="ps">Plan</param>
        /// <returns>Absolute dose value at reference point</returns>
        public double DoseAtRefPt(ReferencePoint rpt, ExternalPlanSetup ps)
        {
            ps.DoseValuePresentation = DoseValuePresentation.Absolute;
            DoseValue dose = ps.Dose.GetDoseToPoint(rpt.GetReferencePointLocation(ps));
            return dose.Dose;
        }

        // class for QATrack+ test content
        public class TestData
        {
            public double value { get; set; }
            public string string_value { get; set; }
            //public string filename { get; set; } //for file uploads value needs to be text or bas64 encoded content!
            //public string encoding { get; set; }
            public bool skipped { get; set; }
            public string comment { get; set; }
        }
        // class for QATrack+ testlist content
        public class TestList_Post
        {
            public string unit_test_collection { get; set; }
            public int day { get; set; }
            public bool in_progress { get; set; }
            public bool include_for_scheduling { get; set; }
            public DateTime work_started { get; set; }
            public DateTime work_completed { get; set; }
            public string user_key { get; set; }
            public string comment { get; set; } //comment may not be null or blank, comment if not needed
            public Dictionary<string, TestData> tests { get; set; }
        }
        // class for QATrack+ unit test collection
        public class UnitTestCollection
        {
            public string url { get; set; }
            public string tests_object { get; set; }
            public string next_test_list { get; set; }
            public int next_day { get; set; }
            public string due_date { get; set; } //Newtonsoft.Json can't handle NULL DateTime. Replaced with string.
            public bool auto_schedule { get; set; }
            public bool active { get; set; }
            public int object_id { get; set; }
            public string name { get; set; }
            public string unit { get; set; }
            public string frequency { get; set; }
            public string assigned_to { get; set; }
            public string content_type { get; set; }
            public string last_instance { get; set; }
            public IList<string> visible_to { get; set; }
        }
        // class to deserialize QATrack+ json answer to utc query
        public class UnitTestCollection_Results
        {
            public int count { get; set; }
            public object next { get; set; }
            public object previous { get; set; }
            public IList<UnitTestCollection> results { get; set; }
        }

    }
}
