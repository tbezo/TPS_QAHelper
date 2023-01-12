# TPS_QAHelper
ESAPI Script (Binary-Plugin) to upload test data from Eclipse to QAtrack+ via API

How we use the script:
1. Open TPS QA patient
2. Open series/course "QS Eclipse QAT" (contains 30 plans and two plansums)
3. Recalculate all plans
3. Run script TPS_QAHelper
4. Go to QATrack+ Website and Review test lists

### Hints on how to use the script
__In Eclipse:__
Create a course and plans on a water phantom with different field sizes, energies, beam shapes, wedges, etc. Enumerate the plans with the plan name (not the id) a.e. 01...99. Insert some reference points into plans and structures in the phantom. Add all plans to a plan sum for easy recalculation. Add some plans to a second plan sum (that is the one that will get the DVH data evaluated).

__In QATrack+:__
Create a test for each reference point in each plan, the slug should be _planname_refptName_ (a.e. 05_Rpt01). Also create one test for each plan to store the calculated monitor units (slug: _planname_MU_). Add the tests to a test list (or create sublists first).
Create tests for the DVH data. Structures starting with "structString" (see script config) are being evaluated for min, max, mean and median dose. The test slugs should be named {first five characters of the structure name}_min, _max, _mean, _median (geo_D_min, geo_K_median, etc.). If you need more than five characters to discriminate the different structures change the ExtractDVHData method. Add all tests for the DVH data to a second test list.
Don't forget to assing both test lists to your Eclipse unit.

__Visual Studio:__
Add NewtonSoft.Json via NuGet. Adjust path to Esapi .dll files if needed. (Tested with Esapi 16.1 - default path C:\Program Files\Varian\RTM\16.1\esapi\API\).
Insert API base address, token, unit name and the names of the QATrack+ test lists. Also set the Eclipse course name, plansum name and the starting string
of the DVH structures to evaluate. When all this is done compile the binary-plugin.

Qatrack+ documentation on how to create an API key: [Using the QATrack+ API](https://docs.qatrackplus.com/en/stable/api/guide.html)


### Debugging
The script creates two files, httpContent.json and response.txt, in the script folder. The first one contains the API payload that is being uploaded. Look there if you want to see what tests are being expected. The second one contains the API response from QATrack+. In case of errors you should be able to find them there.
