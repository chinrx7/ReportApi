using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ReportApi.genReport;

namespace ReportApi
{
    public class Write
    {
        public class Datasets
        {
            public string ItemName { get; set; }
            public Records[] Records { get; set; }

        }

        public class WRequest
        {
            public Datasets[] Datasets { get; set; }
        }

        public class Records
        {
            public string Timestamp { get; set; }
            public string Value { get; set; }
            public string Quality { get; set; }
        }

        public static string resStr;

        public static async void WriteBill(string Amount, string Units, DateTime date)
        {
            string TMP = date.ToString("yyyy-MM-ddT17:00:00+07:00");

            Records rec = new Records();
            rec.Value = Amount;
            rec.Quality = "Good";
            rec.Timestamp = TMP;

            Datasets wReq = new Datasets();
            wReq.ItemName = "PLC1\\WEBCAL0\\AMOUNT";
            wReq.Records = new Records[1];
            wReq.Records[0] = rec;

            Datasets[] REQS = new Datasets[2];
            REQS[0] = wReq;

            Records rec2 = new Records();
            rec2.Value = Units;
            rec2.Quality = "Good";
            rec2.Timestamp = TMP;

            Datasets wReq2 = new Datasets();
            wReq2.ItemName = "PLC1\\WEBCAL0\\UNIT";
            wReq2.Records = new Records[1];
            wReq2.Records[0] = rec2;

            REQS[1] = wReq2;

            WRequest wRequest = new WRequest();
            wRequest.Datasets= new Datasets[1];
            wRequest.Datasets = REQS;


            var body = JsonConvert.SerializeObject(wRequest);
            var content = new StringContent(body, Encoding.UTF8, "application/json");


            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri("http://192.168.111.11:9000/");
            var res = client.PostAsync("api/data/write", content);
            var strRes = await res.Result.Content.ReadAsStringAsync();
            resStr = strRes;
        }

        public static async void WriteBillBuilding(string site, string Amount, string Units,string Bath, DateTime date)
        {
            string TMP = date.ToString("yyyy-MM-ddT17:00:00+07:00");

            Records rec = new Records();
            rec.Value = Amount;
            rec.Quality = "Good";
            rec.Timestamp = TMP;

            Datasets wReq = new Datasets();
            wReq.ItemName = "PLC1\\Billing" + site + "\\Bill\\AMOUNT";
            wReq.Records = new Records[1];
            wReq.Records[0] = rec;

            Datasets[] REQS = new Datasets[3];
            REQS[0] = wReq;

            Records rec2 = new Records();
            rec2.Value = Units;
            rec2.Quality = "Good";
            rec2.Timestamp = TMP;

            Datasets wReq2 = new Datasets();
            wReq2.ItemName = "PLC1\\Billing" + site + "\\Bill\\UNIT";
            wReq2.Records = new Records[1];
            wReq2.Records[0] = rec2;

            REQS[1] = wReq2;

            Records rec3 = new Records();
            rec3.Value = Bath;
            rec3.Quality = "Good";
            rec3.Timestamp = TMP;

            Datasets wReq3 = new Datasets();
            wReq3.ItemName = "PLC1\\Billing" + site + "\\Bill\\BATH";
            wReq3.Records = new Records[1];
            wReq3.Records[0] = rec3;

            REQS[2] = wReq3;

            WRequest wRequest = new WRequest();
            wRequest.Datasets = new Datasets[1];
            wRequest.Datasets = REQS;


            var body = JsonConvert.SerializeObject(wRequest);
            var content = new StringContent(body, Encoding.UTF8, "application/json");


            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri("http://192.168.111.11:9000/");
            var res = client.PostAsync("api/data/write", content);
            var strRes = await res.Result.Content.ReadAsStringAsync();
           
        }
    }
}
