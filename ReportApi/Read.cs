using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading.Tasks.Sources;

namespace ReportApi
{
    public class Read
    {
        public class Request
        {
            public string[] ItemNames { get; set; }
            public string Mode { get; set; }
            public string Timestamp { get; set; }
        }

        public class Datasets
        {
            public string ItemName { get; set; }
            public int ItemId { get; set; }
            public Records[] Records { get; set; }
        }

        public class Records
        {
            public string Timestamp { get; set; }
            public string Value { get; set; }
            public string Quality { get; set; }
        }

        public class items
        {
            public string[] ItemName { get; set; }
        }


        public static Datasets[] res;


        public static async void ReadAtime(string[] tags)
        {
            Request req = new Request()
            {
                ItemNames = tags,
                Mode = "AtTime",
                //Timestamp = DateTime.Now.ToString("yyyy-MM-dd") + "T17:00:00+07:00"
                Timestamp = DateTime.Now.ToString("yyyy-MM-dd") + "T" + DateTime.Now.AddMinutes(-1).ToString("HH:mm:ss")
            };
            HttpResponseMessage response = new HttpResponseMessage();
            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri("http://192.168.111.11:9000/");
            client.DefaultRequestHeaders.Clear();

            var body = JsonConvert.SerializeObject(req);

            var content = new StringContent(body, Encoding.UTF8, "application/json");

            var responseMessage = client.PostAsync("api/data/read", content);

            var cnt = await responseMessage.Result.Content.ReadAsStringAsync();

            dynamic json = JObject.Parse(cnt);
            //JToken jToken = json["DataSets"][0]["Records"];
            JToken jToken = json["DataSets"];

            var Datasets = jToken.ToObject<Datasets[]>();

            res = Datasets;
        }
    }
}
