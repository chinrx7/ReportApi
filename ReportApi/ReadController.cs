using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ReportApi
{
    public class ReadController : ApiController
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
            public Records Records { get; set; }
        }

        public class Records
        {
            public string Timestamp { get; set; }
            public double Value { get; set; }
            public string Quality { get; set; }
        }



        public HttpResponseMessage ReadAtTime([FromBody] string[] Tag)
        {

            try
            {
                Read.ReadAtime(Tag);
            }
            catch { }

            HttpResponseMessage resp = new HttpResponseMessage();
            resp.StatusCode = HttpStatusCode.OK;
            var res = JsonConvert.SerializeObject(Read.res);
            var content = new StringContent(res, Encoding.UTF8, "application/json");
            resp.Content = content;
            return resp;

        }
    }
}
