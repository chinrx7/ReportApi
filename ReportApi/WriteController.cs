using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;

namespace ReportApi
{
    public class WriteController : ApiController
    {

        public HttpResponseMessage WriteVal(string date)
        {
            Write.WriteBill("100000", "20", DateTime.Now);
            HttpResponseMessage httpResponseMessage = new HttpResponseMessage();
            httpResponseMessage.StatusCode = System.Net.HttpStatusCode.OK;
            var res = JsonConvert.SerializeObject(Write.resStr);
            var content = new StringContent(res, Encoding.UTF8, "application/json");
            httpResponseMessage.Content = content;
            return httpResponseMessage;
        }

    }
}
