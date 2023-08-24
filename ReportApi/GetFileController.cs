using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http.Headers;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Cors;
using System.Web.Http.Results;

namespace ReportApi
{
    public class GetFile
    {
        public int id { get; set; }
    }

    [EnableCors(origins: "*", headers: "*", methods: "*")]
    public class GetFileController : ApiController
    {
        public HttpResponseMessage GetFile(string date, string no, string mode)
        {
            HttpResponseMessage response = new HttpResponseMessage();
            string folderPath = Program._set.Report;

            DateTime rDate;

            bool resPDate = DateTime.TryParse(date, out rDate);

            if (resPDate)
            {
                string ext = "";
                if (mode == "view") { ext = ".pdf"; } else { ext = ".xlsx"; }

                string fFormat = folderPath + "{0}\\{1}\\{2}" + ext;

                string rFilePath = string.Format(fFormat, rDate.ToString("yyyy"),rDate.ToString("MMM"),no);

                var byteArr = File.ReadAllBytes(rFilePath);
                var stream = new MemoryStream(byteArr);
                var res = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ByteArrayContent(stream.ToArray())
                };
                res.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment")
                {
                    FileName = rDate.ToString("MMM-yyyy") + "_" + no+ext
                };
                res.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                response = res;
            }
            else
            {
                var res = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(
                        "Ivalid date format!!!",
                        Encoding.UTF8,
                        "text/html"
                        )
                };
                response= res;
            }

            return response;
        }

       
    }
}
