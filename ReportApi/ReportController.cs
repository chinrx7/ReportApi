using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Cors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Extensions.Configuration;
using System.Xml.Linq;

namespace ReportApi
{
    public class Report
    {
        public int id   { get; set; }
        public string title { get; set; }
    }

    [EnableCors(origins: "*", headers: "*", methods: "*")]
    public class ReportController : ApiController
    {
        Report[] reports = new Report[]
        {
            new Report{id=1,title="Report1"},
            new Report{id=2,title="Report2"}
        };

        public IEnumerable<Report> Get()
        {
            Util.Logging("", "Get All Requested");
            //System.Diagnostics.Process.Start("CMD.exe");
            //genReports();
            genReport.GetAllReport();
            return reports;
        }

        public string Get(int id)
        {
            Util.Logging("", "Get Requested");
            genReport.GetAllReport();
            //genReports();

            Thread.Sleep(5000);

            return "success";
        }

        public string Get(string Date, string Unit, string Amount,string Type)
        {
            bool res = false;
            DateTime dateRecive;

            res = DateTime.TryParse(Date, out dateRecive);

            string response = "success";
            if (!res) { response= "error"; }
            else { genReport.SelectedReport(dateRecive,Unit,Amount,Type); }

            return response;
        }

        public IHttpActionResult Post()
        {
            return Ok();
        }

        public void genReports()
        {

            string CurFolder = Directory.GetCurrentDirectory();
            string Fpath = CurFolder + @"\Reports\report1.xlsx";
            var ExcelApp = new Application { Visible = false };
            var workbook = ExcelApp.Workbooks.Open(Fpath);
            ExcelApp.CalculateFull();
            workbook.SaveAs(CurFolder + @"\Reports\reportxxx.xlsx");
            workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, (CurFolder + @"\Reports\reportxxx.pdf"));
            workbook.Close();
            ExcelApp.Quit();
        }

    }
}
