using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using System.Net;
using Newtonsoft.Json;
using System.Data;
using System.Runtime.InteropServices;
using System.Globalization;

namespace ReportApi
{
    internal class genReport
    {
        public static BUILDING[] _building;
        public static DateTime repDate;

        public static double SUnit, SAmount;
        public class BUILDING
        {
            public string PLANTID { get; set; }
            public string NAME { get; set; }
            public string CHMIMAP { get; set; }
            public string TYPE { get; set; }
        }

        public class Records
        {
            public string Timestamp { get; set; }
            public string Value { get; set; }
            public string Quality { get; set; }
        }

        public class Req
        {
            public string[] ItemNames { get; set; }
            public string Mode { get; set; }
            public string Timestamp { get; set; }
        }

        public class BUILDINGS
        {
            public string NAME { get; set; }
            public string CODE { get; set; }
            public string PLANTID { set; get; }
            public string M2 { set; get; }
            public string M3 { set; get; }
        }

        public class LIST
        {
            public string NO { get; set; }
            public string SITE { set; get; }
            public string CODE { set; get; }
            public string ID { set; get; }
            public List<BUILDINGS> BUILDINGS { get; set; }
        }

        public class ReportVal
        {
            public string NAME { get; set; }
            public string WHL { set; get; }
            public string WHR { set; get; }
            public string NO {  set; get; }
        }

        public static Records GetVal(string ItemName, string TimeStmp, string port)
        {
            Records res = new Records();

            Req jsonObj = new Req();
            jsonObj.ItemNames = new string[1];
            jsonObj.ItemNames[0] = ItemName;
            jsonObj.Mode = "AtTime";
            jsonObj.Timestamp = TimeStmp;



            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create("http://192.168.111.11:" + port + "/api/data/read");
            httpWebRequest.Method = "POST";
            httpWebRequest.ContentType = "application/json";

            using (var streamWriter = new System.IO.StreamWriter(httpWebRequest.GetRequestStream()))
            {
                string json = JsonConvert.SerializeObject(jsonObj);
                streamWriter.Write(json);
            }
            try
            {
                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new System.IO.StreamReader(httpResponse.GetResponseStream()))
                {
                    var jsonRes = streamReader.ReadToEnd();
                    dynamic json = JObject.Parse(jsonRes);

                    JToken jToken = json["DataSets"][0]["Records"][0];

                    var response = jToken.ToObject<Records>();
                    res = response;
                }
            }
            catch (Exception ex) { Util.Logging("Error", ex.Message); }

            return res;
        }

        public static void GetAllReport()
        {
            //GetBuilding();

            /*foreach(BUILDING bs in _building)
            {
                generateReport(bs.PLANTID, bs.NAME, bs.CHMIMAP);
            }*/

            KillExcel();
        }

        public static void SelectedReport(DateTime reportDate,string Unit, string Amount, string Type)
        {
            SUnit = double.Parse(Unit);
            SAmount= double.Parse(Amount);

            repDate = reportDate;
            GetBuilding(Type);
            KillExcel();
        }

        public static void KillExcel()
        {
            var process = from p in Process.GetProcessesByName("EXCEL")
                          select p;
            foreach (var p in process)
            {
                p.Kill();
            }
        }

        public static string GetMonName(string MonthNO)
        {
            string res = "";

            switch (MonthNO)
            {
                case "01":
                    res = "มกราคม";
                    break;
                case "02":
                    res = "กุมภาพันธ์";
                    break;
                case "03":
                    res = "มีนาคม";
                    break;
                case "04":
                    res = "เมษายน";
                    break;
                case "05":
                    res = "พฤษภาคม";
                    break;
                case "06":
                    res = "มิถุนายน";
                    break;
                case "07":
                    res = "กรกฎาคม";
                    break;
                case "08":
                    res = "สิงหาคม";
                    break;
                case "09":
                    res = "กันยายน";
                    break;
                case "10":
                    res = "ตุลาคม";
                    break;
                case "11":
                    res = "พฤศจิกายน";
                    break;
                case "12":
                    res = "ธันวาคม";
                    break;
            }

            return res;
        }

        public static void GetBuilding(string Type)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            string FMeterName = string.Format("\\meter{0}.json", Type);

            string jstring = File.ReadAllText(Program.path + FMeterName, Encoding.ASCII);

            dynamic json = JObject.Parse(jstring);

            JToken jtk = json["LIST"];

            var building = jtk.ToObject<LIST[]>();

            double price, dprice;


            price = SAmount / SUnit;
            dprice = price - (price*0.35);


            Write.WriteBill(SAmount.ToString(), SUnit.ToString(), repDate);

            foreach(LIST L in building)
            {
                string mName = "\\BillMaster{0}";
                mName = string.Format(mName,L.BUILDINGS.Count.ToString());

                var m2Val = new List<ReportVal>();
                var m3Val = new List<ReportVal>();



                string METERID = "";

                foreach (BUILDINGS B in L.BUILDINGS)
                {
                    m2Val.Add(GetM2(B.CODE, B.PLANTID, B.M2, B.M3.Replace("M2","")));
                    m3Val.Add(GetM3(B.CODE, B.PLANTID, B.M3, B.M3.Replace("TOU", "")));
                    METERID = B.PLANTID;
                }

                string filePath = Program.path + mName;
                Application excelApp = new Application();
                excelApp.DisplayAlerts = false;
                Workbook workbook = excelApp.Workbooks.Open(filePath);
                Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

                CultureInfo cTH = CultureInfo.CreateSpecificCulture("th-TH");

                worksheet.Cells[4, 4] = "หนังสือแจ้งค่าไฟฟ้า ระบบโซล่าเซลล์ ประจำเดือนเดือน " + repDate.ToString("MMMM", cTH) + " " + repDate.ToString("yyyy",cTH);

                worksheet.Cells[15, 9] = price.ToString("##.00");
                worksheet.Cells[17,9] = dprice.ToString("##.00");
                worksheet.Cells[11, 2] = L.CODE;
                worksheet.Cells[11, 4] = L.ID;
                worksheet.Cells[8, 4] = L.SITE;
                worksheet.Cells[11, 6] = METERID;

                int fix = 15;
                int row = fix;
                int nrow =fix + m2Val.Count;

                for (int i=0;i<m3Val.Count;i++)
                {
                    worksheet.Cells[row, 2] = m3Val[i].NAME;
                    worksheet.Cells[row, 5] = m3Val[i].WHL;
                    worksheet.Cells[row, 4] = m3Val[i].WHR;

                    double AMount, Unit;
                    Unit = double.Parse(m3Val[i].WHR) - double.Parse(m3Val[i].WHL);
                    AMount = Unit * dprice;

                    Write.WriteBillBuilding(m3Val[i].NO, AMount.ToString("00.00"), Unit.ToString("00.00"), dprice.ToString("00.00"), repDate);

                    row++;
                }

                for (int ii = 0; ii < m2Val.Count; ii++)
                {
                    worksheet.Cells[nrow, 2] = m2Val[ii].NAME;
                    worksheet.Cells[nrow, 5] = m2Val[ii].WHL;
                    worksheet.Cells[nrow, 4] = m2Val[ii].WHR;

                    nrow++;
                }

                // worksheet.Cells[30, 5] = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 20).ToShortDateString();

                DateTime dateTime = repDate;

                string yDir = Program._set.OUTPUT + "\\" + dateTime.ToString("yyyy");
                string mDir = Program._set.OUTPUT + "\\" + dateTime.ToString("yyyy") + "\\" + dateTime.ToString("MMM");

                if (!Directory.Exists(yDir)) { Directory.CreateDirectory(yDir); }
                if (!Directory.Exists(mDir)) { Directory.CreateDirectory(mDir); }

                string fFormat = Program._set.OUTPUT + "{0}\\{1}\\{2}.xlsx";
                string fFormat2 = Program._set.OUTPUT + "{0}\\{1}\\{2}.pdf";

                string Savepath = string.Format(fFormat, dateTime.ToString("yyyy"), dateTime.ToString("MMM"), L.NO);
                string Savepath2 = string.Format(fFormat2, dateTime.ToString("yyyy"), dateTime.ToString("MMM"), L.NO);
                //Savepath = "C:\\PVV\\test" + ChmiMAP +".xlsx";
                if (File.Exists(Savepath)) { File.Delete(Savepath); }
                if (File.Exists(Savepath2)) { File.Delete(Savepath2); }

                workbook.SaveAs(Savepath);
                workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, Savepath2);
                workbook.Close();
                excelApp.Quit();
            }
            GenSumReport(price, dprice);
            
        }

        public static void GenSumReport(double uPrice, double dPrice)
        {
            DateTime dateTime = repDate;
            string fPdf = Program._set.OUTPUT + "{0}\\{1}\\00.pdf";
            string xCel = Program._set.OUTPUT + "{0}\\{1}\\00.xlsx";

            string SavePdf = string.Format(fPdf, dateTime.ToString("yyyy"), dateTime.ToString("MMM"));
            string SaveXcel = string.Format(xCel, dateTime.ToString("yyyy"), dateTime.ToString("MMM"));

            string Tag = "PLC1\\Calculation00\\Energy\\WH";
            string StartTime = dateTime.ToString("yyyy-MM-dd") + "T17:00:00+07:00";

            Records rec = GetVal(Tag, StartTime, "9000");

            double sWH = double.Parse(rec.Value);

            dateTime = repDate.AddMonths(1);
            string EndTime = dateTime.ToString("yyyy-MM-dd") + "T17:00:00+07:00";

            rec = GetVal(Tag, EndTime , "9000");
            double eWh = double.Parse(rec.Value);

            string MasterPath = Program.path + "\\BSUMBILLMASTER.xlsx";
            Application excelApp = new Application();
            excelApp.DisplayAlerts = false;
            Workbook workbook = excelApp.Workbooks.Open(MasterPath);
            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

            CultureInfo cTH = CultureInfo.CreateSpecificCulture("th-TH");

            worksheet.Cells[2,1] = "หนังสือแจ้งค่าไฟฟ้าระบบโซลาเซลล์ ประจำเดือน " + repDate.ToString("MMMM", cTH) + " " + repDate.ToString("yyyy", cTH);
            worksheet.Cells[10, 10] = dateTime.ToString("dd MMMM yyyy", cTH);

            worksheet.Cells[14, 4] = eWh.ToString();
            worksheet.Cells[14, 5] = sWH.ToString();

            worksheet.Cells[13, 10] = uPrice.ToString();
            worksheet.Cells[14, 10] = dPrice.ToString();

            if(File.Exists(SaveXcel)) { File.Delete(SaveXcel); }
            if (File.Exists(SavePdf)) { File.Delete(SavePdf); }

            workbook.SaveAs(SaveXcel);
            workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, SavePdf);
            workbook.Close();
            excelApp.Quit();
        }

        public static ReportVal GetM2(string Name,string PlantID, string M2MAP, string NO)
        {
            var Rval = new ReportVal();
            Rval.NAME = "พลังงาน solar ที่ไหลออก อาคาาร " + Name;

            DateTime dateTime1 = repDate;
            string StartTime = dateTime1.ToString("yyyy-MM-dd") + "T17:00:00+07:00";


            string Tag = "PLC1\\M2\\" + M2MAP + "\\TOU_WHR";

            Records rec = GetVal(Tag, StartTime, "9000");

            Rval.WHL = rec.Value;
            if (Rval.WHL == null) { Rval.WHL = "0"; }

            DateTime EndTime1 = repDate.AddMonths(1);
            string EndTime = EndTime1.ToString("yyyy-MM-dd") + "T17:00:00+07:00";

            rec = GetVal(Tag, EndTime, "9000");

            Rval.WHR = rec.Value;
            if (Rval.WHR == null) { Rval.WHR = "0"; }
            Rval.NO= NO;

            return Rval;
        }

        public static ReportVal GetM3(string Name, string PlantID, string M3MAP, string NO)
        {
            var Rval = new ReportVal();
            Rval.NAME = "พลังงาน Solar อาคาร " + Name;

            DateTime dateTime1 = repDate;
            string StartTime = dateTime1.ToString("yyyy-MM-dd") + "T17:00:00+07:00";



            string Tag = "PLC1\\M3\\" + M3MAP + "\\WH_EXP";

            Records rec = GetVal(Tag, StartTime, "9000");

            Rval.WHL = rec.Value;
            if (Rval.WHL == null) { Rval.WHL = "0"; }

            DateTime EndTime1 = repDate.AddMonths(1);
            string EndTime = EndTime1.ToString("yyyy-MM-dd") + "T17:00:00+07:00";

            rec = GetVal(Tag, EndTime, "9000");

            Rval.WHR = rec.Value;
            if (Rval.WHR == null) { Rval.WHR = "0"; }
            Rval.NO= NO;

            return Rval;
        }



        public static void generateReport(string PlantID,string Name,string ChmiMAP)
        {
            DateTime dateTime1 = new DateTime(DateTime.Now.Year, DateTime.Now.Month - 1, 1);
            dateTime1 = dateTime1.AddMonths(-1);
            string StartTime = dateTime1.ToString("yyyy-MM-dd") + "T17:00:00+07:00";


            Dictionary<string,string> Tags = new Dictionary<string,string>();
            Dictionary<string, string> Vals = new Dictionary<string, string>();
            Tags["TOU_WHR"] = "PLC1\\M2\\"+ ChmiMAP + "\\TOU_WHD";
            Tags["WH"] = "PLC1\\M2\\" +  ChmiMAP.Replace("M2","TOU") + "\\WH";

            foreach (KeyValuePair<string,string> kv in Tags)
            {
               Records rec = GetVal(kv.Value,StartTime, "9000");

   
               Vals[kv.Key + "_N"] = rec.Value;
                if (Vals[kv.Key + "_N"] == null) { Vals[kv.Key + "_N"] = "0"; }


            }

            DateTime EndTime1 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            string EndTime = EndTime1.ToString("yyyy-MM-dd") + "T17:00:00+07:00";



            foreach (KeyValuePair<string, string> kv in Tags)
            {
                Records rec = GetVal(kv.Value, EndTime, "9000");
                Vals[kv.Key + "_P"] = rec.Value;
                if (Vals[kv.Key + "_P"] == null) { Vals[kv.Key + "_P"] = "0"; }

            }


            string filePath = Program.path + "\\BillMaster.xlsx";

            Application excelApp = new Application();
            excelApp.DisplayAlerts = false;
            Workbook workbook = excelApp.Workbooks.Open(filePath);
            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

            worksheet.Cells[11, 2] = ChmiMAP;
            worksheet.Cells[8, 4] = Name;
            worksheet.Cells[11, 6] = PlantID;

            worksheet.Cells[16, 4] = Vals["TOU_WHR_N"];
            worksheet.Cells[15, 4] = Vals["WH_N"];
            worksheet.Cells[16, 5] = Vals["TOU_WHR_P"];
            worksheet.Cells[15, 5] = Vals["WH_P"];

            DateTime dateTime = DateTime.Now;
            
            string yDir = Program._set.OUTPUT + "\\" + dateTime.ToString("yyyy");
            string mDir = Program._set.OUTPUT + "\\" + dateTime.ToString("yyyy") + "\\" + dateTime.ToString("MMM");

            if (!Directory.Exists(yDir)) { Directory.CreateDirectory(yDir); }
            if (!Directory.Exists(mDir)) { Directory.CreateDirectory(mDir); }

            string fFormat = Program._set.OUTPUT + "{0}\\{1}\\{2}.xlsx";
            string fFormat2 = Program._set.OUTPUT + "{0}\\{1}\\{2}.pdf";

            string Savepath = string.Format(fFormat, dateTime.ToString("yyyy"), dateTime.ToString("MMM"), ChmiMAP);
            string Savepath2 = string.Format(fFormat2, dateTime.ToString("yyyy"), dateTime.ToString("MMM"), ChmiMAP);
            //Savepath = "C:\\PVV\\test" + ChmiMAP +".xlsx";
            if (File.Exists(Savepath)) { File.Delete(Savepath); }
            if (File.Exists(Savepath2)) { File.Delete(Savepath2); }

            workbook.SaveAs(Savepath);
            workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, Savepath2);
            workbook.Close();
            excelApp.Quit();

            //Worksheet worksheet = workbook.

        }

    }
}
