using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using excel2vcard.Models;
using System.IO;
using System.Text;

namespace excel2vcard.Controllers
{
    public class HomeController : Controller
    {
        Reader excelRead = new Reader();
        
      

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }
        public ActionResult Index()
        {
            returned tosend = new returned() { messageType = String.Empty };
            return View(tosend);
        }

        [HttpPost]
        public ActionResult Contact_v()
        { 

            returned tosend = new returned() { messageType = String.Empty };
            DataTable DATA = new DataTable();

            if (Request.Files["FileUpload1"].ContentLength > 0)
            {
                string extension = System.IO.Path.GetExtension(Request.Files["FileUpload1"].FileName).ToLower();
                string connString = "";


                string[] validFileTypes = { ".xls", ".xlsx", ".csv" };

                string path1 = string.Format("{0}/{1}", Server.MapPath("~/Content/Uploads"), Request.Files["FileUpload1"].FileName);
                if (!Directory.Exists(path1))
                {
                    Directory.CreateDirectory(Server.MapPath("~/Content/Uploads"));
                }
                if (validFileTypes.Contains(extension))
                {
                    if (System.IO.File.Exists(path1))
                    { System.IO.File.Delete(path1);
                    }
                    Request.Files["FileUpload1"].SaveAs(path1);
                    if (extension == ".csv")
                    {
                        DataTable dt = excelRead.ConvertCSVtoDataTable(path1);
                        ViewBag.Data = dt;
                        DATA = dt;
                    }
                    //Connection String to Excel Workbook  
                    else if (extension.Trim() == ".xls")
                    {
                        connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path1 + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                        DataTable dt = excelRead.ConvertXSLXtoDataTable(path1, connString);
                        ViewBag.Data = dt;
                        DATA = dt;
                    }
                    else if (extension.Trim() == ".xlsx")
                    {
                        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path1 + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                        DataTable dt = excelRead.ConvertXSLXtoDataTable(path1, connString);
                        ViewBag.Data = dt;
                        DATA = dt;
                    }

                }

            }
            else
            {
                ViewBag.Error = "excelError";
            }



            StringBuilder vcf = new StringBuilder();
            
            foreach (DataRow dr in (DATA as System.Data.DataTable).Rows)
            {
                System.Text.Encoding utf_8 = System.Text.Encoding.UTF8;

                string s = dr[DATA.Columns["name"]].ToString();
                string adress = dr[DATA.Columns["note"]].ToString();

                if (s == String.Empty || String.IsNullOrWhiteSpace(s))
                {
                    continue;
                }

                vcf.Append("BEGIN:VCARD" + System.Environment.NewLine);
                vcf.Append("VERSION:2.1" + System.Environment.NewLine);
                vcf.Append("N;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:;" + string_convert(s).ToUpper() + ";;;" + System.Environment.NewLine);
                vcf.Append("FN;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:" + string_convert(s).ToUpper() + System.Environment.NewLine);
                vcf.Append("TEL;CELL;PREF:" + dr[DATA.Columns["phone"]].ToString() + System.Environment.NewLine);
                vcf.Append("ADR;HOME;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:;;" + string_convert(adress).ToUpper()+";;;;" + System.Environment.NewLine);
                vcf.Append("END:VCARD" + System.Environment.NewLine);
            }

            string name = Server.MapPath(@"~\temp\") + DateTime.Now.Day.ToString() + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + DateTime.Now.Millisecond.ToString() + ".vcf";

            FileInfo info = new FileInfo(name);


            if (!info.Exists)
            {
                using (StreamWriter writer = info.CreateText())
                {
                    
                    writer.Write(vcf.ToString());
                }
            }
            try
            {
            FileStream send = info.OpenRead();
            
            return File(name , "text/plain","contact.vcf");
            }
            catch (Exception)
            {

                throw;
            }
       
        }

        private String string_convert(String S)
        {

            byte[] utf = System.Text.Encoding.UTF8.GetBytes(S);

            String temp_s = "";
            foreach (byte item in utf)
            {
                decimal x = Convert.ToDecimal(item);
                int temp = Convert.ToInt32(x);
                temp_s += String.Format("={0:x}", temp);
            }

            return temp_s;

        }
    }
}