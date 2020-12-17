using ClosedXML.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using System.Web.UI.WebControls;
using WebApplication1.Models;
using System.Net;

namespace WebApplication1.Controllers
{
   
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Title = "Home Page";

            return View();
        }
        public ActionResult NewIndex()
        {
            ViewBag.Title = "Home Page";

            return View();
        }
        [AllowAnonymous]
        public ActionResult VC()
        {
            
            return View();
        }
        [HttpPost]
        public ActionResult ValidateCaptcha()
        {
            var response = Request["g-recaptcha-response"];
            //secret that was generated in key value pair
            const string secret = "6LeuyKwUAAAAALASWxF-Aqev7lrGj3gsJiNadvgK";

            var client = new WebClient();
            var reply =
                client.DownloadString(
                    string.Format("https://www.google.com/recaptcha/api/siteverify?secret={0}&response={1}", secret, response));

            var captchaResponse = JsonConvert.DeserializeObject<Data>(reply);

            //when response is false check for the error message
            if (!captchaResponse.Success)
            {
                if (captchaResponse.ErrorCodes.Count <= 0)
                    return View();

                var error = captchaResponse.ErrorCodes[0].ToLower();
                switch (error)
                {
                    case ("missing-input-secret"):
                        ViewBag.Message = "The secret parameter is missing.";
                        break;
                    case ("invalid-input-secret"):
                        ViewBag.Message = "The secret parameter is invalid or malformed.";
                        break;

                    case ("missing-input-response"):
                        ViewBag.Message = "The response parameter is missing.";
                        break;
                    case ("invalid-input-response"):
                        ViewBag.Message = "The response parameter is invalid or malformed.";
                        break;

                    default:
                        ViewBag.Message = "Error occured. Please try again";
                        break;
                }
            }
            else
            {
                ViewBag.Message = "Valid";
                return RedirectToAction("NewIndex", "Home");

            }
            return View();
        }

        [HttpPost]
        public async Task<ActionResult> ExportToExcel(string cid)
        {
       
            List<Data> EmpInfo = new List<Data>();

            using (var client = new HttpClient())
            {
                //Passing service base url  
                client.BaseAddress = new Uri(Baseurl);

                client.DefaultRequestHeaders.Clear();
                //Define request data format  
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                //Sending request to find web api REST service resource GetAllEmployees using HttpClient  
                HttpResponseMessage Res = await client.GetAsync("posts");

                //Checking the response is successful or not which is sent using HttpClient  
                if (Res.IsSuccessStatusCode)
                {
                    //Storing the response details recieved from web api   
                    var EmpResponse = Res.Content.ReadAsStringAsync().Result;

                    //Deserializing the response recieved from web api and storing into the Employee list  
                    EmpInfo = JsonConvert.DeserializeObject<List<Data>>(EmpResponse);

                }
               
            }

            DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[4] { new DataColumn("userId"),
                                            new DataColumn("id"),
                                            new DataColumn("title"),
                                            new DataColumn("body") });

            
            foreach (var customer in EmpInfo)
            {
                dt.Rows.Add(customer.userId, customer.id, customer.title, customer.body);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "posts.xlsx");
                    
                }
            }
            
        }
        [HttpPost]
        public async Task<ActionResult> ExportSelToExcel(FormCollection frm)
        {
            string g = frm["gat"].ToString();
            string[] glist = g.Split(',');
            
            List<Data> EmpInfo = new List<Data>();
           
            using (var client = new HttpClient())
            {
                //Passing service base url  
                client.BaseAddress = new Uri(Baseurl);

                client.DefaultRequestHeaders.Clear();
                //Define request data format  
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                //Sending request to find web api REST service resource GetAllEmployees using HttpClient  
                HttpResponseMessage Res = await client.GetAsync("posts");

                //Checking the response is successful or not which is sent using HttpClient  
                if (Res.IsSuccessStatusCode)
                {
                    //Storing the response details recieved from web api   
                    var EmpResponse = Res.Content.ReadAsStringAsync().Result;

                    //Deserializing the response recieved from web api and storing into the Employee list  
                    EmpInfo = JsonConvert.DeserializeObject<List<Data>>(EmpResponse);

                }

            }

            DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[4] { new DataColumn("userId"),
                                            new DataColumn("id"),
                                            new DataColumn("title"),
                                            new DataColumn("body") });


            foreach (var it in glist)
            {
                foreach (var customer in EmpInfo)
                {
                    if (Convert.ToInt32(it) == customer.id)
                    {
                        dt.Rows.Add(customer.userId, customer.id, customer.title, customer.body);
                    }
                   
                }
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Selposts.xlsx");

                }
            }

        }
        string Baseurl = "https://jsonplaceholder.typicode.com/";
        public async Task<ActionResult> posts()
        {
            List<Data> EmpInfo = new List<Data>();

            using (var client = new HttpClient())
            {
                //Passing service base url  
                client.BaseAddress = new Uri(Baseurl);

                client.DefaultRequestHeaders.Clear();
                //Define request data format  
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                //Sending request to find web api REST service resource GetAllEmployees using HttpClient  
                HttpResponseMessage Res = await client.GetAsync("posts");

                //Checking the response is successful or not which is sent using HttpClient  
                if (Res.IsSuccessStatusCode)
                {
                    //Storing the response details recieved from web api   
                    var EmpResponse = Res.Content.ReadAsStringAsync().Result;

                    //Deserializing the response recieved from web api and storing into the Employee list  
                    EmpInfo = JsonConvert.DeserializeObject<List<Data>>(EmpResponse);


                }
                //returning the employee list to view  
                return View(EmpInfo);
            }
        }
        public ActionResult PostId()
        {
            return View();
        }
        
        public async Task<ActionResult> postsid(FormCollection frm)
        {
            int idd = Convert.ToInt32(frm["id"]);
            Data EmpInfo = new Data();

            using (var client = new HttpClient())
            {
                //Passing service base url  
                client.BaseAddress = new Uri(Baseurl);

                client.DefaultRequestHeaders.Clear();
                //Define request data format  
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                //Sending request to find web api REST service resource GetAllEmployees using HttpClient  
                HttpResponseMessage Res = await client.GetAsync("posts/"+idd);

                //Checking the response is successful or not which is sent using HttpClient  
                if (Res.IsSuccessStatusCode)
                {
                    //Storing the response details recieved from web api   
                    var EmpResponse = Res.Content.ReadAsStringAsync().Result;

                    //Deserializing the response recieved from web api and storing into the Employee list  
                    EmpInfo = JsonConvert.DeserializeObject<Data>(EmpResponse);

                }
                int code = (int)Res.StatusCode;
                ViewBag.code = code;
                //returning the employee list to view  
                return View(EmpInfo);
            }
        }

        public ActionResult PostIdFetch()
        {
            return View();
        }

        public async Task<ActionResult> postsidfetch(FormCollection frm)
        {
            int idd = Convert.ToInt32(frm["id"]);
           List<Data> EmpInfo = new List<Data>();

            using (var client = new HttpClient())
            {
                //Passing service base url  
                client.BaseAddress = new Uri(Baseurl);

                client.DefaultRequestHeaders.Clear();
                //Define request data format  
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                //Sending request to find web api REST service resource GetAllEmployees using HttpClient  
                HttpResponseMessage Res = await client.GetAsync("posts/" + idd+"/comments");

                //Checking the response is successful or not which is sent using HttpClient  
                if (Res.IsSuccessStatusCode)
                {
                    //Storing the response details recieved from web api   
                    var EmpResponse = Res.Content.ReadAsStringAsync().Result;

                    //Deserializing the response recieved from web api and storing into the Employee list  
                    EmpInfo = JsonConvert.DeserializeObject<List<Data>>(EmpResponse);
                    int code = (int)EmpInfo.Count;
                    ViewBag.code = code;

                }
                
                //returning the employee list to view  
                return View(EmpInfo);
            }
        }

        public ActionResult create()
        {
            int cc = Convert.ToInt32(Session["Sess"]);
            Session["Ses"] = cc;
            return View();
        }

        [HttpPost]
        public ActionResult create(Data student)
        {
            
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri("https://jsonplaceholder.typicode.com/posts");

                    //HTTP POST
                    var postTask = client.PostAsJsonAsync<Data>("posts", student);
                    postTask.Wait();

                    var result = postTask.Result;
                    int code = (int)result.StatusCode;
                    
                    if (result.IsSuccessStatusCode)
                    {
                        Session["Sess"] = code;
                        return RedirectToAction("Create");
                    }
                }
           

            return View();
        }

        public ActionResult delete()
        {
            int cc = Convert.ToInt32(Session["Sesss"]);
            Session["Se"] = cc;
            return View();
        }

        public ActionResult deletepost(FormCollection frm)
        {
            int id = Convert.ToInt32(frm["id"]);
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri("https://jsonplaceholder.typicode.com/");

                //HTTP DELETE
                var deleteTask = client.DeleteAsync("posts/" + id);
                deleteTask.Wait();

                var result = deleteTask.Result;
                int code = (int)result.StatusCode;
                if (result.IsSuccessStatusCode)
                {
                    Session["Sesss"] = code;
                    return RedirectToAction("delete");
                }
            }

            return RedirectToAction("delete");
        }
    }
}