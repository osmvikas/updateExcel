using RestSharp;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using WebApplication1.Models;

namespace WebApplication1
{
    public partial class index : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            var client = new RestClient("http://localhost:44380/api/UpdateExcelSheet/Writexcel");
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);
            request.AlwaysMultipartFormData = false;
            request.AddParameter("NameAddress", "Vikas Sharma, New Delhi");
            request.AddParameter("PAN", "DTPPS2956N");
            request.AddParameter("FinancialYear", "20-21");
            request.AddParameter("Place", "New Delhi");
            request.AddParameter("Date", "22-05-2020");
            request.AddParameter("Designation", "Dot Net Developer");
            IRestResponse response = client.Execute(request);
            Response.Write(response.Content);
        }
    }
}