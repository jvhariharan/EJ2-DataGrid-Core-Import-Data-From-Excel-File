using TestSample.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.EJ2;
using System;
using System.Linq;

namespace TestSample.Controllers
{
    [Produces("application/json")]
    [Route("api/Orders")]
    public class OrdersController : Controller
    {
        // GET: api/Orders
        [HttpGet("GetTest")]
        public object GetTest()
        {

            var queryString = Request.Query;
            int skip = Convert.ToInt32(queryString["$skip"]);
            int take = Convert.ToInt32(queryString["$top"]);
            var data = OrdersDetails.GetAllRecords().ToList();
            return Json(new { Items = data.Skip(skip).Take(take), Count = data.Count() });

        }

        // GET: api/Orders/5
        [HttpGet("{id}", Name = "Get")]
        public object Get(string id)
        {
            var queryString = Request.Query;
            var dataa = Convert.ToString(queryString["id"]);
            var data = OrdersDetails.GetAllRecords().Where(user => user.CustomerID == dataa).ToList();
            return new { Items = data, Count = data.Count() };
        }



        [HttpPost("GetTest")]
        public object PostTest([FromBody]OrdersDetails value)
        {
            OrdersDetails.GetAllRecords().Add(value);
            var Data = OrdersDetails.GetAllRecords().ToList();
            int count = Data.Count();
            return Json(new { result = Data, count = count });
        }


        // PUT: api/Orders/5
        [HttpPut("GetTest")]
        public object Put([FromBody]OrdersDetails value)
        {
            var ord = value;
            OrdersDetails val = OrdersDetails.GetAllRecords().Where(or => or.OrderID == ord.OrderID).FirstOrDefault();
            val.OrderID = ord.OrderID;
            val.EmployeeID = ord.EmployeeID;
            val.CustomerID = ord.CustomerID;
            val.Freight = ord.Freight;
            val.OrderDate = ord.OrderDate;
            val.ShipCity = ord.ShipCity;
            return value;
        }

        // DELETE: api/ApiWithActions/5
        [HttpDelete("{id:int}")]
        [Route("Orders/{id:int}")]
        public object Delete(int id)
        {
            OrdersDetails.GetAllRecords().Remove(OrdersDetails.GetAllRecords().Where(or => or.OrderID == id).FirstOrDefault());
            return Json(id);
        }
    }
    public class Data
    {

        public bool requiresCounts { get; set; }
        public int skip { get; set; }
        public int take { get; set; }
    }

}