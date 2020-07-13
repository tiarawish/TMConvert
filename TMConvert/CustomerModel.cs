using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMConvert
{
    public class CustomerModel
    {
        public int id { get; set; }
        public string company { get; set; }
        public string customerNumber { get; set; }
        public List<CustomPrice> lCustomerPrices;
    }

    public class resultListCM
    {
        public List<CustomerModel> result { get; set; }
    }
    public class resultListCP
    {
        public List<CustomPrice> result { get; set; }
    }

    public class CustomPrice
    {
        public int articleId { get; set; }
        public string articleNumber { get; set; }
        public decimal price { get; set; }
        public string endDate { get; set; }
        public string description { get; set; }
    }

}
