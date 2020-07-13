using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace TMConvert
{
    public class ApiProcessor
    {

        public async Task<CustomerModel> loadCustomer(int customerID)
        {
            string url = "https://szgwdsfnutmhvnz.weclapp.com/webapp/api/v1/customer/id/" + customerID;

            using (HttpResponseMessage response = await ApiHelper.ApiClient.GetAsync(url))
            {
                if (response.IsSuccessStatusCode)
                {
                    CustomerModel customer = await response.Content.ReadAsAsync<CustomerModel>();
                    return customer;
                }
                else
                {
                    throw new Exception(response.ReasonPhrase);
                }
            }
        }

        public async Task<resultListCP> getCustomPricesForCustomer(int customerId)
        {
            string url = "https://szgwdsfnutmhvnz.weclapp.com/webapp/api/v1/articlePrice?customerId-eq=" + customerId.ToString() + "&endDate-null&pageSize=1000";

            using (HttpResponseMessage response = await ApiHelper.ApiClient.GetAsync(url))
            {
                if (response.IsSuccessStatusCode)
                {
                    string result = await response.Content.ReadAsStringAsync();

                    resultListCP listCP = JsonConvert.DeserializeObject<resultListCP>(result);
                    return listCP;
                }
                else
                {
                    throw new Exception(response.ReasonPhrase);
                }
            }
        }
    

        public async Task<resultListCP> getCustomersWithCustomPricesList()
        {
            string url = "https://szgwdsfnutmhvnz.weclapp.com/webapp/api/v1/articlePrice?customerId-notnull&endDate-null&properties=customerId,endDate&page=2&pageSize=1000"; //&properties=articleId,articleNumber,customerId,price

            using (HttpResponseMessage response = await ApiHelper.ApiClient.GetAsync(url))
            {
                if (response.IsSuccessStatusCode)
                {
                    string result = await response.Content.ReadAsStringAsync();

                    resultListCP customersWithCustomPricesList = JsonConvert.DeserializeObject<resultListCP>(result);

                    return customersWithCustomPricesList;
                }
                else
                {
                    throw new Exception(response.ReasonPhrase);
                }
            }
            
        }

        public async Task<resultListCM> getCustomers()
        {
            string url = "https://szgwdsfnutmhvnz.weclapp.com/webapp/api/v1/customer?customerNumber-notnoll&properties=id,company,customerNumber&pageSize=1000";

            using (HttpResponseMessage response = await ApiHelper.ApiClient.GetAsync(url))
            {
                if (response.IsSuccessStatusCode)
                {
                    string customers = await response.Content.ReadAsStringAsync();

                    resultListCM result = JsonConvert.DeserializeObject<resultListCM>(customers);
                    return result;
                }
                else
                {
                    throw new Exception(response.ReasonPhrase);
                }
            }
        }
    }
}
