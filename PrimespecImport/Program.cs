using Newtonsoft.Json.Linq;
using RestSharp;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace PrimespecImport
{
    class Program
    {
        static void Main(string[] args)
        {

            /// Frequency limit: 5 times per 10 minutes, if you have to run it at a greater frequency, store the access token and run access token and feed jobs requests seperately. Access token expires every hour
            /// individual product structure can be found in the products class down below
            List<Products> primespecProducts = getInventory();
        }

        /// <summary>
        /// This function gets the token required for getting the inventory feeed
        /// </summary>
        /// <returns>Inventory Feeed API Token</returns>
        public static string getAccessToken()
        {
            var client = new RestClient("https://accounts.zoho.com/oauth/v2/token?refresh_token=1000.12c64a6fc52bb4bf38958ecc295b7e0e.ff590fd3e3b86ac2703ae741ac5e1c06&client_id=1000.COXSO8LCSDU6J3QX16XME3Z0ZWVLJU&client_secret=b893969bdbf273a3420d07b0ff80c36327ca0ca389&grant_type=refresh_token");
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);
            IRestResponse response = client.Execute(request);
            JObject json = JObject.Parse(response.Content);
            return json.GetValue("access_token").ToString();
        }

        /// <summary>
        /// Gets primespec inventory feed from Zoho converts it to a Product List Object
        /// You can change parmeter ZOHO_OUTPUT_FORMAT value to XML to get the feed in XML
        /// </summary>
        /// <returns>Products List</returns>
        public static List<Products> getInventory()
        {
            var client = new RestClient("https://analyticsapi.zoho.com/api/phil@primespec.com/PRIMESPEC/Price List For eSt@ff System?ZOHO_ACTION=EXPORT&ZOHO_API_VERSION=1.0&KEY_VALUE_FORMAT=true&ZOHO_ERROR_FORMAT=JSON&ZOHO_OUTPUT_FORMAT=JSON");
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);
            request.AddHeader("ZANALYTICS-ORGID", "638592711");
            request.AddHeader("Authorization", "Bearer " + getAccessToken());
            IRestResponse response = client.Execute(request);
            PrimespecProducts myDeserializedClass = JsonConvert.DeserializeObject<PrimespecProducts>(response.Content);
            return myDeserializedClass.data;
        }

        public class Products
        {
            [JsonProperty("Manufacture Part Number")]
            public string ManufacturePartNumber { get; set; }

            [JsonProperty("Item Description")]
            public string ItemDescription { get; set; }
            public string Cost { get; set; }
            public string MSRP { get; set; }
            public string UPC { get; set; }

            [JsonProperty("Quantity In Stock")]
            public string QuantityInStock { get; set; }
            public string Weight { get; set; }
            public string Category { get; set; }

            [JsonProperty("Manufacturer (UDF)")]
            public string ManufacturerUDF { get; set; }
            public string Image { get; set; }

            [JsonProperty("Price Level 9")]
            public string PriceLevel9 { get; set; }
        }

        public class PrimespecProducts
        {
            public List<Products> data { get; set; }
        }

    }
}
