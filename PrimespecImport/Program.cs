using Newtonsoft.Json.Linq;
using RestSharp;
using System.Collections.Generic;
using Newtonsoft.Json;
using System;
using CsvHelper;
using System.IO;
using System.Globalization;
using System.Threading;

namespace PrimespecImport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Set up a timer to trigger every 15 minutes
            // Timer timer = new Timer(Run, null, TimeSpan.Zero, TimeSpan.FromMinutes(15));
            Run();
            RunDescriptions();
        }

        static void Run()
        {
            try
            {
                Console.WriteLine($"Program started at {DateTime.Now}.");

                List<Products> primespecProducts = getInventory();

                if (primespecProducts != null && primespecProducts.Count > 0)
                {
                    // Print each product to the console
                    foreach (var product in primespecProducts)
                    {
                        Console.WriteLine($"Product: {product.ManufacturePartNumber}, {product.ItemDescription}, {product.Cost}, {product.MSRP}");
                    }

                    // Write products to a CSV file
                    string fileName = "products.csv";
                    using (var writer = new StreamWriter(fileName))
                    using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                    {
                        csv.WriteRecords(primespecProducts);
                    }

                    Console.WriteLine($"Products downloaded to {fileName}");
                }
                else
                {
                    Console.WriteLine("No products retrieved.");
                }

                Console.WriteLine($"Capture products ended at {DateTime.Now}.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
        static void RunDescriptions()
        {
            try
            {
                Console.WriteLine($"Program started at {DateTime.Now}.");

                List<Descriptions> primespecProducts = getDescription();

                if (primespecProducts != null && primespecProducts.Count > 0)
                {
                    // Print each product to the console
                    foreach (var product in primespecProducts)
                    {
                        Console.WriteLine($"Product: {product.PartNo}, {product.ExtendedDescription}");
                    }

                    // Write products to a CSV file
                    string fileName = "descriptions.csv";
                    using (var writer = new StreamWriter(fileName))
                    using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                    {
                        csv.WriteRecords(primespecProducts);
                    }

                    Console.WriteLine($"Descriptions downloaded to {fileName}");
                }
                else
                {
                    Console.WriteLine("No products retrieved.");
                }

                Console.WriteLine($"Capture descriptions ended at {DateTime.Now}.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
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
        public class Descriptions
        {
            [JsonProperty("part_no")]
            public string PartNo { get; set; }

            [JsonProperty("extended_description")]
            public string ExtendedDescription { get; set; }
           
        }

            public class PrimespecDesc
        {
            public List<Descriptions> data { get; set; }
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

        //https://analyticsapi.zoho.com/api/phil@primespec.com/PRIMESPEC/level 9 description %23aR11-3

        public static List<Descriptions> getDescription()
        {
            var client = new RestClient("https://analyticsapi.zoho.com/api/phil@primespec.com/PRIMESPEC/level 9 description?ZOHO_ACTION=EXPORT&ZOHO_API_VERSION=1.0&KEY_VALUE_FORMAT=true&ZOHO_ERROR_FORMAT=JSON&ZOHO_OUTPUT_FORMAT=JSON");
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);
            request.AddHeader("ZANALYTICS-ORGID", "638592711");
            request.AddHeader("Authorization", "Bearer " + getAccessToken());
            IRestResponse response = client.Execute(request);
            PrimespecDesc myDeserializedClass = JsonConvert.DeserializeObject<PrimespecDesc>(response.Content);
            return myDeserializedClass.data;
        }
    }
}
