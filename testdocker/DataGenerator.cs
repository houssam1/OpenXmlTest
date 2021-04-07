using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Runtime.Serialization.Json;
using System.Threading.Tasks;

namespace OpenXmlPocDocker
{
    public static class DataGenerator
    {
        
        public static async Task<List<Country>> GetCountries()
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Add("X-API-KEY", "f64ac212e83540589dfd505505970c46");
            Stream respStream = await client.GetStreamAsync("https://randommer.io/api/Misc/Cultures");
            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(List<Country>));
            List<Country> countries = (List<Country>) ser.ReadObject(respStream);
            return countries;
        }
    }
}
