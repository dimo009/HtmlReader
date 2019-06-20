using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace HtmlReader
{
    class Program
    {
        static void Main(string[] args)
        {
            string html = string.Empty;
            string url = "http://lists.suse.com/pipermail/sle-security-updates/2019-April/005274.html";

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            using (Stream stream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(stream))
            {


                html = reader.ReadToEnd();
                html.ToString();
                Console.WriteLine(html);
  
            }

            //Console.WriteLine(html);

            //Task t = new Task(DownloadPageAsync);
            //t.Start();
            //Console.WriteLine("Downloading page...");
            //Console.ReadLine();
        }

        static async void DownloadPageAsync()
        {
            string page = "http://lists.suse.com/pipermail/sle-security-updates/2019-April/005274.html";

            using (HttpClient client = new HttpClient())
            using (HttpResponseMessage response = await client.GetAsync(page))
            using (HttpContent content = response.Content)
            {
                
                string result = await content.ReadAsStringAsync();
                
                Console.WriteLine(result);
            }
        }
    }
}
