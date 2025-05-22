using HtmlAgilityPack;
using System.Net;
using Azure.Storage.Blobs;
using Microsoft.Data.SqlClient;
using System.Xml;
using System.Security.Policy;
using System.Security.Cryptography;

namespace SitemapDownload
{
    class Program
    {
        static void Main(string[] args)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;  //https접속에러를 막는 프로토콜
            HtmlWeb web = new HtmlWeb();
            BlobServiceClient blobServiceClient = new BlobServiceClient("DefaultEndpointsProtocol=https;AccountName=leocomfile;AccountKey=9RzrYKzmdJNgwTh2gxJ+ioYs+/YmbHjY08ee8QizahIYh7W+XyECYrzrIwEBCJIgwqR8JR+69As8E+T7o/BktQ==;EndpointSuffix=core.windows.net");
            BlobContainerClient container1 = blobServiceClient.GetBlobContainerClient("image");

            String _connString = System.Configuration.ConfigurationManager.ConnectionStrings["sam"].ToString();
            using (SqlConnection cn = new SqlConnection(_connString))
            {
                cn.Open();

                //string weblink = "https://www.lcsc.com/sitemap-2.xml";
                //HtmlDocument doc = web.Load(weblink);

                int i = 2;

                while (i <= 48)
                {
                    WebClient client = new WebClient();
                    client.Headers["User-Agent"] = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.1 (KHTML, like Gecko) Chrome/14.0.835.202 Safari/535.1";
                    client.Headers["Accept"] = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                    string data = client.DownloadString("https://www.lcsc.com/sitemap-"+i+".xml");
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(data);

                    var urls = doc.GetElementsByTagName("loc");

                    foreach (XmlElement loc in urls)
                    {
                        var weburl = loc.InnerText;
                        var partnumber = weburl.Substring(weburl.LastIndexOf("_") + 1, weburl.Length - weburl.LastIndexOf("_") - 6);

                        SqlCommand Cmd = new SqlCommand("", cn);
                        Cmd.CommandText = " Insert into LCSC_master Values ('" + partnumber + "', NULL, NULL, NULL, NULL, NULL,NULL,NULL, '" + weburl + "', NULL, NULL, NULL, NULL, NULL, NULL)";
                        //Cmd.CommandText = "INSERT INTO LCSC_master (LCSCPartNo, WebURL) SELECT '" + partnumber + "', '" + weburl + "' FROM DUAL WHERE NOT EXISTS (SELECT * FROM LCSC_master WHERE LCSCPartNo='" + partnumber + "' AND WebURL='" + weburl + "' LIMIT 1)";
                        Cmd.ExecuteNonQuery();

                    }
                    System.Threading.Thread.Sleep(1000);     //시간 지연 1초

                    i++;
                }
                cn.Close();







            }


        }
    }
}