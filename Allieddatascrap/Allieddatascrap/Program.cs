using HtmlAgilityPack;
using System.Data.SqlClient;
using System.Net;
using static System.Net.Mime.MediaTypeNames;
using System.Text.RegularExpressions;
using System.Reflection.Metadata;

namespace RSComponents
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            HtmlWeb web = new HtmlWeb();
            web.UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.1 (KHTML, like Gecko) Chrome/14.0.835.202 Safari/535.1";
            //BlobServiceClient blobServiceClient = new BlobServiceClient("DefaultEndpointsProtocol=https;AccountName=leocomfile;AccountKey=9RzrYKzmdJNgwTh2gxJ+ioYs+/YmbHjY08ee8QizahIYh7W+XyECYrzrIwEBCJIgwqR8JR+69As8E+T7o/BktQ==;EndpointSuffix=core.windows.net");
            //BlobContainerClient container1 = blobServiceClient.GetBlobContainerClient("image");

            String _connString = System.Configuration.ConfigurationManager.ConnectionStrings["sam"].ToString();
            using (SqlConnection cn = new SqlConnection(_connString))
            {
                try
                {
                    cn.Open();
                    SqlCommand Cmd = new SqlCommand("", cn);
                    SqlDataReader dr;
                    String sql = "select RSPartNo, Description, WebURL from Allied_master where Description IS NULL";
                    Cmd.CommandText = sql;
                    dr = Cmd.ExecuteReader();

                    if (dr.HasRows)
                    {
                        int xx = 1;
                        while (dr.Read())
                        {
                            //string weblink = "https://us.rs-online.com/product/eao/45-1819-1c90-003-108/70915158/";

                            string weblink = dr["WebURL"].ToString();

                            try
                            {

                                HtmlDocument doc = web.Load(weblink);

                                //var search = doc.DocumentNode.SelectSingleNode("//table[@class='info-table']");
                                //var test = doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/main[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/table[1]");

                                //RSPartno
                                string RSPartNo = "", interim = "";
                                int inter = 0, inter2 = 0;
                                try
                                {
                                    //RSPartNo = doc.DocumentNode.SelectSingleNode("//div[@class='identifiers']/p[1]/span[1]/span[1]").InnerText;
                                    interim = weblink.TrimEnd('/');
                                    inter = interim.LastIndexOf('/');
                                    inter2 = weblink.Length - interim.LastIndexOf('/')-2;
                                    RSPartNo = interim.Substring(inter+1, inter2);
                                }
                                catch
                                {

                                }
                                //description
                                var des = "";
                                try
                                {
                                    //des = doc.DocumentNode.SelectSingleNode("//div[@class='description']/p[1]").InnerText.Replace("\r", "");
                                    des = doc.DocumentNode.SelectSingleNode("//h2[@class='product-display-description']").InnerText.Replace("\r", "").Replace("\n", "").Trim();
                                }
                                catch
                                {

                                }
                                //mfrpartno
                                string mfrpartno = "";
                                try
                                {
                                    mfrpartno = doc.DocumentNode.SelectSingleNode("//strong[@class='manufacturer-part-number']/span[1]").InnerText;
                                }
                                catch
                                {

                                }

                                //manufacturer
                                string manufacturer = "";
                                try
                                {
                                    manufacturer = doc.DocumentNode.SelectSingleNode("//h1[@class='product-display-name OneLinkNoTx']").InnerText;
                                    manufacturer = manufacturer.Substring(0, manufacturer.Length - mfrpartno.Length-2).Replace("\r", "").Replace("\n", "").Trim();
                                }
                                catch
                                {

                                }
                                //stockcount
                                var stcount = "";
                                try
                                {
                                    stcount = doc.DocumentNode.SelectSingleNode("//div[@class='availability-value in-stock-message']/p[2]").InnerText;
                                }
                                catch
                                {

                                }
                                //DatasheetLink
                                var datasheetlink = "";
                                try
                                {
                                    var datasheet = doc.DocumentNode.SelectSingleNode("//p[@class='resource-item']/a[1]");
                                    datasheetlink = "https://us.rs-online.com" + datasheet.Attributes["href"].Value;
                                }
                                catch
                                {

                                }

                                //imglink
                                var imglink = "";
                                try
                                {
                                    var img = doc.DocumentNode.SelectSingleNode("//img[@class='spin-image-img']");
                                    imglink = img.Attributes["src"].Value;
                                }
                                catch
                                {
                                    try
                                    {
                                        var img = doc.DocumentNode.SelectSingleNode("//div[@class='image-gallery-slider ']/img[1]");
                                        imglink = img.Attributes["src"].Value;
                                    }
                                    catch
                                    {
                                        try
                                        {
                                            var img = doc.DocumentNode.SelectSingleNode("//div[@class='image-gallery-slider simple-carousel']/div[1]/img[1]");
                                            imglink = img.Attributes["src"].Value;
                                        }
                                        catch
                                        {

                                        }
                                    }
                                }
                                //category
                                var category = "";
                                try
                                {
                                    var categorycount = doc.DocumentNode.SelectNodes("//div[@class='material-navigation']/a").Count;
                                    category = doc.DocumentNode.SelectSingleNode("//div[@class='material-navigation']/a[" + categorycount + "]").InnerText;
                                }
                                catch
                                {

                                }

                                //minimum qty
                                var minqty = "";
                                try
                                {
                                    minqty = doc.DocumentNode.SelectSingleNode("//div[@class='min-mult']/div[1]").InnerText;
                                    minqty = minqty.Substring(minqty.LastIndexOf(":") + 1, minqty.Length - minqty.LastIndexOf(":") - 1).Replace("\r", "");
                                }
                                catch
                                {

                                }

                                //multiple qty
                                var multipleqty = "";
                                try
                                {
                                    multipleqty = doc.DocumentNode.SelectSingleNode("//div[@class='min-mult']/div[2]").InnerText;
                                    multipleqty = multipleqty.Substring(multipleqty.LastIndexOf(":") + 1, multipleqty.Length - multipleqty.LastIndexOf(":") - 1).Replace("\r", "");
                                }
                                catch
                                {

                                }

                                //salesunit
                                var salesunit = "";
                                try
                                {
                                    salesunit = doc.DocumentNode.SelectSingleNode("//div[@class='value-and-uom OneLinkNoTx']/span[1]").InnerText.Replace("/", "");
                                }
                                catch
                                {

                                }



                                //Standardpacking





                                //저장
                                SqlCommand Cmd1 = new SqlCommand("", cn);
                                Cmd1.CommandText = "update Allied_master set SalesUnit = '" + salesunit + "', MinimumOrderQty = '" + minqty + "', MultipleQty = '" + multipleqty + "',  Description = '" + des + "', MfrPartNo = '" + mfrpartno + "', Manufacturer = '" + manufacturer + "',  StockCount = '" + stcount + "', DatasheetLink = '" + datasheetlink + "', ImgLink = '" + imglink + "', Category = '" + category + "', ChangedDate = CONVERT(datetimeoffset, getdate()) AT TIME ZONE 'Korea Standard Time', ChangedBy = 'Yohan' where RSPartNo = '" + RSPartNo + "'";
                                Cmd1.ExecuteNonQuery();







                                //price
                                var pricetable = new Dictionary<int, decimal>();
                                try
                                {


                                    //var search = doc.DocumentNode.SelectSingleNode("//div[@class='new-material-available-standard-pricing-body']").InnerText;
                                    int count3 = doc.DocumentNode.SelectNodes("//div[@class='new-material-available-standard-pricing-body']/div").Count;
                                    int k = 2;
                                    while (k <= count3)
                                    {
                                        var key = Convert.ToInt16(doc.DocumentNode.SelectSingleNode("//div[@class='new-material-available-standard-pricing-body']/div[" + k + "]/div[1]/p[1]").InnerText);
                                        var value = Convert.ToDecimal(doc.DocumentNode.SelectSingleNode("//div[@class='new-material-available-standard-pricing-body']/div[" + k + "]/div[2]/p[1]").InnerText.Replace("$", ""));
                                        pricetable.Add(key, value);
                                        k++;
                                    }
                                }
                                catch
                                {

                                }
                                //price 저장
                                foreach (var buycount in pricetable)
                                {
                                    SqlCommand Cmd2 = new SqlCommand("", cn);
                                    Cmd2.CommandText = "insert into Allied_Price values ('" + RSPartNo + "','" + buycount.Key + "', '" + buycount.Value + "')";
                                    Cmd2.ExecuteNonQuery();
                                }



                                //packaging
                                var packaging = "";


                                //package type
                                var packagetype = "";


                                //제품 specification
                                var spectable = new Dictionary<string, string>();

                                //var search = doc.DocumentNode.SelectSingleNode("//div[@class='new-material-available-standard-pricing-body']").InnerText;
                                try
                                {


                                    int count = doc.DocumentNode.SelectNodes("//div[@class='product-specifications-body']/div").Count;
                                    int p = 2;
                                    while (p <= count)
                                    {
                                        try
                                        {
                                            var key = doc.DocumentNode.SelectSingleNode("//div[@class='product-specifications-body']/div[" + p + "]/div[1]/span[1]").InnerText.Replace("\'", "");
                                            var value = doc.DocumentNode.SelectSingleNode("//div[@class='product-specifications-body']/div[" + p + "]/div[2]/span[1]").InnerText.Replace("\'", "");
                                            spectable.Add(key, value);
                                            if (key == "Packaging")
                                            {
                                                packaging = value;
                                            }
                                            if (key == "Package Type")
                                            {
                                                packagetype = value;
                                            }
                                        }
                                        catch
                                        {
                                            var key = doc.DocumentNode.SelectSingleNode("//div[@class='product-specifications-body']/div[" + p + "]/div[1]/span[1]").InnerText.Replace("\'", "");
                                            var value = doc.DocumentNode.SelectSingleNode("//div[@class='product-specifications-body']/div[" + p + "]/div[2]/a[1]/span[1]").InnerText.Replace("\'", "");
                                            spectable.Add(key, value);
                                            if (key == "Packaging")
                                            {
                                                packaging = value;
                                            }
                                            if (key == "Package Type")
                                            {
                                                packagetype = value;
                                            }
                                        }


                                        p++;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("exmsg=" + ex.Message);
                                }
                                //spec 저장
                                foreach (var speccount in spectable)
                                {
                                    SqlCommand Cmd2 = new SqlCommand("", cn);
                                    Cmd2.CommandText = "insert into Allied_Spec values ('" + RSPartNo + "',N'" + speccount.Key + "', N'" + speccount.Value + "')";
                                    Cmd2.ExecuteNonQuery();
                                }

                                SqlCommand Cmd4 = new SqlCommand("", cn);
                                Cmd4.CommandText = "update Allied_master set Packaging = '" + packaging + "', Package = '" + packagetype + "' where RSPartNo = '" + RSPartNo + "'";
                                Cmd4.ExecuteNonQuery();




                            }
                            catch (Exception ex)
                            {

                                Console.WriteLine("exmsg=" + ex.Message);
                            }
                            System.Threading.Thread.Sleep(1000);     //시간 지연 15초

                        }
                        xx++;
                    }
                }

                catch (Exception ex)
                {

                    Console.WriteLine("exmsg=" + ex.Message);
                }

            }

        }
    }
}