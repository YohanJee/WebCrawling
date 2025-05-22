using HtmlAgilityPack;
using System.Net;
using System.Xml;
using System.Security.Policy;
using System.Security.Cryptography;
using static System.Net.Mime.MediaTypeNames;
using System.Data.SqlClient;

namespace TestingWeb
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            HtmlWeb web = new HtmlWeb();

            String _connString = System.Configuration.ConfigurationManager.ConnectionStrings["sam"].ToString();
            using (SqlConnection cn = new SqlConnection(_connString))
            {
                try
                {
                    cn.Open();
                    SqlCommand Cmd = new SqlCommand("", cn);
                    SqlDataReader dr;
                    String sql = "select LCSCPartNo, Description, WebURL from LCSC_master where Description IS NULL";
                    Cmd.CommandText = sql;
                    dr = Cmd.ExecuteReader();

                    if (dr.HasRows)
                    {
                        int xx = 1;
                        while (dr.Read())
                        {
                            //string weblink = "https://www.lcsc.com/product-detail/Multilayer-Ceramic-Capacitors-MLCC-Leaded_KEMET_C1017252.html";

                            string weblink = dr["WebURL"].ToString();

                            try
                            {


                                HtmlDocument doc = web.Load(weblink);

                                //var search = doc.DocumentNode.SelectSingleNode("//table[@class='info-table']");

                                var mytable = new Dictionary<string, string>();

                                //var test = doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/main[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/table[1]");

                                int count = doc.DocumentNode.SelectNodes("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/main[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/table[1]/tbody[1]/tr").Count;

                                int i = 1;

                                while (i <= count)
                                {
                                    var key = doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/main[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/table[1]/tbody[1]/tr[" + i + "]/td[1]").InnerHtml.Replace("\n", "").Replace("\t", "");
                                    var value = doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/main[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/table[1]/tbody[1]/tr[" + i + "]/td[2]").InnerHtml.Replace("\n", "").Replace("\t", "").Replace("  ", "");
                                    if (key == "Datasheet")
                                    {
                                        value = doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/main[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/table[1]/tbody[1]/tr[" + i + "]/td[2]/a[1]").Attributes["href"].Value;
                                    }
                                    if (key == "Manufacturer")
                                    {
                                        value = doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/main[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/table[1]/tbody[1]/tr[" + i + "]/td[2]/a[1]").InnerHtml.Replace("\n", "").Replace("\t", "").Replace("  ", "");
                                    }
                                    mytable.Add(key, value);
                                    i++;
                                }

                                if (mytable.ContainsKey("Customer #"))
                                {
                                    mytable.Remove("Customer #");
                                }
                                if (mytable.ContainsKey("EasyEDA"))
                                {
                                    mytable.Remove("EasyEDA");
                                }
                                if (mytable.ContainsKey("Datasheet"))
                                {

                                }
                                else
                                {
                                    mytable.Add("Datasheet", "");
                                }

                                //imagelink 사진 없는 제품도 존재
                                try
                                {
                                    var image = doc.DocumentNode.SelectSingleNode("//div[@class='v-slide-group__content']/div[1]/img[1]");
                                    var imgHref = image.Attributes["src"].Value;
                                    mytable.Add("imglink", imgHref);
                                }
                                catch
                                {
                                    mytable.Add("imglink", "");
                                }



                                //instock
                                var stock = doc.DocumentNode.SelectSingleNode("//div[@class='right']/div[1]/div[1]/div[1]").InnerText;
                                var stocknumber = stock.Substring(stock.IndexOf(":") + 2).Replace("\n", "").Replace(" ", "");
                                if (stocknumber == "GetaQuote")
                                {
                                    stocknumber = "Need to get a quote";
                                }
                                mytable.Add("In Stock", stocknumber);

                                //category
                                var cat = doc.DocumentNode.SelectSingleNode("//table[@class='products-specifications']/tbody[1]/tr[1]/td[2]/a[1]").InnerText.Replace("\n", "").Replace("  ", "");
                                mytable.Add("Category", cat);

                                //MinimumOrderQty, MultipleQty, Packaging, SalesUnit, StandardPacking
                                var BuyTips = doc.DocumentNode.SelectSingleNode("//div[@class='buy-tips']").InnerText.Replace("\n", "").Replace("  ", "");
                                string[] arr1 = BuyTips.Split(' ');
                                List<string> list1 = new List<string>();
                                list1 = arr1.ToList();
                                mytable.Add("MinimumOrderQty", list1[2]);
                                mytable.Add("MultipleQty", list1[5]);

                                var SalesUnit = doc.DocumentNode.SelectSingleNode("//div[@class='pb-4']/span[1]").InnerText.Replace("\n", "").Replace("  ", "");
                                SalesUnit = SalesUnit.Substring(SalesUnit.LastIndexOf(":") + 2);
                                mytable.Add("SalesUnit", SalesUnit);

                                var StandardPacking = doc.DocumentNode.SelectSingleNode("//span[@class='ml-4']").InnerText.Replace("\n", "").Replace("  ", "");
                                mytable.Add("StandardPacking", StandardPacking);






                                //저장
                                SqlCommand Cmd1 = new SqlCommand("", cn);
                                Cmd1.CommandText = "update LCSC_master set MinimumOrderQty = '" + mytable["MinimumOrderQty"] + "', MultipleQty = '" + mytable["MultipleQty"] + "', SalesUnit = '" + mytable["SalesUnit"] + "', StandardPacking = '" + mytable["StandardPacking"] + "', Description = '" + mytable["Description"] + "', MfrPartNo = '" + mytable["Mfr. Part #"] + "', Manufacturer = '" + mytable["Manufacturer"] + "', Package = '" + mytable["Package"] + "', StockCount = '" + mytable["In Stock"] + "', DatasheetLink = '" + mytable["Datasheet"] + "', ImgLink = '" + mytable["imglink"] + "', Category = '" + mytable["Category"] + "', ChangedDate = CONVERT(datetimeoffset, getdate()) AT TIME ZONE 'Korea Standard Time', ChangedBy = 'Yohan' where LCSCPartNo = '" + mytable["LCSC Part #"] + "'";
                                Cmd1.ExecuteNonQuery();


                                //price
                                var pricetable = new Dictionary<int, decimal>();

                                var search = doc.DocumentNode.SelectSingleNode("//td[@class='tbody-num ahover']");
                                int count3 = doc.DocumentNode.SelectNodes("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/main[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[4]/table[1]/tbody[1]/tr").Count;
                                int k = 1;
                                while (k <= count3)
                                {
                                    var key = Convert.ToInt16(doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/main[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[4]/table[1]/tbody[1]/tr[" + k + "]/td[1]").InnerHtml.Replace(" ", "").Replace("+", ""));
                                    var value = Convert.ToDecimal(doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/main[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[4]/table[1]/tbody[1]/tr[" + k + "]/td[2]/span[1]").InnerHtml.Replace(" ", "").Replace("US$", ""));
                                    pricetable.Add(key, value);
                                    k++;
                                }

                                //price 저장
                                foreach (var buycount in pricetable)
                                {
                                    SqlCommand Cmd2 = new SqlCommand("", cn);
                                    Cmd2.CommandText = "insert into LCSC_Price values ('" + mytable["LCSC Part #"] + "','" + buycount.Key + "', '" + buycount.Value + "')";
                                    Cmd2.ExecuteNonQuery();
                                }



                                //제품 specification
                                try
                                {
                                    var spectable = new Dictionary<string, string>();

                                    var test = doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/script[1]").InnerText;
                                    int o = 0;
                                    string[] arr = test.Split(',');
                                    List<string> list = new List<string>();
                                    list = arr.ToList();
                                    foreach (string s in list)
                                    {
                                        if (s.Contains("paramNameEn") && list[o + 1] != "paramValue:h" && list[o + 1] != "paramValue:g" && list[o + 1] != "paramValue:f" && list[o + 1] != "paramValue:c" && list[o + 1] != "paramValue:a")
                                        {
                                            if (list[o + 2].Length > 15)
                                            {
                                                spectable.Add(s.Substring(s.LastIndexOf(":") + 2, s.Length - s.LastIndexOf(":") - 3).Replace("\u002F", "/"), list[o + 2].Substring(list[o + 2].LastIndexOf(":") + 2, list[o + 2].Length - list[o + 2].LastIndexOf(":") - 3).Replace("\"", "").Replace(")", "").Replace(";", "").Replace("\u002F", "/"));

                                            }
                                            else
                                            {
                                                spectable.Add(s.Substring(s.LastIndexOf(":") + 2, s.Length - s.LastIndexOf(":") - 3).Replace("\u002F", "/"), "");
                                            }
                                        }
                                        o++;
                                    }


                                    List<string> keys = spectable.Keys.ToList();
                                    List<string> values = spectable.Values.ToList();

                                    int p = 0;

                                    var q = values.GroupBy(x => x).Select(g => new { Value = g.Key, Count = g.Count() }).OrderByDescending(x => x.Count);
                                    foreach (var x in q)
                                    {
                                        if (x.Value == "")
                                        {
                                            p = x.Count;
                                        }
                                    }

                                    int r = 1;
                                    int w = 0;
                                    while (r <= spectable.Count)
                                    {
                                        if (spectable[keys[r - 1]] == "")
                                        {
                                            spectable[keys[r - 1]] = list[o - p + w].Replace("\"", "").Replace(")", "").Replace(";", "").Replace("\u002F", "/");
                                            w++;

                                        }
                                        r++;

                                    }

                                    //specification 저장
                                    foreach (var parameter in spectable)
                                    {
                                        SqlCommand Cmd3 = new SqlCommand("", cn);
                                        Cmd3.CommandText = "insert into LCSC_Spec values ('" + mytable["LCSC Part #"] + "',N'" + parameter.Key + "', N'" + parameter.Value + "')";
                                        Cmd3.ExecuteNonQuery();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("exmsg=" + ex.Message);

                                }


                                System.Threading.Thread.Sleep(1000);     //시간 지연 15초


                            }
                            catch (Exception ex)
                            {

                                Console.WriteLine("exmsg=" + ex.Message);
                            }
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