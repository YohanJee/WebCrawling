using HtmlAgilityPack;
using System.Net;
using IronXL;



namespace Takachi
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            HtmlWeb web = new HtmlWeb();
            web.UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.1 (KHTML, like Gecko) Chrome/14.0.835.202 Safari/535.1";
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");
            int o = 1;
            try
            {
                var web0 = "https://www.takachi-enclosure.com/";
                HtmlDocument doc0 = web.Load(web0);

                List<string> list = new List<string>();
                var count = doc0.DocumentNode.SelectNodes("//ul[@class='flex-wrap col-4']/li").Count;
                for (int i = 1; i <= count; i++)
                {
                    list.Add(doc0.DocumentNode.SelectSingleNode("//ul[@class='flex-wrap col-4']/li[" + i + "]/a[1]").Attributes["href"].Value);
                }

                var y = 1;

                while (y <= count)
                {
                    string weblink1 = list[y - 1];

                    workSheet["A"+o].Value = weblink1;

                    HtmlDocument doc1 = web.Load(weblink1);

                    try

                    {
                        var categorycount = doc1.DocumentNode.SelectNodes("//section[@class='products-list2 products-list2-en']/div[1]/div").Count;
                        var u = 1;
                        while (u <= categorycount)
                        {
                            var weblink5 = doc1.DocumentNode.SelectSingleNode("//section[@class='products-list2 products-list2-en']/div[1]/div[" + u + "]/figure[1]/a[1]").Attributes["href"].Value;
                            if(weblink5.StartsWith("/"))
                            {
                                weblink5 = "https://www.takachi-enclosure.com" + weblink5;
                            }
                            workSheet["B"+o].Value = weblink5;
                            o++;
                            u++;
                        }
                    }
                    catch
                    {
                        var categorycount1 = doc1.DocumentNode.SelectNodes("//div[@class='product-row-list']/article").Count;
                        var q = 1;
                        while (q <= categorycount1)
                        {
                            var weblink2 = doc1.DocumentNode.SelectSingleNode("//div[@class='product-row-list']/article[" + q + "]/div[1]/a[1]").Attributes["href"].Value;
                            if (weblink2.StartsWith("/"))
                            {
                                weblink2 = "https://www.takachi-enclosure.com" + weblink2;
                            }
                            workSheet["b"+o].Value = weblink2;

                            HtmlDocument doc2 = web.Load(weblink2);
                            try
                            {
                                var categorycount2 = doc2.DocumentNode.SelectNodes("//section[@class='products-list2 products-list2-en']/div[1]/div").Count;
                                var t = 1;
                                while (t <= categorycount2)
                                {
                                    var weblink3 = doc2.DocumentNode.SelectSingleNode("//section[@class='products-list2 products-list2-en']/div[1]/div[" + t + "]/figure[1]/a[1]").Attributes["href"].Value;
                                    if (weblink3.StartsWith("/"))
                                    {
                                        weblink3 = "https://www.takachi-enclosure.com" + weblink3;
                                    }
                                    workSheet["C"+o].Value = weblink3;
                                    o++;
                                    t++;

                                }

                            }
                            catch
                            {
                                var categorycount3 = doc2.DocumentNode.SelectNodes("//div[@class='product-row-list']/article").Count;
                                var e = 1;

                                while (e <= categorycount3)
                                {
                                    var weblink4 = doc2.DocumentNode.SelectSingleNode("//div[@class='product-row-list']/article[" + e + "]/div[1]/a[1]").Attributes["href"].Value;
                                    if (weblink4.StartsWith("/"))
                                    {
                                        weblink4 = "https://www.takachi-enclosure.com" + weblink4;
                                    }
                                    workSheet["C"+o].Value = weblink4;

                                    HtmlDocument doc3 = web.Load(weblink4);

                                    try
                                    {

                                        var categorycount4 = doc3.DocumentNode.SelectNodes("//section[@class='products-list2 products-list2-en']/div[1]/div").Count;
                                        var r = 1;
                                        while (r <= categorycount4)
                                        {
                                            var weblink5 = doc3.DocumentNode.SelectSingleNode("//section[@class='products-list2 products-list2-en']/div[1]/div[" + r + "]/figure[1]/a[1]").Attributes["href"].Value;
                                            if (weblink5.StartsWith("/"))
                                            {
                                                weblink5 = "https://www.takachi-enclosure.com" + weblink5;
                                            }
                                            workSheet["D"+o].Value = weblink5;
                                            o++;

                                            r++;
                                        }
                                    }
                                    catch
                                    {

                                    }
                                    e++;
                                }
                            }

                            q++;
                        }

                    }
                    y++;
                }




                }
                catch(Exception ex)
                {
                    Console.WriteLine("exmsg=" + ex.Message);
                }

            workBook.SaveAs("OfficialTakachiURL.xlsx");


        }










    }
}
