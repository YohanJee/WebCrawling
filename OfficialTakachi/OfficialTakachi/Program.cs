using HtmlAgilityPack;
using System.Data.SqlClient;
using System.Net;



namespace Takachi
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            HtmlWeb web = new HtmlWeb();
            web.UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.1 (KHTML, like Gecko) Chrome/14.0.835.202 Safari/535.1";
            String _connString = System.Configuration.ConfigurationManager.ConnectionStrings["sam"].ToString();
            using (SqlConnection cn = new SqlConnection(_connString))
            {
                try
                {
                    cn.Open();
                    SqlCommand Cmd = new SqlCommand("", cn);
                    //i 값을 mssql에 ID로 사용

                    int id = 1;

                    SqlCommand mid = new SqlCommand("", cn);
                    mid.CommandText = "select TOP 1 * from OfficialTakachiCategory order by ID desc";
                    SqlDataReader dt = mid.ExecuteReader();
                    if (dt.HasRows)
                    {
                        dt.Read();
                        id = (int)dt["ID"] + 1;
                    }


                    int pid;

                    var web0 = "https://www.takachi-enclosure.com/";

                    HtmlDocument doc0 = web.Load(web0);
                    /*var img0 = "https://www.takachi-enclosure.com" + doc0.DocumentNode.SelectSingleNode("//div[@class='swiper-slide key-slide1']/img[1]").Attributes["src"].Value;
                    var dsc0 = doc0.DocumentNode.SelectSingleNode("//p[@class='description']").InnerText;
                    //takachi 첫 페이지 저장
                    SqlCommand Cmd0 = new SqlCommand("", cn);
                    Cmd0.CommandText = "insert into OfficialTakachiCategory values ('" + id + "', NULL, NULL , NULL, N'" + dsc0 + "', NULL, '" + img0 + "', NULL, NULL, '" + web0 + "')";
                    Cmd0.ExecuteNonQuery();*/
                    //저장을 하고 다음 페이지로 넘어갈때마다 ID값 1 증가
                    id++;


                    List<string> list = new List<string>();
                    var count = doc0.DocumentNode.SelectNodes("//ul[@class='flex-wrap col-4']/li").Count;
                    for (int i = 1; i <= count; i++)
                    {
                        list.Add(doc0.DocumentNode.SelectSingleNode("//ul[@class='flex-wrap col-4']/li["+i+"]/a[1]").Attributes["href"].Value);
                    }


                    //var y = 1;
                    var y = 2;
                    while (y <= count)
                    {

                        //https://www.leocom.kr/takachi/main_cat/plastic_enclosures.html
                        //정상작동
                        pid = 1;
                        string weblink1 = list[y - 1];
                        HtmlDocument doc1 = web.Load(weblink1);

                        var longtitle = doc1.DocumentNode.SelectSingleNode("//div[@class='title-box']/h2[1]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "").Replace("'", "''");
                        var shorttitle = doc0.DocumentNode.SelectSingleNode("//ul[@class='flex-wrap col-4']/li[" + y + "]/p[1]/a[1]").InnerText.Replace("'", "''");
                        var imglink = "https://www.takachi-enclosure.com" + doc0.DocumentNode.SelectSingleNode("//ul[@class='flex-wrap col-4']/li[" + y + "]/a[1]/img[1]").Attributes["src"].Value;
                        var datasheetlink = "";
                        try
                        {
                            datasheetlink = doc1.DocumentNode.SelectSingleNode("//div[@class='title-box']/a[1]").Attributes["href"].Value.Replace("'", "''");
                        }
                        catch
                        {

                        }
                        var category = "";
                        try
                        {
                            category = doc1.DocumentNode.SelectSingleNode("//p[@class='page-sub-title']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "").Replace("'", "''");
                        }
                        catch
                        {
                            category = "";
                        }
                        var longdsc = doc1.DocumentNode.SelectSingleNode("//p[@class='page-description']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "").Replace("'", "''");

                        //저장
                        SqlCommand Cmd1 = new SqlCommand("", cn);
                        Cmd1.CommandText = "insert into OfficialTakachiCategory values ('" + id + "','" + pid + "', N'" + longtitle + "', '"+shorttitle+"', N'" + longdsc + "', NULL, '"+imglink+"', '" + datasheetlink + "', N'" + category + "', '" + weblink1 + "')";
                        Cmd1.ExecuteNonQuery();
                        pid = id;
                        id++;

                        System.Threading.Thread.Sleep(8000);     //시간 지연 15초

                        //위 카테고리 1개있을시
                        //제품페이지임
                        try
                        {

                            var categorycount = doc1.DocumentNode.SelectNodes("//section[@class='products-list2 products-list2-en']/div[1]/div").Count;
                            var u = 1;
                            //var u = 7;
                            while (u <= categorycount)
                            {
                                var test = doc1.DocumentNode.SelectSingleNode("//section[@class='products-list2 products-list2-en']/div[1]/div[" + u + "]/a[1]").Attributes["href"].Value;
                                if(test.StartsWith("/"))
                                {
                                    test = "https://www.takachi-enclosure.com" + test;
                                }
                                var weblink5 = test;

                                HtmlDocument doc5 = web.Load(weblink5);
                                if (doc5.DocumentNode.InnerText != "")
                                {
                                    if (doc5.DocumentNode.SelectSingleNode("html[1]/head[1]/title[1]").InnerText != "\r\n        500 | 레오콤\r\n    " && doc5.DocumentNode.SelectSingleNode("html[1]/head[1]/title[1]").InnerText != "\r\n        404 | 레오콤\r\n    ")
                                    {


                                        var longtitle2 = doc5.DocumentNode.SelectSingleNode("//h2[@class='page-title']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "").Replace("'", "''");
                                        var shorttitle2 = doc1.DocumentNode.SelectSingleNode("//section[@class='products-list2 products-list2-en']/div[1]/div[" + u + "]/figure[1]/a[1]/img[1]").Attributes["alt"].Value.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                        var category2 = doc5.DocumentNode.SelectSingleNode("//p[@class='page-sub-title']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "").Replace("'", "''");
                                        var longdsc2 = "";
                                        try
                                        {
                                            longdsc2 = doc5.DocumentNode.SelectSingleNode("//div[@class='description with-page-links']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "").Replace("'", "''");
                                        }
                                        catch
                                        {
                                            longdsc2 = doc5.DocumentNode.SelectSingleNode("//div[@class='description ']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                        }

                                        var shortdsc2 = "";
                                        try
                                        {
                                            var shortdsccount = doc1.DocumentNode.SelectNodes("//section[@class='products-list2 products-list2-en']/div[1]/div[" + u + "]/figure[1]/figcaption[1]/ul[1]/li").Count;
                                            var o = 2;
                                            while (o <= shortdsccount)
                                            {
                                                shortdsc2 = shortdsc2 + doc1.DocumentNode.SelectSingleNode("//section[@class='products-list2 products-list2-en']/div[1]/div[" + u + "]/figure[1]/figcaption[1]/ul[1]/li[" + o + "]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                o++;
                                            }
                                        }
                                        catch
                                        {

                                        }

                                        var datasheetlink2 = "";
                                        try
                                        {
                                            datasheetlink2 = doc5.DocumentNode.SelectSingleNode("//div[@class='title-box']/a[1]").Attributes["href"].Value;
                                        }
                                        catch
                                        {

                                        }

                                        var imgHref1 = "";
                                        try
                                        {
                                            var image1 = doc5.DocumentNode.SelectSingleNode("//div[@class='slide-small']/div[1]/div[1]/img[1]");
                                            imgHref1 = image1.Attributes["src"].Value;
                                        }
                                        catch
                                        {

                                        }


                                        //저장
                                        SqlCommand Cmd3 = new SqlCommand("", cn);
                                        Cmd3.CommandText = "insert into OfficialTakachiCategory values ('" + id + "','" + pid + "', N'" + longtitle2 + "', N'" + shorttitle2 + "', N'" + longdsc2 + "', N'" + shortdsc2 + "', '" + imgHref1 + "', '" + datasheetlink2 + "', N'" + category2 + "', '" + weblink5 + "')";
                                        Cmd3.ExecuteNonQuery();
                                        pid = id;
                                        id++;

                                        System.Threading.Thread.Sleep(8000);     //시간 지연 15초

                                        //takachiproduct
                                        //https://www.leocom.kr/takachi/products/WC.aspx
                                        var productcount = doc5.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr").Count;
                                        var r = 2;
                                        int speccount = 0;
                                        var checking = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[1]").InnerText;

                                        if (checking == "W" || checking == "w" || checking == "H" || checking == "h" || checking == "D" || checking == "D")
                                        {
                                            speccount = doc5.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th").Count + 4;
                                            r = 3;

                                            while (r <= productcount)
                                            {
                                                var mytable = new Dictionary<string, string>();
                                                var a = 1;
                                                if (doc5.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td").Count == 6)
                                                {


                                                    while (a <= speccount)
                                                    {
                                                        var key = "";
                                                        var value = "";

                                                        if (a >= 2 && a <= 4)
                                                        {
                                                            key = "외경" + doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[" + (a - 1) + "]").InnerText.ToUpper();

                                                        }
                                                        else if (a >= 5 && a <= 7)
                                                        {
                                                            key = "내경" + doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[" + (a - 1) + "]").InnerText.ToUpper();

                                                        }
                                                        else if (a == 1)
                                                        {
                                                            key = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + a + "]").InnerText;
                                                            if (key == "Product No.")
                                                            {
                                                                key = "형번";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            key = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + (a - 4) + "]").InnerText.Replace("'", "''");

                                                        }
                                                        if (key == "PDF" || key == "3D PDF" || key == "DXF" || key == "DWG" || key == "STP")
                                                        {
                                                            try
                                                            {
                                                                value = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + r + "]/td[" + a + "]/a[1]").Attributes["href"].Value.Replace("'", "''");
                                                            }
                                                            catch
                                                            {
                                                                value = "";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            value = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + r + "]/td[" + a + "]").InnerText.Replace("'", "''");
                                                        }
                                                        mytable.Add(key, value);
                                                        a++;
                                                    }
                                                }
                                                else if (checking == "H" || checking == "h")
                                                {
                                                    while (a <= speccount)
                                                    {
                                                        var key = "";
                                                        var value = "";

                                                        if (a >= 2 && a <= 3)
                                                        {
                                                            key = "외경" + doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[" + (a - 1) + "]").InnerText.ToUpper();

                                                        }
                                                        else if (a >= 4 && a <= 6)
                                                        {
                                                            key = "내경" + doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[" + (a - 1) + "]").InnerText.ToUpper();

                                                        }
                                                        else if (a == 1)
                                                        {
                                                            key = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + a + "]").InnerText;
                                                            if (key == "Product No.")
                                                            {
                                                                key = "형번";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            key = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + (a - 3) + "]").InnerText.Replace("'", "''");

                                                        }
                                                        if (key == "PDF" || key == "3D PDF" || key == "DXF" || key == "DWG" || key == "STP")
                                                        {
                                                            try
                                                            {
                                                                value = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + r + "]/td[" + a + "]/a[1]").Attributes["href"].Value.Replace("'", "''");
                                                            }
                                                            catch
                                                            {
                                                                value = "";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            value = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + r + "]/td[" + a + "]").InnerText.Replace("'", "''");
                                                        }
                                                        mytable.Add(key, value);
                                                        a++;
                                                    }
                                                }
                                                else
                                                {
                                                    while (a <= speccount - 2)
                                                    {
                                                        var key = "";
                                                        var value = "";

                                                        if (a >= 2 && a <= 4)
                                                        {
                                                            if (doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[2]").InnerText == "외경")
                                                            {
                                                                key = "외경" + doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[" + (a - 1) + "]").InnerText.ToUpper();
                                                            }
                                                            else
                                                            {
                                                                key = "내경" + doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[" + (a - 1) + "]").InnerText.ToUpper();
                                                            }
                                                        }
                                                        else if (a == 1)
                                                        {
                                                            key = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + a + "]").InnerText;
                                                            if (key == "Product No.")
                                                            {
                                                                key = "형번";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            key = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + (a - 2) + "]").InnerText.Replace("'", "''");

                                                        }
                                                        if (key == "PDF" || key == "3D PDF" || key == "DXF" || key == "DWG" || key == "STP")
                                                        {
                                                            try
                                                            {
                                                                value = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + r + "]/td[" + a + "]/a[1]").Attributes["href"].Value.Replace("'", "''");
                                                            }
                                                            catch
                                                            {
                                                                value = "";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            value = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + r + "]/td[" + a + "]").InnerText.Replace("'", "''");
                                                        }
                                                        mytable.Add(key, value);
                                                        a++;
                                                    }
                                                }
                                                //저장

                                                TakachiProduct(mytable);


                                                SqlCommand Cmd5 = new SqlCommand("", cn);
                                                Cmd5.CommandText = "update OfficialTakachiProduct set ParentID = '" + pid + "', Weblink = '" + weblink5 + "' where ModelNumber = N'" + mytable["형번"] + "'";
                                                Cmd5.ExecuteNonQuery();
                                                r++;
                                            }

                                        }
                                        else
                                        {
                                            speccount = doc5.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th").Count;
                                            while (r <= productcount)
                                            {
                                                var mytable = new Dictionary<string, string>();
                                                var a = 1;
                                                while (a <= speccount)
                                                {
                                                    var key = "";
                                                    var value = "";
                                                    key = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + a + "]").InnerText;
                                                    if (key == "Product No.")
                                                    {
                                                        key = "형번";
                                                    }
                                                    if (key == "PDF" || key == "3D PDF" || key == "DXF" || key == "DWG" || key == "STP")
                                                    {
                                                        try
                                                        {
                                                            value = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + r + "]/td[" + a + "]/a[1]").Attributes["href"].Value.Replace("'", "''");
                                                        }
                                                        catch
                                                        {
                                                            value = "";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        value = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + r + "]/td[" + a + "]").InnerText.Replace("'", "''");
                                                    }
                                                    mytable.Add(key, value);
                                                    a++;
                                                }
                                                //저장
                                                TakachiProduct(mytable);


                                                SqlCommand Cmd5 = new SqlCommand("", cn);
                                                Cmd5.CommandText = "update OfficialTakachiProduct set ParentID = '" + pid + "', Weblink = '" + weblink5 + "' where ModelNumber = N'" + mytable["형번"] + "'";
                                                Cmd5.ExecuteNonQuery();
                                                r++;
                                            }

                                        }
                                        //같은페이지에 있는 별매품
                                        var tablecount = doc5.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table").Count;
                                        if (tablecount == 2)
                                        {
                                            var productcount2 = doc5.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr").Count;
                                            var m = 3;

                                            var speccount2 = doc5.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[1]/th").Count;

                                            while (m <= productcount2)
                                            {
                                                if (doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr["+m+"]/td[1]").InnerText != "")
                                                {


                                                    var mytable = new Dictionary<string, string>();
                                                    var s = 1;

                                                    while (s <= speccount2)
                                                    {
                                                        var key = "";
                                                        var value = "";
                                                        if (s == 1)
                                                        {
                                                            key = "형번";
                                                        }
                                                        else
                                                        {
                                                            key = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[1]/th[" + s + "]").InnerText.Replace("'", "''");
                                                        }
                                                        if (key == "PDF" || key == "3D PDF" || key == "DXF" || key == "DWG" || key == "STP")
                                                        {
                                                            try
                                                            {
                                                                value = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[" + m + "]/td[" + s + "]/a[1]").Attributes["href"].Value.Replace("'", "''");
                                                            }
                                                            catch
                                                            {
                                                                value = "";
                                                            }
                                                        }
                                                        else
                                                        {

                                                            value = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[" + m + "]/td[" + s + "]").InnerText.Replace("'", "''");

                                                        }
                                                        mytable.Add(key, value);
                                                        s++;
                                                    }
                                                    mytable.Add("별매품 용도", doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[1]/th[1]").InnerText);
                                                    TakachiProduct(mytable);


                                                    SqlCommand Cmd5 = new SqlCommand("", cn);
                                                    Cmd5.CommandText = "update OfficialTakachiProduct set ParentID = '" + (id - 1) + "', Weblink = '" + weblink5 + "' where ModelNumber = N'" + mytable["형번"] + "'";
                                                    Cmd5.ExecuteNonQuery();
                                                }

                                                m++;
                                            }
                                        }
                                    }
                                    else
                                    {

                                    }
                                }
                                u++;
                            }
                        }
                        catch
                        {


                            var categorycount1 = doc1.DocumentNode.SelectNodes("//div[@class='product-row-list']/article").Count;
                            var q = 1;
                            //var q = 6;
                            while (q <= categorycount1)
                            {
                                //https://www.leocom.kr/takachi/cat/handheld_enclosures.html#top
                                //제품페이지 직전
                                //정상작동
                                //var weblink2 = "https://www.leocom.kr/takachi/cat/wallmount_enclosures.html#top";
                                var weblink2 = doc1.DocumentNode.SelectSingleNode("//div[@class='product-row-list']/article[" + q + "]/div[1]/a[1]").Attributes["href"].Value;
                                if (weblink2.StartsWith("/"))
                                {
                                    weblink2 = "https://www.takachi-enclosure.com" + weblink2;
                                }
                                HtmlDocument doc2 = web.Load(weblink2);

                                var shorttitle1 = doc1.DocumentNode.SelectSingleNode("//div[@class='product-row-list']/article[" + q + "]/div[1]/h3[1]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                var longtitle1 = doc2.DocumentNode.SelectSingleNode("//div[@class='title-box']/h2[1]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");

                                var longdsc1 = doc2.DocumentNode.SelectSingleNode("//p[@class='page-description']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "").Replace("'", "''");
                                var shortdsc1 = doc1.DocumentNode.SelectSingleNode("//div[@class='product-row-list']/article[" + q + "]/div[1]/p[1]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "").Replace("'", "''");

                                var datasheetlink1 = "";
                                try
                                {
                                    datasheetlink1 = doc2.DocumentNode.SelectSingleNode("//div[@class='title-box']/a[1]").Attributes["href"].Value;
                                }
                                catch
                                {

                                }

                                var category1 = "";
                                try
                                {
                                    category1 = doc2.DocumentNode.SelectSingleNode("//p[@class='page-sub-title']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                }
                                catch
                                {
                                    category1 = "";
                                }
                                var imgHref = "";
                                try
                                {
                                    var image = doc1.DocumentNode.SelectSingleNode("//div[@class='product-row-list']/article[" + q + "]/div[2]/img[1]");
                                    imgHref = "https://www.takachi-enclosure.com" + image.Attributes["src"].Value;
                                }
                                catch
                                {

                                }


                                //저장
                                SqlCommand Cmd2 = new SqlCommand("", cn);
                                Cmd2.CommandText = "insert into OfficialTakachiCategory values ('" + id + "','" + pid + "', N'" + longtitle1 + "', N'" + shorttitle1 + "', N'" + longdsc1 + "', N'" + shortdsc1 + "', '" + imgHref + "', '" + datasheetlink1 + "', N'" + category1 + "', '" + weblink2 + "')";
                                Cmd2.ExecuteNonQuery();
                                pid = id;
                                id++;

                                System.Threading.Thread.Sleep(8000);     //시간 지연 15초

                                //제품페이지
                                try
                                {
                                    var categorycount2 = doc2.DocumentNode.SelectNodes("//section[@class='products-list2 products-list2-en']/div[1]/div").Count;

                                    var t = 1;
                                    //var t = 5;
                                    while (t <= categorycount2)
                                    {

                                        //위 카테고리 2개 있을때
                                        //정상작동
                                        //https://www.leocom.kr/takachi/products/WC.aspx
                                        var weblink3 = doc2.DocumentNode.SelectSingleNode("//section[@class='products-list2 products-list2-en']/div[1]/div[" + t + "]/figure[1]/a[1]").Attributes["href"].Value;
                                        if (weblink3.StartsWith("/"))
                                        {
                                            weblink3 = "https://www.takachi-enclosure.com" + weblink3;
                                        }
                                        HtmlDocument doc3 = web.Load(weblink3);
                                        if (doc3.DocumentNode.InnerText != "")
                                        {
                                            if (doc3.DocumentNode.SelectSingleNode("html[1]/head[1]/title[1]").InnerText != "\r\n        500 | 레오콤\r\n    " && doc3.DocumentNode.SelectSingleNode("html[1]/head[1]/title[1]").InnerText != "\r\n        404 | 레오콤\r\n    ")
                                            {
                                                var longtitle2 = doc3.DocumentNode.SelectSingleNode("//h2[@class='page-title']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                var category2 = doc3.DocumentNode.SelectSingleNode("//p[@class='page-sub-title']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                var longdsc2 = "";
                                                try
                                                {
                                                    longdsc2 = doc3.DocumentNode.SelectSingleNode("//div[@class='description with-page-links']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "").Replace("'","''");
                                                }
                                                catch
                                                {
                                                    longdsc2 = doc3.DocumentNode.SelectSingleNode("//div[@class='description ']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "").Replace("'", "''");
                                                }

                                                var shorttitle2 = "";
                                                var shortdsc2 = "";
                                                try
                                                {
                                                    shorttitle2 = doc2.DocumentNode.SelectSingleNode("//section[@class='products-list2 products-list2-en']/div[1]/div[" + t + "]/figure[1]/figcaption[1]/ul[1]/li[1]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                    var shortdsccount = doc2.DocumentNode.SelectNodes("//section[@class='products-list2 products-list2-en']/div[1]/div[" + t + "]/figure[1]/figcaption[1]/ul[1]/li").Count;
                                                    var o = 2;
                                                    while (o <= shortdsccount)
                                                    {
                                                        shortdsc2 = shortdsc2 + doc2.DocumentNode.SelectSingleNode("//section[@class='products-list2 products-list2-en']/div[1]/div[" + t + "]/figure[1]/figcaption[1]/ul[1]/li[" + o + "]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                        o++;
                                                    }
                                                }
                                                catch
                                                {

                                                }

                                                var datasheetlink2 = "";
                                                try
                                                {
                                                    datasheetlink2 = doc3.DocumentNode.SelectSingleNode("//div[@class='title-box']/a[1]").Attributes["href"].Value;
                                                }
                                                catch
                                                {

                                                }

                                                var imgHref1 = "";
                                                try
                                                {
                                                    var image1 = doc3.DocumentNode.SelectSingleNode("//div[@class='slide-small']/div[1]/div[1]/img[1]");
                                                    imgHref1 = image1.Attributes["src"].Value;
                                                }
                                                catch
                                                {

                                                }


                                                //저장
                                                SqlCommand Cmd3 = new SqlCommand("", cn);
                                                Cmd3.CommandText = "insert into OfficialTakachiCategory values ('" + id + "','" + pid + "', N'" + longtitle2 + "', N'" + shorttitle2 + "', N'" + longdsc2 + "', N'" + shortdsc2 + "', '" + imgHref1 + "', '" + datasheetlink2 + "', N'" + category2 + "', '" + weblink3 + "')";
                                                Cmd3.ExecuteNonQuery();
                                                id++;

                                                System.Threading.Thread.Sleep(8000);     //시간 지연 15초

                                                //takachiproduct
                                                //정상작동
                                                //https://www.leocom.kr/takachi/products/WC.aspx
                                                var productcount = doc3.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr").Count;
                                                var u = 2;
                                                int speccount = 0;
                                                var checking = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[1]").InnerText;
                                                if (checking == "W" || checking == "w" || checking == "H" || checking == "h" || checking == "D" || checking == "d")
                                                {
                                                    speccount = doc3.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th").Count + doc3.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td").Count - 2;
                                                    u = 3;

                                                    while (u <= productcount)
                                                    {
                                                        var mytable = new Dictionary<string, string>();
                                                        var a = 1;
                                                        if (doc3.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td").Count == 6)
                                                        {
                                                            while (a <= speccount)
                                                            {
                                                                var key = "";
                                                                var value = "";

                                                                if (a >= 2 && a <= 4)
                                                                {
                                                                    key = "외경" + doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[" + (a - 1) + "]").InnerText.ToUpper();

                                                                }
                                                                else if (a >= 5 && a <= 7)
                                                                {
                                                                    key = "내경" + doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[" + (a - 1) + "]").InnerText.ToUpper();

                                                                }
                                                                else if (a == 1)
                                                                {
                                                                    key = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + a + "]").InnerText;
                                                                    if (key == "Product No.")
                                                                    {
                                                                        key = "형번";
                                                                    }

                                                                }
                                                                else
                                                                {
                                                                    key = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + (a - 4) + "]").InnerText;

                                                                }
                                                                if (key == "PDF" || key == "3D PDF" || key == "DXF" || key == "DWG" || key == "STP")
                                                                {
                                                                    try
                                                                    {
                                                                        value = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + u + "]/td[" + a + "]/a[1]").Attributes["href"].Value;
                                                                    }
                                                                    catch
                                                                    {
                                                                        value = "";
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    value = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + u + "]/td[" + a + "]").InnerText;
                                                                }
                                                                mytable.Add(key, value);
                                                                a++;
                                                            }
                                                        }
                                                        else if (checking == "H" || checking == "h")
                                                        {
                                                            while (a <= speccount)
                                                            {
                                                                var key = "";
                                                                var value = "";

                                                                if (a >= 2 && a <= 3)
                                                                {
                                                                    key = "외경" + doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[" + (a - 1) + "]").InnerText.ToUpper();

                                                                }
                                                                else if (a >= 4 && a <= 6)
                                                                {
                                                                    key = "내경" + doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[" + (a - 1) + "]").InnerText.ToUpper();

                                                                }
                                                                else if (a == 1)
                                                                {
                                                                    key = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + a + "]").InnerText;
                                                                    if (key == "Product No.")
                                                                    {
                                                                        key = "형번";
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    key = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + (a - 3) + "]").InnerText;

                                                                }
                                                                if (key == "PDF" || key == "3D PDF" || key == "DXF" || key == "DWG" || key == "STP")
                                                                {
                                                                    try
                                                                    {
                                                                        value = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + u + "]/td[" + a + "]/a[1]").Attributes["href"].Value;
                                                                    }
                                                                    catch
                                                                    {
                                                                        value = "";
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    value = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + u + "]/td[" + a + "]").InnerText;
                                                                }
                                                                mytable.Add(key, value);
                                                                a++;
                                                            }
                                                        }
                                                        else
                                                        {
                                                            while (a <= speccount - 2)
                                                            {
                                                                var key = "";
                                                                var value = "";

                                                                if (a >= 2 && a <= 4)
                                                                {
                                                                    if (doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[2]").InnerText == "외경")
                                                                    {
                                                                        key = "외경" + doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[" + (a - 1) + "]").InnerText.ToUpper();
                                                                    }
                                                                    else
                                                                    {
                                                                        key = "내경" + doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[" + (a - 1) + "]").InnerText.ToUpper();
                                                                    }

                                                                }
                                                                else if (a == 1)
                                                                {
                                                                    key = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + a + "]").InnerText;
                                                                    if (key == "Product No.")
                                                                    {
                                                                        key = "형번";
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    key = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + (a - 2) + "]").InnerText;

                                                                }
                                                                if (key == "PDF" || key == "3D PDF" || key == "DXF" || key == "DWG" || key == "STP")
                                                                {
                                                                    try
                                                                    {
                                                                        value = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + u + "]/td[" + a + "]/a[1]").Attributes["href"].Value;
                                                                    }
                                                                    catch
                                                                    {
                                                                        value = "";
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    value = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + u + "]/td[" + a + "]").InnerText.Replace("'", "''");
                                                                }
                                                                mytable.Add(key, value);
                                                                a++;
                                                            }
                                                        }
                                                        //저장

                                                        TakachiProduct(mytable);


                                                        SqlCommand Cmd5 = new SqlCommand("", cn);
                                                        Cmd5.CommandText = "update OfficialTakachiProduct set ParentID = '" + (id - 1) + "', Weblink = '" + weblink3 + "' where ModelNumber = N'" + mytable["형번"] + "'";
                                                        Cmd5.ExecuteNonQuery();
                                                        u++;
                                                    }



                                                }
                                                else
                                                {
                                                    speccount = doc3.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th").Count;
                                                    while (u <= productcount)
                                                    {
                                                        var mytable = new Dictionary<string, string>();
                                                        var a = 1;
                                                        while (a <= speccount)
                                                        {
                                                            var key = "";
                                                            var value = "";
                                                            key = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + a + "]").InnerText.Replace("'", "''");
                                                            if (key == "Product No.")
                                                            {
                                                                key = "형번";
                                                            }
                                                            if (key == "PDF" || key == "3D PDF" || key == "DXF" || key == "DWG" || key == "STP")
                                                            {
                                                                try
                                                                {
                                                                    value = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + u + "]/td[" + a + "]/a[1]").Attributes["href"].Value.Replace("'", "''");
                                                                }
                                                                catch
                                                                {
                                                                    value = "";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                value = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + u + "]/td[" + a + "]").InnerText.Replace("'", "''");
                                                            }
                                                            mytable.Add(key, value);
                                                            a++;
                                                        }
                                                        //저장
                                                        TakachiProduct(mytable);


                                                        SqlCommand Cmd5 = new SqlCommand("", cn);
                                                        Cmd5.CommandText = "update OfficialTakachiProduct set ParentID = '" + (id - 1) + "', Weblink = '" + weblink3 + "' where ModelNumber = N'" + mytable["형번"] + "'";
                                                        Cmd5.ExecuteNonQuery();
                                                        u++;
                                                    }

                                                }

                                                //같은페이지에 있는 별매품
                                                var tablecount = doc3.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table").Count;
                                                if (tablecount == 2)
                                                {
                                                    var productcount2 = doc3.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr").Count;
                                                    var m = 3;

                                                    var speccount2 = doc3.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[1]/th").Count;

                                                    while (m <= productcount2)
                                                    {
                                                        if (doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr["+m+"]/td[1]").InnerText != "")
                                                        {
                                                            var mytable = new Dictionary<string, string>();
                                                            var s = 1;

                                                            while (s <= speccount2)
                                                            {
                                                                var key = "";
                                                                var value = "";
                                                                if (s == 1)
                                                                {
                                                                    key = "형번";
                                                                }
                                                                else
                                                                {
                                                                    key = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[1]/th[" + s + "]").InnerText;
                                                                }
                                                                if (key == "PDF" || key == "3D PDF" || key == "DXF" || key == "DWG" || key == "STP")
                                                                {
                                                                    try
                                                                    {
                                                                        value = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[" + m + "]/td[" + s + "]/a[1]").Attributes["href"].Value;
                                                                    }
                                                                    catch
                                                                    {
                                                                        value = "";
                                                                    }
                                                                }
                                                                else
                                                                {

                                                                    value = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[" + m + "]/td[" + s + "]").InnerText;

                                                                }
                                                                mytable.Add(key, value);
                                                                s++;
                                                            }

                                                            mytable.Add("별매품 용도", doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[1]/th[1]").InnerText);
                                                            TakachiProduct(mytable);

                                                            SqlCommand Cmd5 = new SqlCommand("", cn);
                                                            Cmd5.CommandText = "update OfficialTakachiProduct set ParentID = '" + (id - 1) + "', Weblink = '" + weblink3 + "' where ModelNumber = N'" + mytable["형번"] + "'";
                                                            Cmd5.ExecuteNonQuery();
                                                        }
                                                        else
                                                        {

                                                        }

                                                        
                                                        m++;
                                                    }
                                                }
                                            }
                                            else
                                            {

                                            }
                                        }
                                        t++;

                                    }





                                }
                                catch
                                {
                                    var categorycount3 = doc2.DocumentNode.SelectNodes("//div[@class='product-row-list']/article").Count;
                                    var e = 1;
                                    //var e = 4;

                                    while (e <= categorycount3)
                                    {
                                        //https://www.leocom.kr/takachi/cat/hinged_plastic_boxes.html#top
                                        //제품페이지 직전
                                        var weblink4 = doc2.DocumentNode.SelectSingleNode("//div[@class='product-row-list']/article[" + e + "]/div[1]/a[1]").Attributes["href"].Value;
                                        if (weblink4.StartsWith("/"))
                                        {
                                            weblink4 = "https://www.takachi-enclosure.com" + weblink4;
                                        }
                                        HtmlDocument doc3 = web.Load(weblink4);


                                        var shorttitle2 = doc2.DocumentNode.SelectSingleNode("//div[@class='product-row-list']/article[" + e + "]/div[1]/h3[1]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                        var longtitle2 = doc3.DocumentNode.SelectSingleNode("//div[@class='title-box']/h2[1]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");

                                        var longdsc2 = doc3.DocumentNode.SelectSingleNode("//p[@class='page-description']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "").Replace("'","");
                                        var shortdsc2 = doc2.DocumentNode.SelectSingleNode("//div[@class='product-row-list']/article[" + e + "]/div[1]/p[1]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "").Replace("'", "");
                                        var datasheetlink2 = "";
                                        try
                                        {
                                            datasheetlink2 = doc3.DocumentNode.SelectSingleNode("//div[@class='title-box']/a[1]").Attributes["href"].Value;
                                        }
                                        catch
                                        {

                                        }

                                        var category2 = "";
                                        try
                                        {
                                            category2 = doc3.DocumentNode.SelectSingleNode("//p[@class='page-sub-title']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                        }
                                        catch
                                        {
                                            category2 = "";
                                        }
                                        var imgHref1 = "";
                                        try
                                        {
                                            var image1 = doc3.DocumentNode.SelectSingleNode("//div[@class='product-row-list']/article[" + e + "]/div[2]/img[1]");
                                            imgHref1 = image1.Attributes["src"].Value;
                                        }
                                        catch
                                        {

                                        }

                                        //저장
                                        SqlCommand Cmd3 = new SqlCommand("", cn);
                                        Cmd3.CommandText = "insert into OfficialTakachiCategory values ('" + id + "','" + pid + "', N'" + longtitle2 + "', N'" + shorttitle2 + "', N'" + longdsc2 + "', N'" + shortdsc2 + "', '" + imgHref1 + "', '" + datasheetlink2 + "', N'" + category2 + "', '" + weblink4 + "')";
                                        Cmd3.ExecuteNonQuery();
                                        id++;

                                        System.Threading.Thread.Sleep(8000);     //시간 지연 15초


                                        //제품페이지
                                        try
                                        {

                                            var categorycount4 = doc3.DocumentNode.SelectNodes("//section[@class='products-list2 products-list2-en']/div[1]/div").Count;
                                            var r = 1;
                                            //var r = 5;
                                            while (r <= categorycount4)
                                            {
                                                //위 카테고리 3개 있을때
                                                //https://www.leocom.kr/takachi/products/UPC.aspx
                                                var weblink5 = doc3.DocumentNode.SelectSingleNode("//section[@class='products-list2 products-list2-en']/div[1]/div[" + r + "]/figure[1]/a[1]").Attributes["href"].Value;
                                                if (weblink5.StartsWith("/"))
                                                {
                                                    weblink5 = "https://www.takachi-enclosure.com" + weblink5;
                                                }
                                                HtmlDocument doc4 = web.Load(weblink5);
                                                if (doc4.DocumentNode.InnerText != "")
                                                {
                                                    if (doc4.DocumentNode.SelectSingleNode("html[1]/head[1]/title[1]").InnerText != "\r\n        500 | 레오콤\r\n    " && doc4.DocumentNode.SelectSingleNode("html[1]/head[1]/title[1]").InnerText != "\r\n        404 | 레오콤\r\n    ")
                                                    {


                                                        var longtitle3 = doc4.DocumentNode.SelectSingleNode("//h2[@class='page-title']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                        var shorttitle3 = doc3.DocumentNode.SelectSingleNode("//section[@class='products-list2 products-list2-en']/div[1]/div[" + r + "]/figure[1]/figcaption[1]/ul[1]/li[1]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                        var category3 = doc4.DocumentNode.SelectSingleNode("//p[@class='page-sub-title']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                        var longdsc3 = "";
                                                        try
                                                        {
                                                            longdsc3 = doc4.DocumentNode.SelectSingleNode("//div[@class='description with-page-links']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                        }
                                                        catch
                                                        {
                                                            longdsc3 = doc4.DocumentNode.SelectSingleNode("//div[@class='description ']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                        }
                                                        var shortdsccount = doc3.DocumentNode.SelectNodes("//section[@class='products-list2 products-list2-en']/div[1]/div[" + r + "]/figure[1]/figcaption[1]/ul[1]/li").Count;
                                                        var shortdsc3 = "";
                                                        var o = 2;
                                                        while (o <= shortdsccount)
                                                        {
                                                            shortdsc2 = shortdsc2 + doc3.DocumentNode.SelectSingleNode("//section[@class='products-list2 products-list2-en']/div[1]/div[" + r + "]/figure[1]/figcaption[1]/ul[1]/li[" + o + "]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                            o++;
                                                        }
                                                        var datasheetlink3 = doc4.DocumentNode.SelectSingleNode("//div[@class='title-box']/a[1]").Attributes["href"].Value;
                                                        var image2 = doc4.DocumentNode.SelectSingleNode("//div[@class='slide-small']/div[1]/div[1]/img[1]");
                                                        var imgHref2 = image2.Attributes["src"].Value;

                                                        //저장
                                                        SqlCommand Cmd4 = new SqlCommand("", cn);
                                                        Cmd4.CommandText = "insert into OfficialTakachiCategory values ('" + id + "','" + pid + "', N'" + longtitle3 + "', N'" + shorttitle3 + "', N'" + longdsc3 + "', N'" + shortdsc3 + "', '" + imgHref2 + "', '" + datasheetlink3 + "', N'" + category3 + "', '" + weblink5 + "')";
                                                        Cmd4.ExecuteNonQuery();
                                                        id++;

                                                        System.Threading.Thread.Sleep(12000);     //시간 지연 15초

                                                        //takachiproduct
                                                        var productcount = doc4.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr").Count;
                                                        var c = 2;
                                                        int speccount = 0;
                                                        var checking = doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[1]").InnerText;

                                                        if (checking == "W" || checking == "w" || checking == "H" || checking == "h" || checking == "D" || checking == "d")
                                                        {
                                                            speccount = doc4.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th").Count + 4;
                                                            c = 3;

                                                            while (c <= productcount)
                                                            {
                                                                var mytable = new Dictionary<string, string>();

                                                                var a = 1;
                                                                if (doc4.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td").Count == 6)
                                                                {


                                                                    while (a <= speccount)
                                                                    {
                                                                        var key = "";
                                                                        var value = "";

                                                                        if (a >= 2 && a <= 4)
                                                                        {
                                                                            key = "외경" + doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[" + (a - 1) + "]").InnerText.ToUpper();

                                                                        }
                                                                        else if (a >= 5 && a <= 7)
                                                                        {
                                                                            key = "내경" + doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[" + (a - 1) + "]").InnerText.ToUpper();

                                                                        }
                                                                        else if (a == 1)
                                                                        {
                                                                            key = doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + a + "]").InnerText;
                                                                            if (key == "Product No.")
                                                                            {
                                                                                key = "형번";
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            key = doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + (a - 4) + "]").InnerText;

                                                                        }
                                                                        if (key == "PDF" || key == "3D PDF" || key == "DXF" || key == "DWG" || key == "STP")
                                                                        {
                                                                            try
                                                                            {
                                                                                value = doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + c + "]/td[" + a + "]/a[1]").Attributes["href"].Value;
                                                                            }
                                                                            catch
                                                                            {
                                                                                value = "";
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            value = doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + c + "]/td[" + a + "]").InnerText;
                                                                        }
                                                                        mytable.Add(key, value);
                                                                        a++;
                                                                    }
                                                                }
                                                                else if (checking == "H" || checking == "h")
                                                                {
                                                                    while (a <= speccount)
                                                                    {
                                                                        var key = "";
                                                                        var value = "";

                                                                        if (a >= 2 && a <= 3)
                                                                        {
                                                                            key = "외경" + doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[" + (a - 1) + "]").InnerText.ToUpper();

                                                                        }
                                                                        else if (a >= 4 && a <= 6)
                                                                        {
                                                                            key = "내경" + doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[" + (a - 1) + "]").InnerText.ToUpper();

                                                                        }
                                                                        else if (a == 1)
                                                                        {
                                                                            key = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + a + "]").InnerText;
                                                                            if (key == "Product No.")
                                                                            {
                                                                                key = "형번";
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            key = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + (a - 3) + "]").InnerText;

                                                                        }
                                                                        if (key == "PDF" || key == "3D PDF" || key == "DXF" || key == "DWG" || key == "STP")
                                                                        {
                                                                            try
                                                                            {
                                                                                value = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + c + "]/td[" + a + "]/a[1]").Attributes["href"].Value;
                                                                            }
                                                                            catch
                                                                            {
                                                                                value = "";
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            value = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + c + "]/td[" + a + "]").InnerText;
                                                                        }
                                                                        mytable.Add(key, value);
                                                                        a++;
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    while (a <= speccount - 2)
                                                                    {
                                                                        var key = "";
                                                                        var value = "";

                                                                        if (a >= 2 && a <= 4)
                                                                        {
                                                                            if (doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[2]").InnerText == "외경")
                                                                            {
                                                                                key = "외경" + doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[" + (a - 1) + "]").InnerText.ToUpper();
                                                                            }
                                                                            else
                                                                            {
                                                                                key = "내경" + doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[" + (a - 1) + "]").InnerText.ToUpper();
                                                                            }
                                                                        }
                                                                        else if (a == 1)
                                                                        {
                                                                            key = doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + a + "]").InnerText;
                                                                            if (key == "Product No.")
                                                                            {
                                                                                key = "형번";
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            key = doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + (a - 2) + "]").InnerText;

                                                                        }
                                                                        if (key == "PDF" || key == "3D PDF" || key == "DXF" || key == "DWG" || key == "STP")
                                                                        {
                                                                            try
                                                                            {
                                                                                value = doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + c + "]/td[" + a + "]/a[1]").Attributes["href"].Value;
                                                                            }
                                                                            catch
                                                                            {
                                                                                value = "";
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            value = doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + c + "]/td[" + a + "]").InnerText;
                                                                        }
                                                                        mytable.Add(key, value);
                                                                        a++;
                                                                    }
                                                                }
                                                                //저장

                                                                TakachiProduct(mytable);


                                                                SqlCommand Cmd5 = new SqlCommand("", cn);
                                                                Cmd5.CommandText = "update OfficialTakachiProduct set ParentID = '" + (id - 1) + "', Weblink = '" + weblink5 + "' where ModelNumber = N'" + mytable["형번"] + "'";
                                                                Cmd5.ExecuteNonQuery();
                                                                c++;
                                                            }

                                                        }
                                                        else
                                                        {
                                                            speccount = doc4.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th").Count;
                                                            while (c <= productcount)
                                                            {
                                                                var mytable = new Dictionary<string, string>();
                                                                var a = 1;
                                                                while (a <= speccount)
                                                                {
                                                                    var key = "";
                                                                    var value = "";
                                                                    key = doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + a + "]").InnerText;
                                                                    if (key == "Product No.")
                                                                    {
                                                                        key = "형번";
                                                                    }
                                                                    if (key == "PDF" || key == "3D PDF" || key == "DXF" || key == "DWG" || key == "STP")
                                                                    {
                                                                        try
                                                                        {
                                                                            value = doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + c + "]/td[" + a + "]/a[1]").Attributes["href"].Value;
                                                                        }
                                                                        catch
                                                                        {
                                                                            value = "";
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        value = doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + c + "]/td[" + a + "]").InnerText;
                                                                    }
                                                                    mytable.Add(key, value);
                                                                    a++;
                                                                }
                                                                //저장
                                                                TakachiProduct(mytable);


                                                                SqlCommand Cmd5 = new SqlCommand("", cn);
                                                                Cmd5.CommandText = "update OfficialTakachiProduct set ParentID = '" + (id - 1) + "', Weblink = '" + weblink5 + "' where ModelNumber = N'" + mytable["형번"] + "'";
                                                                Cmd5.ExecuteNonQuery();
                                                                c++;
                                                            }

                                                        }
                                                        //같은페이지에 있는 별매품
                                                        var tablecount = doc4.DocumentNode.SelectNodes("//div[@class='table-wrap clr_after']/div[1]/table").Count;
                                                        if (tablecount == 2)
                                                        {
                                                            var productcount2 = doc4.DocumentNode.SelectNodes("//div[@class='table-wrap clr_after']/div[1]/table[2]/tbody[1]/tr").Count;
                                                            var m = 3;

                                                            var speccount2 = doc4.DocumentNode.SelectNodes("//div[@class='table-wrap clr_after']/div[1]/table[2]/tbody[1]/tr[1]/th").Count;

                                                            while (m <= productcount2)
                                                            {
                                                                if (doc4.DocumentNode.SelectSingleNode("//div[@class='table-wrap clr_after']/div[1]/table[2]/tbody[1]/tr["+m+"]/td[1]").InnerText != "")
                                                                {
                                                                    var mytable = new Dictionary<string, string>();
                                                                    var s = 1;

                                                                    while (s <= speccount2)
                                                                    {
                                                                        var key = "";
                                                                        var value = "";
                                                                        if (s == 1)
                                                                        {
                                                                            key = "형번";
                                                                        }
                                                                        else
                                                                        {
                                                                            key = doc4.DocumentNode.SelectSingleNode("//div[@class='table-wrap clr_after']/div[1]/table[2]/tbody[1]/tr[1]/th[" + s + "]").InnerText;
                                                                        }
                                                                        if (key == "PDF" || key == "3D PDF" || key == "DXF" || key == "DWG" || key == "STP")
                                                                        {
                                                                            try
                                                                            {
                                                                                value = doc4.DocumentNode.SelectSingleNode("//div[@class='table-wrap clr_after']/div[1]/table[2]/tbody[1]/tr[" + m + "]/td[" + s + "]/a[1]").Attributes["href"].Value;
                                                                            }
                                                                            catch
                                                                            {
                                                                                value = "";
                                                                            }
                                                                        }
                                                                        else
                                                                        {

                                                                            value = doc4.DocumentNode.SelectSingleNode("//div[@class='table-wrap clr_after']/div[1]/table[2]/tbody[1]/tr[" + m + "]/td[" + s + "]").InnerText;

                                                                        }
                                                                        mytable.Add(key, value);
                                                                        s++;
                                                                    }
                                                                    mytable.Add("별매품 용도", doc4.DocumentNode.SelectSingleNode("//div[@class='table-wrap clr_after']/div[1]/table[2]/tbody[1]/tr[1]/th[1]").InnerText);
                                                                    TakachiProduct(mytable);


                                                                    SqlCommand Cmd5 = new SqlCommand("", cn);
                                                                    Cmd5.CommandText = "update OfficialTakachiProduct set ParentID = '" + (id - 1) + "', Weblink = '" + weblink5 + "' where ModelNumber = N'" + mytable["형번"] + "'";
                                                                    Cmd5.ExecuteNonQuery();
                                                                }
                                                                m++;
                                                            }
                                                        }

                                                    }
                                                    else
                                                    {

                                                    }
                                                }
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
                catch (Exception ex)
                {
                    Console.WriteLine("exmsg=" + ex.Message);
                }

            }

        }







        private static void TakachiProduct(Dictionary<string, string> mytable)
        {
            String _connString = System.Configuration.ConfigurationManager.ConnectionStrings["sam"].ToString();

            using (SqlConnection cn = new SqlConnection(_connString))
            {
                cn.Open();


                foreach (KeyValuePair<string, string> items in mytable)
                {
                    if (items.Key == "형번")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "insert into OfficialTakachiProduct (ModelNumber) values (N'" + mytable["형번"] + "')";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "외경W")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set ExternalWidth = N'" + mytable["외경W"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "외경H")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set ExternalHight = N'" + mytable["외경H"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "외경D")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set ExternalDepth = N'" + mytable["외경D"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "내경W")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set InternalWidth = N'" + mytable["내경W"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "내경H")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set InternalHight = N'" + mytable["내경H"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "내경D")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set InternalDepth = N'" + mytable["내경D"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Battery compartment")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Battery compartment] = N'" + mytable["Battery compartment"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key.Contains("weight") || items.Key.Contains("Weight"))
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set Weight = '" + mytable[items.Key] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key.Contains("description"))
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Parts description] = N'" + mytable[items.Key] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "PDF")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set BluePrintPDF = N'" + mytable["PDF"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "3D PDF")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [3DBluePrint] = N'" + mytable["3D PDF"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "DXF")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set DXFBluePrint = N'" + mytable["DXF"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "DWG")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set DWGBluePrint = N'" + mytable["DWG"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Belt to use with (Option)")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Belt to use with (Option)] = N'" + mytable[items.Key] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "color")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set Color = N'" + mytable["color"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Inclined angle")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Inclined angle] = N'" + mytable["Inclined angle"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "material")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [material] = N'" + mytable["material"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Color")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Color] = N'" + mytable[items.Key] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Rated voltage/current")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Rated voltage/current] = N'" + mytable["Rated voltage/current"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Enclosure W Dimension")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Enclosure W Dimension] = N'" + mytable["Enclosure W Dimension"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Equipment")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Equipment] = N'" + mytable["Equipment"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Enclosure H Dimension")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Enclosure H Dimension] = N'" + mytable["Enclosure H Dimension"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Enclosure D Dimension")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Enclosure D Dimension] = N'" + mytable["Enclosure D Dimension"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "프레임의 도장")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [프레임의도장] = N'" + mytable["프레임의 도장"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Finish")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Finish] = N'" + mytable["Finish"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "투명커버유무")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [투명커버유무] = N'" + mytable["투명커버유무"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Type")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Type] = N'" + mytable["Type"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Base type")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Base type] = N'" + mytable["Base type"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key.Contains("Thickness"))
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Thickness] = N'" + mytable[items.Key] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Fixing screw")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Fixing screw] = N'" + mytable["Fixing screw"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "커넥터외형치수")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [커넥터외형치수] = N'" + mytable["커넥터외형치수"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "커넥터 외형치수")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [커넥터외형치수] = N'" + mytable["커넥터 외형치수"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "적합케이블외경")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [적합케이블외경] = N'" + mytable["적합케이블외경"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Cable range")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Cable range] = N'" + mytable["Cable range"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Number of poles")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Number of poles] = N'" + mytable["Number of poles"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "정격전압")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [정격전압] = N'" + mytable["정격전압"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "정격전류")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [정격전류] = N'" + mytable["정격전류"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "적합연결전선")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [적합연결전선] = N'" + mytable["적합연결전선"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key.Contains("Dimension") || items.Key.Contains("dimension"))
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Dimensions (W×D×t) mm] = N'" + mytable[items.Key] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Suitable case")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Suitable case] = N'" + mytable["Suitable case"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "준수 플라스틱 케이스")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [준수플라스틱케이스] = N'" + mytable["준수 플라스틱 케이스"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Pack of")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Pack of] = N'" + mytable["Pack of"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "손잡이길이")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [손잡이길이] = N'" + mytable["손잡이길이"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Mountable thickness")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Mountable thickness] = N'" + mytable["Mountable thickness"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "지원설치판 두께")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [지원설치판두께] = N'" + mytable["지원설치판 두께"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key.Contains("Length"))
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Length(mm)] = N'" + mytable[items.Key] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Screw size")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Screw size] = N'" + mytable["Screw size"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Material / Surface finish")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Material / Surface finish] = N'" + mytable["Material / Surface finish"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "나사길이")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [나사길이] = N'" + mytable["나사길이"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Strap length(mm)")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Strap length(mm)] = N'" + mytable["Strap length(mm)"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "모양 유형")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [모양 유형] = N'" + mytable["모양 유형"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "크기(폭×길이×두께)mm")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [크기(폭×길이×두께)mm] = N'" + mytable["크기(폭×길이×두께)mm"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Height")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Height] = N'" + mytable["Height"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Suitable boxes")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Suitable boxes] = N'" + mytable["Suitable boxes"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Thread size")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Thread size] = N'" + mytable["Thread size"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Mounting hole")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Mounting hole] = N'" + mytable["Mounting hole"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Outer diameter")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Outer diameter] = N'" + mytable["Outer diameter"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Filter effective diameter")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Filter effective diameter] = N'" + mytable["Filter effective diameter"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Size (Ｄ×Ｗ×Ｈ)")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Size (Ｄ×Ｗ×Ｈ)] = N'" + mytable["Size (Ｄ×Ｗ×Ｈ)"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Protection class")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Protection class] = N'" + mytable["Protection class"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "노출부 외형 치수(폭×깊이×두께)")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [노출부 외형 치수(폭×깊이×두께)] = N'" + mytable["노출부 외형 치수(폭×깊이×두께)"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "부착판 폭×높이")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [부착판 폭×높이] = N'" + mytable["부착판 폭×높이"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "전체길이치수")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [전체길이치수] = N'" + mytable["전체길이치수"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Rubber material")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Rubber material] = N'" + mytable["Rubber material"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Suitable screw")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Suitable screw] = N'" + mytable["Suitable screw"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "적합케이스")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [적합케이스] = N'" + mytable["적합케이스"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "별매품 용도")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [별매품 용도] = N'" + mytable["별매품 용도"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Top cover color")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Top cover color] = N'" + mytable["Top cover color"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "Matching fan size")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update OfficialTakachiProduct set [Matching fan size] = N'" + mytable["Matching fan size"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                }
            }
        }




    }
}
