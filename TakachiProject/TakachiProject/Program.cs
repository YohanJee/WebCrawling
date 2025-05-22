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
                    mid.CommandText = "select TOP 1 * from TakachiCategory order by ID desc";
                    SqlDataReader dt = mid.ExecuteReader();
                    if (dt.HasRows)
                    {
                        dt.Read();
                        id = (int)dt["ID"] + 1;
                    }


                    int pid;

                    /*var web0 = "https://www.leocom.kr/takachi/takachi.aspx";
                    var img0 = "https://www.leocom.kr/takachi/images/takachimain.jpg";
                    var dsc0 = "타카치 전기공업(TAKACHI ENCLOSURE)은 일본을 대표하는 고품질의 다양한 Enclosure(용기)를 제작합니다. 2만여종의 각종 산업용 케이스등을 생산합니다.\r\n\r\n[레오콤이용시 장점]\r\n● 타카치 전기공업(TAKACHI ENCLOSURE) 전제품 국내 최저가 지향합니다. 타사와 가격비교 바랍니다.\r\n● 타카치 전기공업(TAKACHI ENCLOSURE) 전제품 1개부터 주문 가능합니다.\r\n● 예상 납기: 일본 타카치 전기공업(TAKACHI ENCLOSURE) 으로 부터 약10일(영업일기준) 소요 됩니다.(일본에 재고 보유시)";
                    //takachi 첫 페이지 저장
                    SqlCommand Cmd0 = new SqlCommand("", cn);
                    Cmd0.CommandText = "insert into TakachiCategory values ('" + id + "', NULL, NULL , NULL, N'" + dsc0 + "', NULL, '" + img0 + "', NULL, NULL, '" + web0 + "')";
                    Cmd0.ExecuteNonQuery();
                    //저장을 하고 다음 페이지로 넘어갈때마다 ID값 1 증가
                    id++;*/

                    List<string> list = new List<string>();
                    list.Add("https://www.leocom.kr/takachi/main_cat/plastic_enclosures.html");
                    list.Add("https://www.leocom.kr/takachi/main_cat/aluminium_enclosures.html");
                    list.Add("https://www.leocom.kr/takachi/main_cat/carrying_cases.html");
                    list.Add("https://www.leocom.kr/takachi/main_cat/waterproof_plastic_boxes.html");
                    list.Add("https://www.leocom.kr/takachi/main_cat/waterproof_aluminium_boxes.html");
                    list.Add("https://www.leocom.kr/takachi/main_cat/stainless_steel_enclosures.html");
                    list.Add("https://www.leocom.kr/takachi/main_cat/junction_boxes.html");
                    list.Add("https://www.leocom.kr/takachi/main_cat/connector.html");
                    list.Add("https://www.leocom.kr/takachi/main_cat/flexible_size.html");
                    list.Add("https://www.leocom.kr/takachi/main_cat/rackmount.html");
                    list.Add("https://www.leocom.kr/takachi/main_cat/19inch_rack.html");
                    list.Add("https://www.leocom.kr/takachi/main_cat/accessories.html");
                    list.Add("https://www.leocom.kr/takachi/main_cat/metal_panels.html");

                    //var y = 1;
                    var y = 11;
                    while (y <= 13)
                    {

                        //https://www.leocom.kr/takachi/main_cat/plastic_enclosures.html
                        //타카치 메인페이지에 있는 13개 종류 첫페이지들
                        pid = 1;
                        string weblink1 = list[y - 1];
                        HtmlDocument doc1 = web.Load(weblink1);

                        var longtitle = doc1.DocumentNode.SelectSingleNode("//div[@class='title-box']/h2[1]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                        var datasheetlink = "";
                        try
                        {
                            datasheetlink = "https://www.leocom.kr/takachi" + doc1.DocumentNode.SelectSingleNode("//div[@class='title-box']/a[1]").Attributes["href"].Value.Substring(2);
                        }
                        catch
                        {

                        }
                        var category = "";
                        try
                        {
                            category = doc1.DocumentNode.SelectSingleNode("//p[@class='page-sub-title']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                        }
                        catch
                        {
                            category = "";
                        }
                        var longdsc = doc1.DocumentNode.SelectSingleNode("//p[@class='page-description']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");

                        //저장
                        SqlCommand Cmd1 = new SqlCommand("", cn);
                        Cmd1.CommandText = "insert into TakachiCategory values ('" + id + "','" + pid + "', N'" + longtitle + "', NULL, N'" + longdsc + "', NULL, NULL, '" + datasheetlink + "', N'" + category + "', '" + weblink1 + "')";
                        Cmd1.ExecuteNonQuery();
                        pid = id;
                        id++;

                        System.Threading.Thread.Sleep(8000);     //시간 지연 15초

                        //위 카테고리 1개있을시
                        //제품페이지임
                        try
                        {
                            
                            var categorycount = doc1.DocumentNode.SelectNodes("//section[@class='products-list2']/div[1]/div").Count;
                            var u = 1;
                            //var u = 7;
                            while (u <= categorycount)
                            {
                                var test = doc1.DocumentNode.SelectSingleNode("//section[@class='products-list2']/div[1]/div[" + u + "]/a[1]").Attributes["href"].Value;
                                if (test.StartsWith(" "))
                                {
                                    test = test.Substring(3);
                                }
                                else
                                {
                                    test = test.Substring(2);
                                }
                                var weblink5 = "https://www.leocom.kr/takachi" + test;

                                HtmlDocument doc5 = web.Load(weblink5);
                                if (doc5.DocumentNode.InnerText != "")
                                {
                                    if (doc5.DocumentNode.SelectSingleNode("html[1]/head[1]/title[1]").InnerText != "\r\n        500 | 레오콤\r\n    " && doc5.DocumentNode.SelectSingleNode("html[1]/head[1]/title[1]").InnerText != "\r\n        404 | 레오콤\r\n    ")
                                    {


                                        var longtitle2 = doc5.DocumentNode.SelectSingleNode("//h2[@class='page-title']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                        var shorttitle2 = doc1.DocumentNode.SelectSingleNode("//section[@class='products-list2']/div[1]/div[" + u + "]/figure[1]/a[1]/img[1]").Attributes["alt"].Value.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                        var category2 = doc5.DocumentNode.SelectSingleNode("//p[@class='page-sub-title']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                        var longdsc2 = doc5.DocumentNode.SelectSingleNode("//div[@class='description']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");

                                        var shortdsc2 = "";
                                        try
                                        {
                                            var shortdsccount = doc1.DocumentNode.SelectNodes("//section[@class='products-list2']/div[1]/div[" + u + "]/figure[1]/figcaption[1]/ul[1]/li").Count;
                                            var o = 2;
                                            while (o <= shortdsccount)
                                            {
                                                shortdsc2 = shortdsc2 + doc1.DocumentNode.SelectSingleNode("//section[@class='products-list2']/div[1]/div[" + u + "]/figure[1]/figcaption[1]/ul[1]/li[" + o + "]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                o++;
                                            }
                                        }
                                        catch
                                        {
                                            
                                        }

                                        



                                        var datasheetlink2 = "";
                                        try
                                        {
                                            datasheetlink2 = "https://www.leocom.kr/takachi" + doc5.DocumentNode.SelectSingleNode("//div[@class='title-box']/a[1]").Attributes["href"].Value.Substring(2);
                                        }
                                        catch
                                        {

                                        }

                                        var imgHref1 = "";
                                        try
                                        {
                                            var image1 = doc5.DocumentNode.SelectSingleNode("//div[@class='slide-small']/div[1]/div[1]/img[1]");
                                            imgHref1 = "https://www.leocom.kr/takachi" + image1.Attributes["src"].Value.Substring(2);
                                        }
                                        catch
                                        {

                                        }


                                        //저장
                                        SqlCommand Cmd3 = new SqlCommand("", cn);
                                        Cmd3.CommandText = "insert into TakachiCategory values ('" + id + "','" + pid + "', N'" + longtitle2 + "', N'" + shorttitle2 + "', N'" + longdsc2 + "', N'" + shortdsc2 + "', '" + imgHref1 + "', '" + datasheetlink2 + "', N'" + category2 + "', '" + weblink5 + "')";
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
                                                            if (key.Contains("모델"))
                                                            {
                                                                key = "형번";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            key = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + (a - 4) + "]").InnerText;

                                                        }
                                                        if (key == "PDF도면" || key == "3D PDF" || key == "DXF도면" || key == "DWG도면" || key == "주문")
                                                        {
                                                            try
                                                            {
                                                                value = "https://www.leocom.kr/takachi" + doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + r + "]/td[" + a + "]/a[1]").Attributes["href"].Value.Substring(2);
                                                            }
                                                            catch
                                                            {
                                                                value = "";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            value = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + r + "]/td[" + a + "]").InnerText;
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
                                                            if (key.Contains("모델"))
                                                            {
                                                                key = "형번";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            key = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + (a - 3) + "]").InnerText;

                                                        }
                                                        if (key == "PDF도면" || key == "3D PDF" || key == "DXF도면" || key == "DWG도면" || key == "주문")
                                                        {
                                                            try
                                                            {
                                                                value = "https://www.leocom.kr/takachi" + doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + r + "]/td[" + a + "]/a[1]").Attributes["href"].Value.Substring(2);
                                                            }
                                                            catch
                                                            {
                                                                value = "";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            value = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + r + "]/td[" + a + "]").InnerText;
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
                                                            if (key.Contains("모델"))
                                                            {
                                                                key = "형번";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            key = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + (a - 2) + "]").InnerText;

                                                        }
                                                        if (key == "PDF도면" || key == "3D PDF" || key == "DXF도면" || key == "DWG도면" || key == "주문")
                                                        {
                                                            try
                                                            {
                                                                value = "https://www.leocom.kr/takachi" + doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + r + "]/td[" + a + "]/a[1]").Attributes["href"].Value.Substring(2);
                                                            }
                                                            catch
                                                            {
                                                                value = "";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            value = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + r + "]/td[" + a + "]").InnerText;
                                                        }
                                                        mytable.Add(key, value);
                                                        a++;
                                                    }
                                                }
                                                //저장

                                                TakachiProduct(mytable);


                                                SqlCommand Cmd5 = new SqlCommand("", cn);
                                                Cmd5.CommandText = "update TakachiProduct set ParentID = '" + pid + "', Weblink = '" + weblink5 + "' where ModelNumber = N'" + mytable["형번"] + "'";
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
                                                    if (key.Contains("모델"))
                                                    {
                                                        key = "형번";
                                                    }
                                                    if (key == "PDF도면" || key == "3D PDF" || key == "DXF도면" || key == "DWG도면" || key == "주문")
                                                    {
                                                        try
                                                        {
                                                            value = "https://www.leocom.kr/takachi" + doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + r + "]/td[" + a + "]/a[1]").Attributes["href"].Value.Substring(2);
                                                        }
                                                        catch
                                                        {
                                                            value = "";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        value = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + r + "]/td[" + a + "]").InnerText;
                                                    }
                                                    mytable.Add(key, value);
                                                    a++;
                                                }
                                                //저장
                                                TakachiProduct(mytable);


                                                SqlCommand Cmd5 = new SqlCommand("", cn);
                                                Cmd5.CommandText = "update TakachiProduct set ParentID = '" + pid + "', Weblink = '" + weblink5 + "' where ModelNumber = N'" + mytable["형번"] + "'";
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
                                                        key = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[1]/th[" + s + "]").InnerText;
                                                    }
                                                    if (key == "PDF도면" || key == "3D PDF" || key == "DXF도면" || key == "DWG도면" || key == "주문")
                                                    {
                                                        try
                                                        {
                                                            value = "https://www.leocom.kr/takachi" + doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[" + m + "]/td[" + s + "]/a[1]").Attributes["href"].Value.Substring(2);
                                                        }
                                                        catch
                                                        {
                                                            value = "";
                                                        }
                                                    }
                                                    else
                                                    {

                                                        value = doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[" + m + "]/td[" + s + "]").InnerText;

                                                    }
                                                    mytable.Add(key, value);
                                                    s++;
                                                }
                                                mytable.Add("별매품 용도", doc5.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[1]/th[1]").InnerText);
                                                TakachiProduct(mytable);


                                                SqlCommand Cmd5 = new SqlCommand("", cn);
                                                Cmd5.CommandText = "update TakachiProduct set ParentID = '" + (id - 1) + "', Weblink = '" + weblink5 + "' where ModelNumber = N'" + mytable["형번"] + "'";
                                                Cmd5.ExecuteNonQuery();
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
                            //var q = 2;
                            while (q <= categorycount1)
                            {
                                //https://www.leocom.kr/takachi/cat/handheld_enclosures.html#top
                                //제품페이지 직전
                                //var weblink2 = "https://www.leocom.kr/takachi/cat/wallmount_enclosures.html#top";
                                var weblink2 = "https://www.leocom.kr/takachi" + doc1.DocumentNode.SelectSingleNode("//div[@class='product-row-list']/article[" + q + "]/div[1]/a[1]").Attributes["href"].Value.Substring(2);

                                HtmlDocument doc2 = web.Load(weblink2);

                                var shorttitle1 = doc1.DocumentNode.SelectSingleNode("//div[@class='product-row-list']/article[" + q + "]/div[1]/h3[1]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                var longtitle1 = doc2.DocumentNode.SelectSingleNode("//div[@class='title-box']/h2[1]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");

                                var longdsc1 = doc2.DocumentNode.SelectSingleNode("//p[@class='page-description']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                var shortdsc1 = doc1.DocumentNode.SelectSingleNode("//div[@class='product-row-list']/article[" + q + "]/div[1]/p[1]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");

                                var datasheetlink1 = "";
                                try
                                {
                                    datasheetlink1 = "https://www.leocom.kr/takachi" + doc2.DocumentNode.SelectSingleNode("//div[@class='title-box']/a[1]").Attributes["href"].Value.Substring(2);
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
                                    imgHref = "https://www.leocom.kr/takachi" + image.Attributes["src"].Value.Substring(2);
                                }
                                catch
                                {

                                }


                                //저장
                                SqlCommand Cmd2 = new SqlCommand("", cn);
                                Cmd2.CommandText = "insert into TakachiCategory values ('" + id + "','" + pid + "', N'" + longtitle1 + "', N'" + shorttitle1 + "', N'" + longdsc1 + "', N'" + shortdsc1 + "', '" + imgHref + "', '" + datasheetlink1 + "', N'" + category1 + "', '" + weblink2 + "')";
                                Cmd2.ExecuteNonQuery();
                                pid = id;
                                id++;

                                System.Threading.Thread.Sleep(8000);     //시간 지연 15초

                                //제품페이지
                                try
                                {
                                    var categorycount2 = doc2.DocumentNode.SelectNodes("//section[@class='products-list2']/div[1]/div").Count;

                                    var t = 1;
                                    //var t = 17;
                                    while (t <= categorycount2)
                                    {

                                        //위 카테고리 2개 있을때
                                        //https://www.leocom.kr/takachi/products/WC.aspx
                                        var weblink3 = "https://www.leocom.kr/takachi" + doc2.DocumentNode.SelectSingleNode("//section[@class='products-list2']/div[1]/div[" + t + "]/figure[1]/a[1]").Attributes["href"].Value.Substring(2);
                                        //var weblink3 = "https://www.leocom.kr/takachi/products/AWP.aspx";
                                        HtmlDocument doc3 = web.Load(weblink3);
                                        if (doc3.DocumentNode.InnerText != "")
                                        {
                                            if (doc3.DocumentNode.SelectSingleNode("html[1]/head[1]/title[1]").InnerText != "\r\n        500 | 레오콤\r\n    " && doc3.DocumentNode.SelectSingleNode("html[1]/head[1]/title[1]").InnerText != "\r\n        404 | 레오콤\r\n    ")
                                            {
                                                var longtitle2 = doc3.DocumentNode.SelectSingleNode("//h2[@class='page-title']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                var category2 = doc3.DocumentNode.SelectSingleNode("//p[@class='page-sub-title']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                var longdsc2 = doc3.DocumentNode.SelectSingleNode("//div[@class='description']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");

                                                var shorttitle2 = "";
                                                var shortdsc2 = "";
                                                try
                                                {
                                                    shorttitle2 = doc2.DocumentNode.SelectSingleNode("//section[@class='products-list2']/div[1]/div[" + t + "]/figure[1]/figcaption[1]/ul[1]/li[1]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                    var shortdsccount = doc2.DocumentNode.SelectNodes("//section[@class='products-list2']/div[1]/div[" + t + "]/figure[1]/figcaption[1]/ul[1]/li").Count;
                                                    var o = 2;
                                                    while (o <= shortdsccount)
                                                    {
                                                        shortdsc2 = shortdsc2 + doc2.DocumentNode.SelectSingleNode("//section[@class='products-list2']/div[1]/div[" + t + "]/figure[1]/figcaption[1]/ul[1]/li[" + o + "]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                        o++;
                                                    }
                                                }
                                                catch
                                                {

                                                }
                                                
                                                var datasheetlink2 = "";
                                                try
                                                {
                                                    datasheetlink2 = "https://www.leocom.kr/takachi" + doc3.DocumentNode.SelectSingleNode("//div[@class='title-box']/a[1]").Attributes["href"].Value.Substring(2);
                                                }
                                                catch
                                                {

                                                }

                                                var imgHref1 = "";
                                                try
                                                {
                                                    var image1 = doc3.DocumentNode.SelectSingleNode("//div[@class='slide-small']/div[1]/div[1]/img[1]");
                                                    imgHref1 = "https://www.leocom.kr/takachi" + image1.Attributes["src"].Value.Substring(2);
                                                }
                                                catch
                                                {

                                                }


                                                //저장
                                                SqlCommand Cmd3 = new SqlCommand("", cn);
                                                Cmd3.CommandText = "insert into TakachiCategory values ('" + id + "','" + pid + "', N'" + longtitle2 + "', N'" + shorttitle2 + "', N'" + longdsc2 + "', N'" + shortdsc2 + "', '" + imgHref1 + "', '" + datasheetlink2 + "', N'" + category2 + "', '" + weblink3 + "')";
                                                Cmd3.ExecuteNonQuery();
                                                id++;

                                                System.Threading.Thread.Sleep(8000);     //시간 지연 15초

                                                //takachiproduct
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
                                                                    if (key.Contains("모델"))
                                                                    {
                                                                        key = "형번";
                                                                    }

                                                                }
                                                                else
                                                                {
                                                                    key = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + (a - 4) + "]").InnerText;

                                                                }
                                                                if (key == "PDF도면" || key == "3D PDF" || key == "DXF도면" || key == "DWG도면" || key == "주문")
                                                                {
                                                                    try
                                                                    {
                                                                        value = "https://www.leocom.kr/takachi" + doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + u + "]/td[" + a + "]/a[1]").Attributes["href"].Value.Substring(2);
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
                                                                    if (key.Contains("모델"))
                                                                    {
                                                                        key = "형번";
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    key = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + (a - 3) + "]").InnerText;

                                                                }
                                                                if (key == "PDF도면" || key == "3D PDF" || key == "DXF도면" || key == "DWG도면" || key == "주문")
                                                                {
                                                                    try
                                                                    {
                                                                        value = "https://www.leocom.kr/takachi" + doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + u + "]/td[" + a + "]/a[1]").Attributes["href"].Value.Substring(2);
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
                                                                    if (key.Contains("모델"))
                                                                    {
                                                                        key = "형번";
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    key = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + (a - 2) + "]").InnerText;

                                                                }
                                                                if (key == "PDF도면" || key == "3D PDF" || key == "DXF도면" || key == "DWG도면" || key == "주문")
                                                                {
                                                                    try
                                                                    {
                                                                        value = "https://www.leocom.kr/takachi" + doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + u + "]/td[" + a + "]/a[1]").Attributes["href"].Value.Substring(2);
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
                                                        //저장

                                                        TakachiProduct(mytable);


                                                        SqlCommand Cmd5 = new SqlCommand("", cn);
                                                        Cmd5.CommandText = "update TakachiProduct set ParentID = '" + (id - 1) + "', Weblink = '" + weblink3 + "' where ModelNumber = N'" + mytable["형번"] + "'";
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
                                                            key = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + a + "]").InnerText;
                                                            if (key.Contains("모델"))
                                                            {
                                                                key = "형번";
                                                            }
                                                            if (key == "PDF도면" || key == "3D PDF" || key == "DXF도면" || key == "DWG도면" || key == "주문")
                                                            {
                                                                try
                                                                {
                                                                    value = "https://www.leocom.kr/takachi" + doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + u + "]/td[" + a + "]/a[1]").Attributes["href"].Value.Substring(2);
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
                                                        //저장
                                                        TakachiProduct(mytable);


                                                        SqlCommand Cmd5 = new SqlCommand("", cn);
                                                        Cmd5.CommandText = "update TakachiProduct set ParentID = '" + (id - 1) + "', Weblink = '" + weblink3 + "' where ModelNumber = N'" + mytable["형번"] + "'";
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
                                                            if (key == "PDF도면" || key == "3D PDF" || key == "DXF도면" || key == "DWG도면" || key == "주문")
                                                            {
                                                                try
                                                                {
                                                                    value = "https://www.leocom.kr/takachi" + doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[" + m + "]/td[" + s + "]/a[1]").Attributes["href"].Value.Substring(2);
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
                                                        Cmd5.CommandText = "update TakachiProduct set ParentID = '" + (id - 1) + "', Weblink = '" + weblink3 + "' where ModelNumber = N'" + mytable["형번"] + "'";
                                                        Cmd5.ExecuteNonQuery();
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
                                    //var e = 3;

                                    while (e <= categorycount3)
                                    {
                                        //https://www.leocom.kr/takachi/cat/hinged_plastic_boxes.html#top
                                        //제품페이지 직전
                                        var weblink4 = "https://www.leocom.kr/takachi" + doc2.DocumentNode.SelectSingleNode("//div[@class='product-row-list']/article[" + e + "]/div[1]/a[1]").Attributes["href"].Value.Substring(2);

                                        HtmlDocument doc3 = web.Load(weblink4);


                                        var shorttitle2 = doc2.DocumentNode.SelectSingleNode("//div[@class='product-row-list']/article[" + e + "]/div[1]/h3[1]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                        var longtitle2 = doc3.DocumentNode.SelectSingleNode("//div[@class='title-box']/h2[1]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");

                                        var longdsc2 = doc3.DocumentNode.SelectSingleNode("//p[@class='page-description']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                        var shortdsc2 = doc2.DocumentNode.SelectSingleNode("//div[@class='product-row-list']/article[" + e + "]/div[1]/p[1]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                        var datasheetlink2 = "";
                                        try
                                        {
                                            datasheetlink2 = "https://www.leocom.kr/takachi" + doc3.DocumentNode.SelectSingleNode("//div[@class='title-box']/a[1]").Attributes["href"].Value.Substring(2);
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
                                            imgHref1 = "https://www.leocom.kr/takachi" + image1.Attributes["src"].Value.Substring(2);
                                        }
                                        catch
                                        {

                                        }

                                        //저장
                                        SqlCommand Cmd3 = new SqlCommand("", cn);
                                        Cmd3.CommandText = "insert into TakachiCategory values ('" + id + "','" + pid + "', N'" + longtitle2 + "', N'" + shorttitle2 + "', N'" + longdsc2 + "', N'" + shortdsc2 + "', '" + imgHref1 + "', '" + datasheetlink2 + "', N'" + category2 + "', '" + weblink4 + "')";
                                        Cmd3.ExecuteNonQuery();
                                        id++;

                                        System.Threading.Thread.Sleep(8000);     //시간 지연 15초


                                        //제품페이지
                                        try
                                        {

                                            var categorycount4 = doc3.DocumentNode.SelectNodes("//section[@class='products-list2']/div[1]/div").Count;
                                            var r = 1;
                                            //var r = 8;
                                            while (r <= categorycount4)
                                            {
                                                //위 카테고리 3개 있을때
                                                //https://www.leocom.kr/takachi/products/UPC.aspx
                                                var weblink5 = "https://www.leocom.kr/takachi" + doc3.DocumentNode.SelectSingleNode("//section[@class='products-list2']/div[1]/div[" + r + "]/figure[1]/a[1]").Attributes["href"].Value.Substring(2);

                                                HtmlDocument doc4 = web.Load(weblink5);
                                                if (doc4.DocumentNode.InnerText != "")
                                                {
                                                    if (doc4.DocumentNode.SelectSingleNode("html[1]/head[1]/title[1]").InnerText != "\r\n        500 | 레오콤\r\n    " && doc4.DocumentNode.SelectSingleNode("html[1]/head[1]/title[1]").InnerText != "\r\n        404 | 레오콤\r\n    ")
                                                    {


                                                        var longtitle3 = doc4.DocumentNode.SelectSingleNode("//h2[@class='page-title']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                        var shorttitle3 = doc3.DocumentNode.SelectSingleNode("//section[@class='products-list2']/div[1]/div[" + r + "]/figure[1]/figcaption[1]/ul[1]/li[1]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                        var category3 = doc4.DocumentNode.SelectSingleNode("//p[@class='page-sub-title']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                        var longdsc3 = doc4.DocumentNode.SelectSingleNode("//div[@class='description']").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                        var shortdsccount = doc3.DocumentNode.SelectNodes("//section[@class='products-list2']/div[1]/div[" + r + "]/figure[1]/figcaption[1]/ul[1]/li").Count;
                                                        var shortdsc3 = "";
                                                        var o = 2;
                                                        while (o <= shortdsccount)
                                                        {
                                                            shortdsc2 = shortdsc2 + doc3.DocumentNode.SelectSingleNode("//section[@class='products-list2']/div[1]/div[" + r + "]/figure[1]/figcaption[1]/ul[1]/li[" + o + "]").InnerText.Replace("\r", "").Replace("\n", "").Replace("  ", "");
                                                            o++;
                                                        }
                                                        var datasheetlink3 = "https://www.leocom.kr/takachi" + doc4.DocumentNode.SelectSingleNode("//div[@class='title-box']/a[1]").Attributes["href"].Value.Substring(2);
                                                        var image2 = doc4.DocumentNode.SelectSingleNode("//div[@class='slide-small']/div[1]/div[1]/img[1]");
                                                        var imgHref2 = "https://www.leocom.kr/takachi" + image2.Attributes["src"].Value.Substring(2);

                                                        //저장
                                                        SqlCommand Cmd4 = new SqlCommand("", cn);
                                                        Cmd4.CommandText = "insert into TakachiCategory values ('" + id + "','" + pid + "', N'" + longtitle3 + "', N'" + shorttitle3 + "', N'" + longdsc3 + "', N'" + shortdsc3 + "', '" + imgHref2 + "', '" + datasheetlink3 + "', N'" + category3 + "', '" + weblink5 + "')";
                                                        Cmd4.ExecuteNonQuery();
                                                        id++;

                                                        System.Threading.Thread.Sleep(8000);     //시간 지연 15초

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
                                                                            if (key.Contains("모델"))
                                                                            {
                                                                                key = "형번";
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            key = doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + (a - 4) + "]").InnerText;

                                                                        }
                                                                        if (key == "PDF도면" || key == "3D PDF" || key == "DXF도면" || key == "DWG도면" || key == "주문")
                                                                        {
                                                                            try
                                                                            {
                                                                                value = "https://www.leocom.kr/takachi" + doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + c + "]/td[" + a + "]/a[1]").Attributes["href"].Value.Substring(2);
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
                                                                            if (key.Contains("모델"))
                                                                            {
                                                                                key = "형번";
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            key = doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + (a - 3) + "]").InnerText;

                                                                        }
                                                                        if (key == "PDF도면" || key == "3D PDF" || key == "DXF도면" || key == "DWG도면" || key == "주문")
                                                                        {
                                                                            try
                                                                            {
                                                                                value = "https://www.leocom.kr/takachi" + doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + c + "]/td[" + a + "]/a[1]").Attributes["href"].Value.Substring(2);
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
                                                                            if (key.Contains("모델"))
                                                                            {
                                                                                key = "형번";
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            key = doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[1]/th[" + (a - 2) + "]").InnerText;

                                                                        }
                                                                        if (key == "PDF도면" || key == "3D PDF" || key == "DXF도면" || key == "DWG도면" || key == "주문")
                                                                        {
                                                                            try
                                                                            {
                                                                                value = "https://www.leocom.kr/takachi" + doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + c + "]/td[" + a + "]/a[1]").Attributes["href"].Value.Substring(2);
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
                                                                Cmd5.CommandText = "update TakachiProduct set ParentID = '" + (id - 1) + "', Weblink = '" + weblink5 + "' where ModelNumber = N'" + mytable["형번"] + "'";
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
                                                                    if (key.Contains("모델"))
                                                                    {
                                                                        key = "형번";
                                                                    }
                                                                    if (key == "PDF도면" || key == "3D PDF" || key == "DXF도면" || key == "DWG도면" || key == "주문")
                                                                    {
                                                                        try
                                                                        {
                                                                            value = "https://www.leocom.kr/takachi" + doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[1]/tbody[1]/tr[" + c + "]/td[" + a + "]/a[1]").Attributes["href"].Value.Substring(2);
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
                                                                Cmd5.CommandText = "update TakachiProduct set ParentID = '" + (id - 1) + "', Weblink = '" + weblink5 + "' where ModelNumber = N'" + mytable["형번"] + "'";
                                                                Cmd5.ExecuteNonQuery();
                                                                c++;
                                                            }

                                                        }
                                                        //같은페이지에 있는 별매품
                                                        var tablecount = doc4.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table").Count;
                                                        if (tablecount == 2)
                                                        {
                                                            var productcount2 = doc4.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr").Count;
                                                            var m = 3;

                                                            var speccount2 = doc4.DocumentNode.SelectNodes("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[1]/th").Count;

                                                            while (m <= productcount2)
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
                                                                        key = doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[1]/th[" + s + "]").InnerText;
                                                                    }
                                                                    if (key == "PDF도면" || key == "3D PDF" || key == "DXF도면" || key == "DWG도면" || key == "주문")
                                                                    {
                                                                        try
                                                                        {
                                                                            value = "https://www.leocom.kr/takachi" + doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[" + m + "]/td[" + s + "]/a[1]").Attributes["href"].Value.Substring(2);
                                                                        }
                                                                        catch
                                                                        {
                                                                            value = "";
                                                                        }
                                                                    }
                                                                    else
                                                                    {

                                                                        value = doc4.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[" + m + "]/td[" + s + "]").InnerText;

                                                                    }
                                                                    mytable.Add(key, value);
                                                                    s++;
                                                                }
                                                                mytable.Add("별매품 용도", doc3.DocumentNode.SelectSingleNode("//section[@class='common-section detail-section']/div[1]/div[1]/table[2]/tbody[1]/tr[1]/th[1]").InnerText);
                                                                TakachiProduct(mytable);


                                                                SqlCommand Cmd5 = new SqlCommand("", cn);
                                                                Cmd5.CommandText = "update TakachiProduct set ParentID = '" + (id - 1) + "', Weblink = '" + weblink5 + "' where ModelNumber = N'" + mytable["형번"] + "'";
                                                                Cmd5.ExecuteNonQuery();
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
                        Cmd4.CommandText = "insert into TakachiProduct (ModelNumber) values (N'" + mytable["형번"] + "')";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "외경W")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set ExternalWidth = N'" + mytable["외경W"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "외경H")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set ExternalHight = N'" + mytable["외경H"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "외경D")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set ExternalDepth = N'" + mytable["외경D"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "내경W")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set InternalWidth = N'" + mytable["내경W"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "내경H")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set InternalHight = N'" + mytable["내경H"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "내경D")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set InternalDepth = N'" + mytable["내경D"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "전지구조")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set 전지구조 = N'" + mytable["전지구조"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "무게(g)")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set Weight = '" + mytable["무게(g)"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "판매가격(￥)")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [Price(￥)] = N'" + mytable["판매가격(￥)"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "PDF도면")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set BluePrintPDF = N'" + mytable["PDF도면"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "3D PDF")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [3DBluePrint] = N'" + mytable["3D PDF"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "DXF도면")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set DXFBluePrint = N'" + mytable["DXF도면"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "DWG도면")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set DWGBluePrint = N'" + mytable["DWG도면"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "주문")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set OrderLink = N'" + mytable["주문"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "색상")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set Color = N'" + mytable["색상"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "경사각")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [경사각] = N'" + mytable["경사각"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "수지재질")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [수지재질] = N'" + mytable["수지재질"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "단자대극수")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [단자대극수] = N'" + mytable["단자대극수"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "정격")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [정격] = N'" + mytable["정격"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "적합전선")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [적합전선] = N'" + mytable["적합전선"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "부속품")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [부속품] = N'" + mytable["부속품"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "상하커버색상/패널색상")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [상하커버색상/패널색상] = N'" + mytable["상하커버색상/패널색상"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "프레임의도장")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [프레임의도장] = N'" + mytable["프레임의도장"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "프레임의 도장")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [프레임의도장] = N'" + mytable["프레임의 도장"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "표면처리")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [표면처리] = N'" + mytable["표면처리"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "투명커버유무")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [투명커버유무] = N'" + mytable["투명커버유무"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "유형")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [유형] = N'" + mytable["유형"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "보강유무")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [보강유무] = N'" + mytable["보강유무"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "두께")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [두께] = N'" + mytable["두께"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "설치나사개수")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [설치나사개수] = N'" + mytable["설치나사개수"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "커넥터외형치수")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [커넥터외형치수] = N'" + mytable["커넥터외형치수"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "커넥터 외형치수")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [커넥터외형치수] = N'" + mytable["커넥터 외형치수"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "적합케이블외경")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [적합케이블외경] = N'" + mytable["적합케이블외경"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "적합 케이블 직경mm")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [적합 케이블 직경mm] = N'" + mytable["적합 케이블 직경mm"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "단자극수")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [단자극수] = N'" + mytable["단자극수"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "정격전압")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [정격전압] = N'" + mytable["정격전압"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "정격전류")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [정격전류] = N'" + mytable["정격전류"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "적합연결전선")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [적합연결전선] = N'" + mytable["적합연결전선"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "치수(폭×길이×두께)")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [치수(폭x길이x두께)] = N'" + mytable["치수(폭×길이×두께)"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "준수플라스틱케이스")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [준수플라스틱케이스] = N'" + mytable["준수플라스틱케이스"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "준수 플라스틱 케이스")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [준수플라스틱케이스] = N'" + mytable["준수 플라스틱 케이스"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "판매단위")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [판매단위] = N'" + mytable["판매단위"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "손잡이길이")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [손잡이길이] = N'" + mytable["손잡이길이"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "지원설치판두께")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [지원설치판두께] = N'" + mytable["지원설치판두께"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "지원설치판 두께")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [지원설치판두께] = N'" + mytable["지원설치판 두께"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "길이치수(mm)")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [길이치수(mm)] = N'" + mytable["길이치수(mm)"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "전체길이치수")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [길이치수(mm)] = N'" + mytable["전체길이치수"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "나사크기")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [나사크기] = N'" + mytable["나사크기"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "나사길이")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [나사길이] = N'" + mytable["나사길이"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "스트랩길이mm")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [스트랩길이mm] = N'" + mytable["스트랩길이mm"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "스트랩길이")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [스트랩길이mm] = N'" + mytable["스트랩길이"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "모양 유형")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [모양 유형] = N'" + mytable["모양 유형"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "크기(폭×길이×두께)mm")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [크기(폭×길이×두께)mm] = N'" + mytable["크기(폭×길이×두께)mm"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "높이치수")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [높이치수] = N'" + mytable["높이치수"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "준수상자")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [준수상자] = N'" + mytable["준수상자"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "나사규격")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [나사규격] = N'" + mytable["나사규격"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "취부구멍")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [취부구멍] = N'" + mytable["취부구멍"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "외경")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [외경] = N'" + mytable["외경"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "필터 부 유효경")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [필터 부 유효경] = N'" + mytable["필터 부 유효경"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "노출부 외형 크기")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [노출부 외형 크기] = N'" + mytable["노출부 외형 크기"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "방수보호등급")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [방수보호등급] = N'" + mytable["방수보호등급"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "노출부 외형 치수(폭×깊이×두께)")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [노출부 외형 치수(폭×깊이×두께)] = N'" + mytable["노출부 외형 치수(폭×깊이×두께)"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "부착판 폭×높이")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [부착판 폭×높이] = N'" + mytable["부착판 폭×높이"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "전체길이치수")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [전체길이치수] = N'" + mytable["전체길이치수"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "고무재질")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [고무재질] = N'" + mytable["고무재질"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "적용나사 크기")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [적용나사 크기] = N'" + mytable["적용나사 크기"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "적합케이스")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [적합케이스] = N'" + mytable["적합케이스"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "별매품 용도")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [별매품 용도] = N'" + mytable["별매품 용도"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "상단패널색상")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [상단패널색상] = N'" + mytable["상단패널색상"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                    else if (items.Key == "준수 팬")
                    {
                        SqlCommand Cmd4 = new SqlCommand("", cn);
                        Cmd4.CommandText = "update TakachiProduct set [준수 팬] = N'" + mytable["준수 팬"] + "' where ModelNumber = N'" + mytable["형번"] + "'";
                        Cmd4.ExecuteNonQuery();
                    }
                }
            }
        }




    }
}
