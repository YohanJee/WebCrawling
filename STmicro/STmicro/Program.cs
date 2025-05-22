using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Data.SqlClient;
using System.Net;
using System.Text;
using System.IO;
using Azure.Storage.Blobs;

namespace STmicro
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;  //https접속에러를 막는 프로토콜

            HtmlWeb web = new HtmlWeb();

            // Create BlobServiceClient from the connection string.
            BlobServiceClient blobServiceClient = new BlobServiceClient("DefaultEndpointsProtocol=https;AccountName=leocomfile;AccountKey=9RzrYKzmdJNgwTh2gxJ+ioYs+/YmbHjY08ee8QizahIYh7W+XyECYrzrIwEBCJIgwqR8JR+69As8E+T7o/BktQ==;EndpointSuffix=core.windows.net");

            // Get and create the container for the blobs
            BlobContainerClient container1 = blobServiceClient.GetBlobContainerClient("image");
            BlobContainerClient container2 = blobServiceClient.GetBlobContainerClient("datasheet");


            int pmasterid = 0;
            String _connString = System.Configuration.ConfigurationManager.ConnectionStrings["sam"].ToString();
            using (SqlConnection cn = new SqlConnection(_connString))
            {
                try
                {
                    cn.Open();

                    var savedmid = 0;
                    SqlCommand mid = new SqlCommand("", cn);
                    SqlDataReader dt;
                    string sql11 = "select TOP 1 * from savedmasterid where basemfgno = 2908 order by masterid desc";
                    mid.CommandText = sql11;
                    dt = mid.ExecuteReader();
                    if (dt.HasRows)
                    {
                        dt.Read();
                        savedmid = (int)dt["masterid"];
                    }

                    SqlCommand Cmd = new SqlCommand("", cn);
                    SqlDataReader dr;
                    String sql = "Select masterid, mn_code, BaseMfgNo from partlist_master where BaseMfgNo = 2908 and masterid >= " + savedmid + "";
                    Cmd.CommandText = sql;
                    dr = Cmd.ExecuteReader();


                    if (dr.HasRows)
                    {
                        int xx = 1;
                        while (dr.Read())
                        {
                            try
                            {
                                pmasterid = (int)dr["masterid"];
                                //https://www.ti.com/store/ti/en/p/product/?p=5962-8680603V2A
                                //https://www.ti.com/product/TL4050C#product-details
                                //var keyword = doc.DocumentNode.SelectSingleNode("//text()[contains(., 'Parametrics')]");

                                string part_number = dr["mn_code"].ToString();
                                if (part_number[1] == '.')
                                {
                                    StringBuilder sb = new StringBuilder(part_number);
                                    sb[1] = '-';
                                    part_number = sb.ToString();
                                }



                                string weblink = "https://estore.st.com/en/" + part_number + "-cpn.html";


                                HtmlDocument doc = web.Load(weblink);
                                //var search = doc.DocumentNode.SelectSingleNode("//div[@class='value']");

                                //var test = doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/section[7]/div[1]/div[2]/div[1]/div[1]/div[3]/div[9]/i[1]");


                                //description추출
                                var dsc = "";
                                var content = doc.DocumentNode.SelectSingleNode("(//script[@type='application/ld+json'])[2]").InnerText;
                                List<string> abc = content.Split('\"').ToList();
                                foreach (var contents in abc)
                                {
                                    if (contents.Contains("description"))
                                    {
                                        dsc = abc[abc.IndexOf(contents) + 2];
                                        break;
                                    }
                                }
                                if(dsc.Length > 200)
                                {
                                    dsc = dsc.Substring(0, 200);
                                }




                                //status
                                int stat = 0;
                                var status = doc.DocumentNode.SelectSingleNode("//h3[@class='active-product-label']").InnerText;
                                if (status == "Active")
                                {
                                    stat = 1;
                                }

                                var weblink2 = doc.DocumentNode.SelectSingleNode("//div[@class='learn-more']/a[1]");
                                string weblink3 = weblink2.Attributes["href"].Value;
                                HtmlDocument doc2 = web.Load(weblink3);
                                //var search = doc2.DocumentNode.SelectSingleNode("//a[@href='https://estore.st.com/en/1-5ke12ca-cpn.html']");
                                var allmncode = doc2.DocumentNode.SelectSingleNode("//meta[@name='keywords']");
                                var everything = allmncode.Attributes["content"].Value;
                                List<string> bcd = everything.Split(',').ToList();
                                foreach (var mncode in bcd.Skip(1))
                                {
                                    var mncode1 = mncode.Substring(1, mncode.Length - 1);
                                    SqlCommand Cmd7 = new SqlCommand("", cn);
                                    SqlDataReader dx;
                                    string sql4 = "";
                                    //titlevalue2 = "Testing";
                                    sql4 = "select mn_code from partlist_master where mn_code = '" + mncode1 + "'";
                                    Cmd7.CommandText = sql4;
                                    dx = Cmd7.ExecuteReader();
                                    if (!(dx.HasRows))
                                    {
                                        SqlCommand Cmd8 = new SqlCommand("", cn);
                                        SqlDataReader dz;
                                        Cmd8.CommandText = "select MAX(masterid) masterid from partlist_master";
                                        dz = Cmd8.ExecuteReader();
                                        dz.Read();

                                        var masterid = "";
                                        masterid = dz["masterid"].ToString();
                                        int masterid2 = Convert.ToInt32(masterid);
                                        dz.Close();
                                        masterid2 += 1;

                                        Cmd8.CommandText = "insert into partlist_master values(" + masterid2 + ", '" + mncode1 + "', 2908, NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL)";
                                        Cmd8.ExecuteNonQuery();
                                    }
                                }





                                //packagetype, packing, temp, compliance
                                var packagetype = "";
                                var packing = "";
                                var compliance = "";
                                string temp_min = "";
                                string temp_max = "";
                                var temp = "";
                                string parameter = "";
                                int count = doc.DocumentNode.SelectNodes("(//table[@class='table table-striped table-bordered'])[2]/tbody[1]/tr").Count;
                                int k = 0;
                                while ((k <= count - 2))
                                {
                                    parameter = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("(//table[@class='table table-striped table-bordered'])[2]/tbody[1]/tr[" + (k + 2) + "]/td[1]").InnerText);
                                    if (parameter.Contains("Temp Min"))
                                    {
                                        temp_min = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("(//table[@class='table table-striped table-bordered'])[2]/tbody[1]/tr[" + (k + 2) + "]/td[2]").InnerText);
                                    }
                                    else if (parameter.Contains("Temp Max"))
                                    {
                                        temp_max = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("(//table[@class='table table-striped table-bordered'])[2]/tbody[1]/tr[" + (k + 2) + "]/td[2]").InnerText);
                                    }
                                    else if (parameter.Contains("Packing Type"))
                                    {
                                        packing = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("(//table[@class='table table-striped table-bordered'])[2]/tbody[1]/tr[" + (k + 2) + "]/td[2]").InnerText);
                                    }
                                    else if (parameter.Contains("Package Name"))
                                    {
                                        packagetype = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("(//table[@class='table table-striped table-bordered'])[2]/tbody[1]/tr[" + (k + 2) + "]/td[2]").InnerText);
                                    }
                                    else if (parameter.Contains("Compliance"))
                                    {
                                        compliance = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("(//table[@class='table table-striped table-bordered'])[2]/tbody[1]/tr[" + (k + 2) + "]/td[2]").InnerText);
                                    }
                                    temp = temp_min + " ~ " + temp_max;
                                    k++;

                                }


                                //category
                                int catcount = doc.DocumentNode.SelectNodes("//ul[@class='items']/li").Count;
                                var category1 = "";
                                var category2 = "";
                                var category3 = "";
                                var category4 = "";
                                var category5 = "";
                                var basemncode = "";

                                if (catcount == 5)
                                {
                                    category1 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[2]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    category2 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[3]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    basemncode = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[4]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                }
                                else if (catcount == 6)
                                {
                                    category1 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[2]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    category2 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[3]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    category3 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[4]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    basemncode = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[5]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                }
                                else if (catcount == 7)
                                {
                                    category1 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[2]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    category2 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[3]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    category3 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[4]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    category4 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[5]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    basemncode = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[6]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                }
                                else if (catcount == 8)
                                {
                                    category1 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[2]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    category2 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[3]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    category3 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[4]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    category4 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[5]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    category5 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[6]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    basemncode = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[7]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                }
                                else if(catcount == 9)
                                {
                                    category1 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[3]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    category2 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[4]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    category3 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[5]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    category4 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[6]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    category5 = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[7]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                    basemncode = doc.DocumentNode.SelectSingleNode("//ul[@class='items']/li[8]/a[1]").InnerText.Replace("\n                    ", "").Replace("                ", "");
                                }

                                //pricing
                                int pricecount = doc.DocumentNode.SelectNodes("//table[@id='tier_price']/tbody[1]/tr").Count;
                                int y = 1;
                                var pricing = "";
                                while (y <= pricecount - 1)
                                {
                                    var quan = doc.DocumentNode.SelectSingleNode("//table[@id='tier_price']/tbody[1]/tr[" + y + "]/td[1]").InnerText;
                                    var price = doc.DocumentNode.SelectSingleNode("//table[@id='tier_price']/tbody[1]/tr[" + y + "]/td[2]").InnerText;
                                    var total = quan + " / " + price + " , ";
                                    pricing = pricing + total;
                                    y++;
                                }

                                //저장
                                SqlCommand Cmd3 = new SqlCommand("", cn);
                                Cmd3.CommandText = "update partlist_master set Pricing = '" + pricing + "', Dsc = '" + dsc + "', Temperature = '" + temp + "', Category1 = '" + category1 + "', Category2 = '" + category2 + "', Category3 = '" + category3 + "', Category4 = '" + category4 + "', Category5 = '" + category5 + "', Packing = '" + packing + "', Compliance = '" + compliance + "', BaseMnCode = '" + basemncode + "', PartStatus = " + stat + ", WebURL = '" + weblink + "', PackageType = '" + packagetype + "', ChangedDate = CONVERT(datetimeoffset, getdate()) AT TIME ZONE 'Korea Standard Time', ChangedBy = ORIGINAL_LOGIN() where masterid =" + dr["masterid"];
                                Cmd3.ExecuteNonQuery();



                                //Spec 펼쳐서 더 길게 보기

                                int speccount = doc.DocumentNode.SelectNodes("//td[@class='col data']/ul[1]/li").Count;
                                string[] name = new string[speccount];
                                string[] value = new string[speccount];
                                string namevalue = "";
                                int j = 0;

                                while (j <= speccount - 1)
                                {
                                    namevalue = doc.DocumentNode.SelectSingleNode("//td[@class='col data']/ul[1]/li[" + (j + 1) + "]").InnerText.Replace("'","''");

                                    if (namevalue.IndexOf(":") > -1)
                                    {
                                        name[j] = namevalue.Substring(0, namevalue.IndexOf(":"));
                                        value[j] = namevalue.Substring(namevalue.IndexOf(":") + 2);
                                        if (name[j].Length > 255)
                                        {
                                            name[j] = "";
                                        }
                                        if (value[j].Length > 255)
                                        {
                                            value[j] = "";
                                        }
                                    }
                                    else if(namevalue.Length > 350)
                                    {
                                        name[j] = "";
                                        value[j] = "";

                                    }
                                    else
                                    {
                                        name[j] = namevalue;
                                    }

                                    Console.WriteLine("{0}\n{1}", name[j], value[j]);
                                    j++;
                                }

                                int se = 0;
                                SqlCommand Cmd2 = new SqlCommand("", cn);
                                foreach (string specEN in name)
                                {
                                    if (!String.IsNullOrEmpty(name[se]))
                                    {
                                        Cmd2.CommandText = "insert into SpecConnections values (" + dr["masterid"] + ",'" + name[se] + "','" + value[se] + "')";
                                        Cmd2.ExecuteNonQuery();
                                    }
                                    se++;
                                }



                                //image link 추출
                                var image = doc.DocumentNode.SelectSingleNode("//img[@class='gallery-placeholder__image']");
                                string imgHref = image.Attributes["src"].Value;
                                string imagename = imgHref.Substring(imgHref.LastIndexOf("/")+1, imgHref.IndexOf("?")-imgHref.LastIndexOf("/")-1);

                                //Blob file name
                                BlobClient blob = container1.GetBlobClient(imagename);

                                //Store in Local location
                                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(imgHref);
                                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                                Stream inputStream = response.GetResponseStream();
                                blob.Upload(inputStream, true);     //이미 파일이 있는 경우 true로 설정하면 에러를 내지 않고 overwrite하여 날짜와 시간이 바뀐다

                                SqlCommand Cmd4 = new SqlCommand("", cn);
                                Cmd4.CommandText = "insert into partlist_master_image values (" + dr["masterid"] + ",'" + imgHref + "', '" + imagename + "', CONVERT(datetimeoffset, getdate()) AT TIME ZONE 'Korea Standard Time')";
                                Cmd4.ExecuteNonQuery();




                                //datasheet 링크 추출
                                var datesheet = doc.DocumentNode.SelectSingleNode("//li[@class='download_link']/a[1]");
                                string datesheetHerf = datesheet.Attributes["href"].Value;
                                string sheetname = datesheetHerf.Substring(datesheetHerf.LastIndexOf("/")+1);

                                //Blob file name
                                BlobClient blob1 = container2.GetBlobClient(sheetname);

                                //Store in Local location
                                HttpWebRequest request1 = (HttpWebRequest)WebRequest.Create(datesheetHerf);
                                HttpWebResponse response1 = (HttpWebResponse)request1.GetResponse();
                                Stream inputStream1 = response1.GetResponseStream();
                                blob1.Upload(inputStream1, true);     //이미 파일이 있는 경우 true로 설정하면 에러를 내지 않고 overwrite하여 날짜와 시간이 바뀐다

                                SqlCommand Cmd5 = new SqlCommand("", cn);
                                Cmd5.CommandText = "insert into partlist_master_datasheet values (" + dr["masterid"] + ",'" + datesheetHerf + "','" + sheetname + "', CONVERT(datetimeoffset, getdate()) AT TIME ZONE 'Korea Standard Time')";
                                Cmd5.ExecuteNonQuery();






                            }

                            catch (Exception ex)
                            {
                                Console.WriteLine("exmsg=" + ex.Message);
                                SqlCommand Cmd6 = new SqlCommand("", cn);
                                Cmd6.CommandText = "update partlist_master set ChangedDate = CONVERT(datetimeoffset, getdate()) AT TIME ZONE 'Korea Standard Time', ScrapError = '" + ex.Message + "' where masterid =" + pmasterid;
                                Cmd6.ExecuteNonQuery();

                                //if (ex.InnerException.Message == "The remote server returned an error: (404) Not Found.")
                                //{
                                //    Cmd6.CommandText = "update partlist_master set ChangedDate = GETDATE(), ScrapError = '" + ex + "' where masterid =" + pmasterid;
                                //}
                                //else
                                //{
                                //    Cmd6.CommandText = "update partlist_master set ChangedDate = GETDATE(), ScrapError = '" + ex + "' where where masterid =" + pmasterid;
                                //}
                                //Cmd6.ExecuteNonQuery();

                            }

                            Console.WriteLine("masterid = {0}   Mn_code = {1}", dr["masterid"], dr["mn_code"]);
                            System.Threading.Thread.Sleep(25000);     //시간 지연 25초

                            xx++;


                        }
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("exmsg=" + ex.Message);
                    SqlCommand save = new SqlCommand("", cn);
                    save.CommandText = "insert into savedmasterid values(2908, " + pmasterid + ", ORIGINAL_LOGIN(), CONVERT(datetimeoffset, getdate()) AT TIME ZONE 'Korea Standard Time')";
                    save.ExecuteNonQuery();
                }
            }
        }
    }
}
