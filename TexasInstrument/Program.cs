using HtmlAgilityPack;
using System;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.IO;
using Azure.Storage.Blobs;

namespace TexasInstrument
{


    class Program
    {
    static void Main(string[] args)
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
                SqlDataReader di;
                string sql2 = "select TOP 1 * from savedmasterid where basemfgno = 5737 order by masterid desc";
                mid.CommandText = sql2;
                di = mid.ExecuteReader();
                if (di.HasRows)
                {
                    di.Read();
                    savedmid = (int)di["masterid"];
                }

                SqlCommand Cmd = new SqlCommand("", cn);
                SqlDataReader dr;
                String sql = "Select masterid, mn_code, BaseMfgNo from partlist_master where BaseMfgNo = 5737 and masterid > " + savedmid + "";
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
                            //https://www.ti.com/store/ti/en/p/product/?p=SN74LS590N&keyMatch=SN74LS590N&tisearch=search-everything
                            //web.UserAgent = "definitely-not-a-screen-scraper";
                            string weblink = "https://www.ti.com/store/ti/en/p/product/?p=" + dr["mn_code"].ToString();
                            HtmlDocument doc = web.Load(weblink);
                            //var search = doc.DocumentNode.SelectSingleNode("//a[@slot='trigger']");
                            //var search = doc.DocumentNode.SelectSingleNode("//text()[contains(., 'Level-3-260C-168 HR')]");
                            //var path = doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/main[1]/div[1]/div[1]/div[1]/div[2]/section[2]/div[1]/div[1]/div[1]/table[1]/tr[4]/td[1]").InnerText;

                            //status
                            int stat = 0;
                            var status = doc.DocumentNode.SelectSingleNode("//a[@class='u-margin-top-xs']").InnerText;
                            if (status == "ACTIVE")
                            {
                                stat = 1;
                            }

                            //description추출
                            var dsc = doc.DocumentNode.SelectSingleNode("//div[@class='product-header-info']/h2[1]").InnerText.Replace("'", "''");
                            //dsc = System.Net.WebUtility.HtmlDecode(dsc);

                            //Category 추출
                            var category1 = "";
                            var category2 = "";
                            var category3 = "";
                            var category4 = "";
                            var category5 = "";
                            var basemncode = "";
                            var basemncodehref = "";

                            int count = doc.DocumentNode.SelectNodes("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section").Count;

                            if (count == 3)
                            {
                                category1 = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[2]/a[1]").InnerText.Replace("'", "''"));
                                basemncode = doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[3]/a[1]").InnerText.Replace("'", "''");
                                basemncodehref = doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[3]/a[1]").Attributes["href"].Value;
                            }
                            else if (count == 4)
                            {
                                category1 = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[2]/a[1]").InnerText.Replace("'", "''"));
                                category2 = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[3]/a[1]").InnerText.Replace("'", "''"));
                                basemncode = doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[4]/a[1]").InnerText.Replace("'", "''");
                                basemncodehref = doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[4]/a[1]").Attributes["href"].Value;

                            }
                            else if (count == 5)
                            {
                                category1 = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[2]/a[1]").InnerText.Replace("'", "''"));
                                category2 = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[3]/a[1]").InnerText.Replace("'", "''"));
                                category3 = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[4]/a[1]").InnerText.Replace("'", "''"));
                                basemncode = doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[5]/a[1]").InnerText.Replace("'", "''");
                                basemncodehref = doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[5]/a[1]").Attributes["href"].Value;

                            }
                            else if (count == 6)
                            {
                                category1 = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[2]/a[1]").InnerText.Replace("'", "''"));
                                category2 = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[3]/a[1]").InnerText.Replace("'", "''"));
                                category3 = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[4]/a[1]").InnerText.Replace("'", "''"));
                                category4 = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[5]/a[1]").InnerText.Replace("'", "''"));
                                basemncode = doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[6]/a[1]").InnerText.Replace("'", "''");
                                basemncodehref = doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[6]/a[1]").Attributes["href"].Value;

                            }
                            else if (count == 7)
                            {
                                category1 = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[2]/a[1]").InnerText.Replace("'", "''"));
                                category2 = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[3]/a[1]").InnerText.Replace("'", "''"));
                                category3 = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[4]/a[1]").InnerText.Replace("'", "''"));
                                category4 = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[5]/a[1]").InnerText.Replace("'", "''"));
                                category5 = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[6]/a[1]").InnerText.Replace("'", "''"));
                                basemncode = doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[7]/a[1]").InnerText.Replace("'", "''");
                                basemncodehref = doc.DocumentNode.SelectSingleNode("//ti-breadcrumb[@data-lid='breadcrumb']/ti-breadcrumb-section[7]/a[1]").Attributes["href"].Value;
                            }

                            //packagetype, packing, packingqty, compliance, temperature
                            var package = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//table[@class='table-info table-vertical']/tbody[1]/tr[1]/td[1]/a[1]").InnerText);
                            var pack = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//table[@class='table-info table-vertical']/tbody[1]/tr[2]/td[1]/a[1]").InnerText.Replace("\r", "").Replace("\t", "").Replace("\n", ""));
                            var packingqty1 = pack.Substring(0, pack.IndexOf("|")-1).Replace(",","");
                            int packingqty = Convert.ToInt32(packingqty1);
                            var packing = pack.Substring(pack.IndexOf("|") + 2);
                            var temperature = System.Net.WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//table[@class='table-info table-vertical']/tbody[1]/tr[3]/td[1]").InnerText).Replace("\n", "").Replace("\t", "");
                            var rohs = doc.DocumentNode.SelectSingleNode("//table[@class='table-vertical table-info']/tr[1]/th[1]").InnerText.Replace("\n","").Replace("\t","");
                            var rohs2 = doc.DocumentNode.SelectSingleNode("//table[@class='table-vertical table-info']/tr[1]/td[1]").InnerText.Replace("\n", "").Replace("\t", "");
                            var REACH = doc.DocumentNode.SelectSingleNode("//table[@class='table-vertical table-info']/tr[2]/th[1]").InnerText.Replace("\n", "").Replace("\t", "");
                            var REACH2 = doc.DocumentNode.SelectSingleNode("//table[@class='table-vertical table-info']/tr[2]/td[1]").InnerText.Replace("\n", "").Replace("\t", "");
                            var compliance = "";
                            if (rohs == "RoHS" && rohs2 == "Yes")
                            {
                                compliance = rohs;
                                if (REACH == "REACH" && REACH2 == "Yes")
                                {
                                    compliance = rohs + "/" + REACH;
                                }
                            }
                            else if (REACH == "REACH" && REACH2 == "Yes")
                            {
                                compliance = REACH;
                            }

                                //price
                                var pricing = "";
                                var pricecount = doc.DocumentNode.SelectNodes("//table[@id='priceTable']/tbody[1]/tr").Count;
                                int y = 1;
                                while (y <= pricecount/2)
                                {
                                    var pricenum = doc.DocumentNode.SelectSingleNode("//table[@id='priceTable']/tbody[1]/tr[" + y + "]/td[1]").InnerText.Replace("\n", "").Replace("\t", "");
                                    var pricevalue = doc.DocumentNode.SelectSingleNode("//table[@id='priceTable']/tbody[1]/tr[" + y + "]/td[2]/span[1]").InnerText.Replace("\n", "").Replace("\t", "");
                                    pricing = pricing + pricenum + "/$" + pricevalue + " / ";
                                    y++;
                                }


                            //저장
                            SqlCommand Cmd3 = new SqlCommand("", cn);
                            Cmd3.CommandText = "update partlist_master set Dsc = '" + dsc + "', Pricing = '" + pricing + "', WebURL = '" + weblink + "', PackageType = '" + package + "', Packing = '" + packing + "', Compliance = '" + compliance + "', Temperature = '" + temperature + "', packingqty = '" + packingqty + "', BaseMnCode = '" + basemncode + "', PartStatus = " + stat + ", Category1 = '" + category1 + "',  Category2 = '" + category2 + "', Category3 = '" + category3 + "', Category4 = '" + category4 + "', ChangedDate = CONVERT(datetimeoffset, getdate()) AT TIME ZONE 'Korea Standard Time', ChangedBy = ORIGINAL_LOGIN() where masterid =" + dr["masterid"];
                            Cmd3.ExecuteNonQuery();




                            //image link 추출
                            var image = doc.DocumentNode.SelectSingleNode("//img[@class='imageSize']");
                            string imgHref = image.Attributes["src"].Value;
                            string imagename = imgHref.Substring(imgHref.LastIndexOf("/") + 1);

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
                            if (doc.DocumentNode.SelectSingleNode("//div[@class='product-resources']/label[1]").InnerText == "DATA SHEET")
                                {
                                    var datesheet = doc.DocumentNode.SelectSingleNode("//div[@class='product-resources']/ul[1]/li[1]/span[1]/a[1]");
                                    string datesheetHerf = datesheet.Attributes["href"].Value;
                                    string sheetname = datesheetHerf.Substring(datesheetHerf.LastIndexOf("/") + 1) + ".pdf";

                                    //Blob file name
                                    BlobClient blob1 = container2.GetBlobClient(sheetname);

                                    //Store in Local location
                                    HttpWebRequest request1 = (HttpWebRequest)WebRequest.Create(datesheetHerf);
                                    HttpWebResponse response1 = (HttpWebResponse)request1.GetResponse();
                                    Stream inputStream1 = response1.GetResponseStream();
                                    blob1.Upload(inputStream1, true);     //이미 파일이 있는 경우 true로 설정하면 에러를 내지 않고 overwrite하여 날짜와 시간이 바뀐다

                                    SqlCommand Cmd5 = new SqlCommand("", cn);
                                    Cmd5.CommandText = "insert into partlist_master_datasheet values (" + dr["masterid"] + ",'" + datesheetHerf + "', '" + sheetname + "', CONVERT(datetimeoffset, getdate()) AT TIME ZONE 'Korea Standard Time')";
                                    Cmd5.ExecuteNonQuery();
                                }
                            




                            //Product Specification
                            HtmlDocument doc2 = web.Load(basemncodehref);
                            var productlist = doc2.DocumentNode.SelectSingleNode("//meta[@name='PartNumber']");
                            var pnumbers = productlist.Attributes["content"].Value;
                            string[] mncodes = pnumbers.Split(',');
                            mncodes = mncodes.Skip(1).ToArray();

                            foreach (var mncode in mncodes)
                            {
                                SqlCommand Cmd6 = new SqlCommand("", cn);
                                SqlDataReader dx;
                                string sql6 = "";
                                //titlevalue2 = "Testing";
                                sql6 = "select mn_code from partlist_master where mn_code = '" + mncode + "'";
                                Cmd6.CommandText = sql6;
                                dx = Cmd6.ExecuteReader();
                                if (!(dx.HasRows))
                                {
                                    SqlCommand Cmd7 = new SqlCommand("", cn);
                                    SqlDataReader dz;
                                    Cmd7.CommandText = "select MAX(masterid) masterid from partlist_master";
                                    dz = Cmd7.ExecuteReader();
                                    dz.Read();

                                    var masterid = "";
                                    masterid = dz["masterid"].ToString();
                                    int masterid2 = Convert.ToInt32(masterid);
                                    dz.Close();
                                    masterid2 += 1;

                                    Cmd7.CommandText = "insert into partlist_master values(" + masterid2 + ", '" + mncode + "', 5737, NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL)";
                                    Cmd7.ExecuteNonQuery();
                                }
                            }

                                int titlecount = doc2.DocumentNode.SelectNodes("//ti-tab-panel[@tab-title='Parameters']/ti-view-more[1]/div[1]/ti-multicolumn-list[1]/ti-multicolumn-list-row").Count;

                                string[] name = new string[titlecount];
                                string[] value = new string[titlecount];
                                int j = 0;

                                while (j <= titlecount - 1)
                                {
                                    name[j] = doc2.DocumentNode.SelectSingleNode("//ti-tab-panel[@tab-title='Parameters']/ti-view-more[1]/div[1]/ti-multicolumn-list[1]/ti-multicolumn-list-row["+(j+1)+ "]/ti-multicolumn-list-cell[1]/span[1]").InnerText.Replace("\n", "").Replace("\t", "").Replace("'","''");
                                    value[j] = doc2.DocumentNode.SelectSingleNode("//ti-tab-panel[@tab-title='Parameters']/ti-view-more[1]/div[1]/ti-multicolumn-list[1]/ti-multicolumn-list-row[" + (j + 1) + "]/ti-multicolumn-list-cell[2]/span[1]").InnerText.Replace("\n", "").Replace("\t", "").Replace("'", "''");
                                    j++;
                                }

                                int se = 0;
                                SqlCommand Cmd0 = new SqlCommand("", cn);
                                foreach (string specEN in name)
                                {
                                    if (!String.IsNullOrEmpty(name[se]))
                                    {
                                        Cmd0.CommandText = "insert into SpecConnections values (" + dr["masterid"] + ",'" + name[se] + "','" + value[se] + "')";
                                        Cmd0.ExecuteNonQuery();
                                    }
                                    se++;
                                }






                                Console.WriteLine("masterid = {0}   Mn_code = {1}", dr["masterid"], dr["mn_code"]);
                            //System.Threading.Thread.Sleep(25000);     //시간 지연 25초

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
                            Console.WriteLine("masterid = {0}   Mn_code = {1}", dr["masterid"], dr["mn_code"]);
                            System.Threading.Thread.Sleep(25000);     //시간 지연 25초


                        }
                        xx++;


                    }
                }

            }
            catch (Exception ex)
            {
                    SqlCommand save = new SqlCommand("", cn);
                    save.CommandText = "insert into savedmasterid values(5737, " + pmasterid + ", ORIGINAL_LOGIN(), CONVERT(datetimeoffset, getdate()) AT TIME ZONE 'Korea Standard Time')";
                    save.ExecuteNonQuery();
                    Console.WriteLine("exmsg=" + ex.Message);

            }
        }
    }
}
}
