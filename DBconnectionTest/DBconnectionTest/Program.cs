using System;
using System.ComponentModel;
using System.Data.SqlClient;

namespace DBconnectionTest
{
    internal class Program
    {
        static void Main(string[] args)
        {

            String _connString = System.Configuration.ConfigurationManager.ConnectionStrings["sam"].ToString();

            using (SqlConnection cn = new SqlConnection(_connString))
            {
                int MasterId = 0;
                string ConvertedMn_code = "", Dsc ="";
                try
                {
                    String sql;
                    cn.Open();

                    SqlCommand Cmd = new SqlCommand("", cn);
                    SqlDataReader dr;
                    sql = "Select masterid, source, descript From exam_item";
                    Cmd.CommandText = sql;
                    dr = Cmd.ExecuteReader();

                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            Dsc = dr["Descript"].ToString();
                            MasterId = (int)dr["masterid"];
                            //ConvertedMn_code = StringConverter.SpecialCharacter(dr["mn_code"].ToString());
                            //Console.WriteLine("MasterId=" + MasterId);
                            SqlCommand Cmd2 = new SqlCommand("", cn);
                            Cmd2.CommandText = "UPDATE partlist_master SET ConvertedMnCode = '" + ConvertedMn_code + "' where masterid = " + MasterId;
                            Cmd2.ExecuteNonQuery();
                        }
                        dr.Close();
                    }
                }
                catch (Exception ex)
                {
                    //Console.WriteLine("MasterId= {0} exmsg= {1}" + MasterId, ex.Message);
                    //DateTime date1 = DateTime.Now;
                    //DateTime utcDate1 = date1.ToUniversalTime();
                    //DateTime date2 = utcDate1.ToLocalTime();
                    //SqlCommand Cmd3 = new SqlCommand("", cn);
                    //Cmd3.CommandText = "Insert into UpdateHistoryStringConverterPartlistMaster values(" + MasterId + "," + ex.Message + "," + date2 + ")";
                    //Cmd3.ExecuteNonQuery();
                    //Environment.Exit(0);
                }
            }
        }
    }
}
