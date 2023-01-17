using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Http;
using Nancy;
using Nancy.Json;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Xml.Linq;

namespace SampleSpreadSheetApp.model
{
    public class SpreadSheetDataLayer
    {
        SqlConnection scon = new SqlConnection(Startup.workflowDB);
        SqlCommand cmd;
        SqlDataAdapter sda;
        DataTable dt = new DataTable();
        HttpContext _httpContext;

        //public List<SpreadSheetModel> GetParentSheetData()
        //{
        //    List<SpreadSheetModel> result = new List<SpreadSheetModel>();
        //    try
        //    {
        //        sda = new SqlDataAdapter("sp_GetParentSheetData", scon);
        //        scon.Open();
        //        sda.Fill(dt);
        //        if (dt != null)
        //        {
        //            foreach (DataRow row in dt.Rows)
        //            {
        //                result.Add(new SpreadSheetModel
        //                {
        //                    name = row["Name"].ToString(),
        //                    monday = Convert.ToInt32(row["Monday"]),
        //                    tuesday = Convert.ToInt32(row["Tuesday"]),
        //                    wednesday = Convert.ToInt32(row["Wednesday"]),
        //                    thursday = Convert.ToInt32(row["Thursday"]),
        //                    friday = Convert.ToInt32(row["Friday"]),
        //                    saturday = Convert.ToInt32(row["Saturday"]),
        //                    sunday = Convert.ToInt32(row["Sunday"])
        //                });
        //            }
        //        }
        //        scon.Close();

        //    }
        //    catch (Exception ex)
        //    {
        //        if (scon.State == ConnectionState.Open)
        //        {
        //            scon.Close();
        //        }
        //    }
        //    return result;
        //}
        public string GetParentSheetData()
        {
            string sql = "Select * from SpreadSheetInfo";
            List<DataTable> sheetData = new List<DataTable>();
            List<string> names = new List<string>();
            Dictionary<string, string> keyValuePairs = new Dictionary<string, string>();
            DataTable dtList = new DataTable();
            DataTable result = new DataTable();
            string JSONresult = null;
            try
            {
                sda = new SqlDataAdapter(sql, scon);
                scon.Open();
                sda.Fill(dt);
                if (dt != null)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
                        sheetData.Add((DataTable)JsonConvert.DeserializeObject(row["spreadsheetData"].ToString(), typeof(DataTable)));
                        names.Add(row["name"].ToString());
                    }
                    for (int i = 0; i < names.Count; i++)
                    {
                        sheetData[i].Columns.Add("Name").SetOrdinal(0);
                        sheetData[i].Columns.RemoveAt(1);
                        foreach (DataRow row1 in sheetData[i].Rows)
                        {
                            row1["Name"] = names[i];
                        }
                    }
                    foreach (DataTable table in sheetData)
                    {
                        dtList.Merge(table, false, MissingSchemaAction.Add);
                    }
                    result = dtList.AsEnumerable().GroupBy(g => new { Col1 = g["Name"] }).Select(g => g.OrderBy(r => r["Name"]).Last()).CopyToDataTable();

                    DataRow newBlankRow1 = result.NewRow();
                    result.Rows.Add(newBlankRow1);
                    result.Rows[result.Rows.Count - 1][0] = "Total";
                    for (int i = 1; i < result.Columns.Count; i++)
                    {
                        int sum = 0;
                        for (int j = 0; j < result.Rows.Count - 1; j++)
                        {
                            var value = result.Rows[j][i].ToString();
                            int val = (value == "" || value == null) ? 0 : Convert.ToInt32(value); 
                            sum = sum + Convert.ToInt32(val);
                        }
                        result.Rows[result.Rows.Count - 1][i] = sum;
                    }
                    //DataTable dtUnion = sheetData[0].AsEnumerable().Union(sheetData[1].AsEnumerable()).CopyToDataTable<DataRow>();
                    
                    JSONresult = JsonConvert.SerializeObject(result);

                }
                scon.Close();   
            }
            catch (Exception ex)
            {
                if (scon.State == ConnectionState.Open)
                {
                    scon.Close();
                }
            }
            return JSONresult;
        }
        public void InsertChild1Data(List<SpreadSheetModel> model, string Name)
        {
            try
            {
                if(model != null)
                {
                    model.ForEach(x =>
                    {
                        if (x.Products != "Total")
                        {
                            cmd = new SqlCommand("sp_InsertChildData", scon);
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@Products", SqlDbType.VarChar);
                            cmd.Parameters["@Products"].Value = x.Products;
                            cmd.Parameters.Add("@Monday", SqlDbType.VarChar);
                            cmd.Parameters["@Monday"].Value = x.monday;
                            cmd.Parameters.Add("@Tuesday", SqlDbType.VarChar);
                            cmd.Parameters["@Tuesday"].Value = x.tuesday;
                            cmd.Parameters.Add("@Wednesday", SqlDbType.VarChar);
                            cmd.Parameters["@Wednesday"].Value = x.wednesday;
                            cmd.Parameters.Add("@Thursday", SqlDbType.VarChar);
                            cmd.Parameters["@Thursday"].Value = x.thursday;
                            cmd.Parameters.Add("@friday", SqlDbType.VarChar);
                            cmd.Parameters["@friday"].Value = x.friday;
                            cmd.Parameters.Add("@saturday", SqlDbType.VarChar);
                            cmd.Parameters["@saturday"].Value = x.saturday;
                            cmd.Parameters.Add("@sunday", SqlDbType.VarChar);
                            cmd.Parameters["@sunday"].Value = x.sunday;
                            cmd.Parameters.Add("@Name", SqlDbType.VarChar);
                            cmd.Parameters["@Name"].Value = Name;
                            scon.Open();
                            cmd.ExecuteNonQuery();
                            scon.Close();
                        }
                    });
                }

            }
            catch (Exception ex)
            {
                if (scon.State == ConnectionState.Open)
                {
                    scon.Close();
                }
            }
        }
        public void InsertChild2Data(List<SpreadSheetModel> model, string Name)
        {
            try
            {
                if (model != null)
                {
                    model.ForEach(x =>
                    {
                        if (x.Products != "Total")
                        {
                            cmd = new SqlCommand("sp_InsertChildData", scon);
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@Products", SqlDbType.VarChar);
                            cmd.Parameters["@Products"].Value = x.Products;
                            cmd.Parameters.Add("@Monday", SqlDbType.VarChar);
                            cmd.Parameters["@Monday"].Value = x.monday;
                            cmd.Parameters.Add("@Tuesday", SqlDbType.VarChar);
                            cmd.Parameters["@Tuesday"].Value = x.tuesday;
                            cmd.Parameters.Add("@Wednesday", SqlDbType.VarChar);
                            cmd.Parameters["@Wednesday"].Value = x.wednesday;
                            cmd.Parameters.Add("@Thursday", SqlDbType.VarChar);
                            cmd.Parameters["@Thursday"].Value = x.thursday;
                            cmd.Parameters.Add("@friday", SqlDbType.VarChar);
                            cmd.Parameters["@friday"].Value = x.friday;
                            cmd.Parameters.Add("@saturday", SqlDbType.VarChar);
                            cmd.Parameters["@saturday"].Value = x.saturday;
                            cmd.Parameters.Add("@sunday", SqlDbType.VarChar);
                            cmd.Parameters["@sunday"].Value = x.sunday;
                            cmd.Parameters.Add("@Name", SqlDbType.VarChar);
                            cmd.Parameters["@Name"].Value = Name;
                            scon.Open();
                            cmd.ExecuteNonQuery();
                            scon.Close();
                        }
                    });
                }

            }
            catch (Exception ex)
            {
                if (scon.State == ConnectionState.Open)
                {
                    scon.Close();
                }
            }
        }


        public List<SpreadSheetModel> GetChild1SheetData()
        {
            List<SpreadSheetModel> result = new List<SpreadSheetModel>();
            try
            {
                sda = new SqlDataAdapter("sp_GetChild1SheetData", scon);
                scon.Open();
                sda.Fill(dt);
                if (dt != null)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        result.Add(new SpreadSheetModel
                        {
                            //Name = row["Name"].ToString(),
                            Products = row["Products"].ToString(),
                            monday = Convert.ToInt32(row["Monday"]),
                            tuesday = Convert.ToInt32(row["Tuesday"]),
                            wednesday = Convert.ToInt32(row["Wednesday"]),
                            thursday = Convert.ToInt32(row["Thursday"]),
                            friday = Convert.ToInt32(row["Friday"]),
                            saturday = Convert.ToInt32(row["Saturday"]),
                            sunday = Convert.ToInt32(row["Sunday"])
                        });
                    }
                }
                scon.Close();

            }
            catch (Exception ex)
            {
                if (scon.State == ConnectionState.Open)
                {
                    scon.Close();
                }
            }
            return result;
        }

        public List<SpreadSheetModel> GetChild2SheetData()
        {
            List<SpreadSheetModel> result = new List<SpreadSheetModel>();
            try
            {
                sda = new SqlDataAdapter("sp_GetChild2SheetData", scon);
                scon.Open();
                sda.Fill(dt);
                if (dt != null)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        result.Add(new SpreadSheetModel
                        {
                            //Name = row["Name"].ToString(),
                            Products = row["Products"].ToString(),
                            monday = Convert.ToInt32(row["Monday"]),
                            tuesday = Convert.ToInt32(row["Tuesday"]),
                            wednesday = Convert.ToInt32(row["Wednesday"]),
                            thursday = Convert.ToInt32(row["Thursday"]),
                            friday = Convert.ToInt32(row["Friday"]),
                            saturday = Convert.ToInt32(row["Saturday"]),
                            sunday = Convert.ToInt32(row["Sunday"])
                        });
                    }
                }
                scon.Close();

            }
            catch (Exception ex)
            {
                if (scon.State == ConnectionState.Open)
                {
                    scon.Close();
                }
            }
            return result;
        }

        public void InsertDealerDetails(string name, string sheetInfo)
        {
           try
            {
                cmd = new SqlCommand("sp_InsertDealerData", scon);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@name", SqlDbType.VarChar);
                cmd.Parameters["@name"].Value = name;
                cmd.Parameters.Add("@spreadsheetData", SqlDbType.VarChar);
                cmd.Parameters["@spreadsheetData"].Value = sheetInfo;
                scon.Open();
                cmd.ExecuteNonQuery();
                scon.Close();
            }
            catch (Exception ex)
            {
                if (scon.State == ConnectionState.Open)
                {
                    scon.Close();
                }
            }
        }
    }
}
