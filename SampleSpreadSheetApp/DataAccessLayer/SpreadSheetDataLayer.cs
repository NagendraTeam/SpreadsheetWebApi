using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace SampleSpreadSheetApp.model
{
    public class SpreadSheetDataLayer
    {
        SqlConnection scon = new SqlConnection(Startup.workflowDB);
        SqlCommand cmd;
        SqlDataAdapter sda;
        DataTable dt = new DataTable();
        HttpContext _httpContext;

        public List<SpreadSheetModel> GetParentSheetData()
        {
            List<SpreadSheetModel> result = new List<SpreadSheetModel>();
            try
            {
                sda = new SqlDataAdapter("sp_GetParentSheetData", scon);
                scon.Open();
                sda.Fill(dt);
                if (dt != null)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        result.Add(new SpreadSheetModel
                        {
                            name = row["Name"].ToString(),
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


    }
}
