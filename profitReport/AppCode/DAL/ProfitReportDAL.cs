using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace profitReport.AppCode.DAL
{
    public class ProfitReportDAL
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);

        public DataTable getProfitReport(string invoiceYear)
        {
            DataSet ds = new DataSet();
            try
            {
                SqlCommand cmd = new SqlCommand("dbo.sp_selectProfitReport", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@invoiceYear", invoiceYear);               
                SqlDataAdapter adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                cmd.Dispose();
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                ds.Dispose();
            }
            return ds.Tables[0];
        }

        public Int32 deleteAllProfitReport()
        {
            int result;
            try
            {
                SqlCommand cmd = new SqlCommand("dbo.sp_deleteAllProfitReport", conn);
                cmd.CommandType = CommandType.StoredProcedure;

                if (conn.State == ConnectionState.Closed)
                    conn.Open();

                result = cmd.ExecuteNonQuery();
                cmd.Dispose();

                if (result > 0)
                    return result;
                else
                    return 0;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (conn.State != ConnectionState.Closed)
                    conn.Close();
            }
        }
    }
}
