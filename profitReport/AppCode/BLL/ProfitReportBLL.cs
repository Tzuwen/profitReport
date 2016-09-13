using profitReport.AppCode.DAL;
using System;
using System.Data;

namespace profitReport.AppCode.BLL
{
    public class ProfitReportBLL
    {
        public static DataTable getProfitReport(string invoiceYear)
        {
            ProfitReportDAL objDal = new ProfitReportDAL();
            try
            {
                return objDal.getProfitReport(invoiceYear);
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                objDal = null;
            }
        }

        public static Int32 deleteAllProfitReport()
        {
            ProfitReportDAL objDal = new ProfitReportDAL();
            try
            {
                return objDal.deleteAllProfitReport();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                objDal = null;
            }
        }
    }
}
