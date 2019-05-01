using System;
using System.Data;
using System.Data.Odbc;

namespace TOMSSQL
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelParser parser = new ExcelParser();
            DataSet ds = parser.GetDataSetFromExcel();
            DataRowCollection drc = ds.Tables[0].Rows;

            using (OdbcConnection conn = parser.CreateConnectionToMSSQL())
            {
                try
                {
                    conn.Open();
                    parser.CreateTablesInMSSQL(conn);
                    parser.ProcessRows(conn, drc);
                    Console.WriteLine("Done.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

        }
    }
}
