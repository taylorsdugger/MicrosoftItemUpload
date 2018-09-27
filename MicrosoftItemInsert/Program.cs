using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace MicrosoftItemInsert
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str;
            int rCnt = 2;
            int cCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\DUGGTAY\Downloads\LaunchSKUs.xlsx", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            Console.WriteLine(rCnt);

            for (rCnt = 2; rCnt <= rw; rCnt++)
            {
                string sku = ((range.Cells[rCnt, 1] as Excel.Range).Value2).ToString();
                string product = ((range.Cells[rCnt, 2] as Excel.Range).Value2).ToString();
                string lob = ((range.Cells[rCnt, 3] as Excel.Range).Value2).ToString();

                bool lobValue = readRecord(sku, "LineOfBusiness");
                if (lobValue)
                {
                    updateRecord(sku, "LineOfBusiness", lob);
                }
                else
                {
                    insertRecord(sku, "LineOfBusiness", lob);
                }

                bool prodValue = readRecord(sku, "Product");
                if (prodValue)
                {
                    updateRecord(sku, "Product", product);
                }
                else
                {
                    insertRecord(sku, "Product", product);
                }


                if (rCnt % 1000 == 0)
                {
                    Console.WriteLine(rCnt);
                }

            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);


            void insertRecord(string sku, string attributeName, string attributeValue)
            {
                using (System.IO.StreamWriter file =
                    new System.IO.StreamWriter(@"C:\Users\DUGGTAY\Downloads\errorLog.txt", true))
                {

                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["LuminosityIntegration"].ConnectionString))
                    {
                        using (SqlCommand command = new SqlCommand())
                        {
                            command.Connection = connection;
                            command.CommandType = CommandType.StoredProcedure;
                            command.CommandText = "InsertTenantMap";
                            command.Parameters.AddWithValue("@TenantID", 1);
                            command.Parameters.AddWithValue("@Key", "SKU");
                            command.Parameters.AddWithValue("@Value", sku);
                            command.Parameters.AddWithValue("@AttributeName", attributeName);
                            command.Parameters.AddWithValue("@AttributeValue", attributeValue);
                            command.Parameters.AddWithValue("@AttributeValueDataType", "String");

                            try
                            {
                                connection.Open();
                                int recordsAffected = command.ExecuteNonQuery();
                            }
                            catch (SqlException e)
                            {
                                file.WriteLine("Error at SKU: " + sku + ". With message: " + e.Message);
                            }
                            finally
                            {
                                connection.Close();
                            }
                        }
                    }
                }

            }

            void updateRecord(string sku, string attributeName, string attributeValue)
            {
                using (System.IO.StreamWriter file =
                    new System.IO.StreamWriter(@"C:\Users\DUGGTAY\Downloads\errorLog.txt", true))
                {

                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["LuminosityIntegration"].ConnectionString))
                    {
                        using (SqlCommand command = new SqlCommand())
                        {
                            command.Connection = connection;
                            command.CommandType = CommandType.StoredProcedure;
                            command.CommandText = "UpdateTenantMap";
                            command.Parameters.AddWithValue("@TenantID", 1);
                            command.Parameters.AddWithValue("@Key", "SKU");
                            command.Parameters.AddWithValue("@Value", sku);
                            command.Parameters.AddWithValue("@AttributeName", attributeName);
                            command.Parameters.AddWithValue("@AttributeValue", attributeValue);
                            command.Parameters.AddWithValue("@AttributeValueDataType", "String");

                            try
                            {
                                connection.Open();
                                int recordsAffected = command.ExecuteNonQuery();
                            }
                            catch (SqlException e)
                            {
                                file.WriteLine("Error at SKU: " + sku + ". With message: " + e.Message);
                            }
                            finally
                            {
                                connection.Close();
                            }
                        }
                    }
                }

            }

            bool readRecord(string sku, string attributeName)
            {
                using (System.IO.StreamWriter file =
                    new System.IO.StreamWriter(@"C:\Users\DUGGTAY\Downloads\errorLog.txt", true))
                {

                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["LuminosityIntegration"].ConnectionString))
                    {
                        using (SqlCommand command = new SqlCommand())
                        {
                            command.Connection = connection;
                            command.CommandType = CommandType.StoredProcedure;
                            command.CommandText = "ReadTenantMap";
                            command.Parameters.AddWithValue("@TenantID", 1);
                            command.Parameters.AddWithValue("@Key", "SKU");
                            command.Parameters.AddWithValue("@Value", sku);
                            command.Parameters.AddWithValue("@AttributeName", attributeName);

                            try
                            {
                                connection.Open();
                                using (var reader = command.ExecuteReader())
                                {
                                    if (reader.Read())
                                    {
                                        return true;
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                }
                            }
                            catch (SqlException e)
                            {
                                file.WriteLine("Error at SKU: " + sku + ". With message: " + e.Message);
                                return false;
                            }
                            finally
                            {
                                connection.Close();
                            }
                        }
                    }
                }

            }

        }

    }
}
