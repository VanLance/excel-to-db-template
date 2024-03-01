using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace ExcelToSqlServer
{
    public class ExcelToSqlServerImporter
    {
        public void ImportDataFromExcel(string excelFilePath, string connectionString)
        {
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Worksheet worksheet = workbook.Sheets[1]; // Assuming data is in the first sheet

            Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;
            int columnCount = usedRange.Columns.Count;

            List<string[]> data = new List<string[]>();

            // Read data from Excel
            for (int i = 1; i <= rowCount; i++)
            {
                string[] rowData = new string[columnCount];
                for (int j = 1; j <= columnCount; j++)
                {
                    Range cell = usedRange.Cells[i, j];
                    rowData[j - 1] = cell.Value != null ? cell.Value.ToString() : "";
                }
                data.Add(rowData);
            }

            // Import data into SQL Server
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                foreach (var row in data)
                {
                    string query = "INSERT INTO YourTableName (Column1, Column2, Column3) VALUES (@Column1, @Column2, @Column3)"; // Modify as per your table schema
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Column1", row[0]); // Assuming your Excel has three columns
                        command.Parameters.AddWithValue("@Column2", row[1]);
                        command.Parameters.AddWithValue("@Column3", row[2]);
                        command.ExecuteNonQuery();
                    }
                }
            }

            // Cleanup
            workbook.Close(false);
            excelApp.Quit();
            ReleaseObject(worksheet);
            ReleaseObject(workbook);
            ReleaseObject(excelApp);
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
