using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.SqlClient;
using OfficeOpenXml;
using System.IO;
using ConsoleTableExt;

namespace Excel_Reader
{
    internal class Program
    {
        static string conString = ConfigurationManager.AppSettings.Get("conString");
        static string fileString = ConfigurationManager.AppSettings.Get("fileString");
        static string dbString = ConfigurationManager.AppSettings.Get("dbString");
        
        static void Main(string[] args)
        {           
            // Checks if the Database exists, if it does remove it and create a fresh one.
            using (var con = new SqlConnection(dbString))
            {
                using (var cmd = con.CreateCommand())
                {
                    con.Open();
                    // checks if the database is in existence.
                    cmd.CommandText = "SELECT db_id('ExcelBook')";
                    bool state = cmd.ExecuteScalar() != DBNull.Value;
                    
                    if (state)
                    {
                        cmd.CommandText = "DROP DATABASE ExcelBook";
                        cmd.ExecuteNonQuery();
                        Console.WriteLine("Removing Previous Database...");
                    }
                    
                    cmd.CommandText = "CREATE DATABASE ExcelBook";
                    cmd.ExecuteNonQuery();
                    Console.WriteLine("Database Created...");

                }
            }
            // Creates a table in the database to be used.
            using (var con = new SqlConnection(conString))
            {
                using (var cmd = con.CreateCommand())
                {
                    con.Open();
                    cmd.CommandText = "IF NOT EXISTS (SELECT * FROM sysobjects WHERE id = object_id(N'[dbo].[ExcelSheet]')" +
                                            "AND OBJECTPROPERTY(id, N'IsUserTable') = 1)" +
                                            "CREATE TABLE[dbo].[ExcelSheet] (" +
                                            "Date VARCHAR(50)," +
                                            "Product VARCHAR(50)," +
                                            "Quantity INT," +
                                            "Customer VARCHAR(50));";
                    cmd.ExecuteNonQuery();
                    Console.WriteLine("Database Table Created...");
                }
            }
            //creates an ExcelPackage from the XLSX file.
            using (ExcelPackage package = new ExcelPackage(new FileInfo(fileString)))
            {
                // defines which sheet in the workbook to pull data from.
                var sheet = package.Workbook.Worksheets["data"];
                // returns a list of objects from the Excel worksheet.
                var orders = new Program().GetList<Order>(sheet);
                // inserts all the objects and their data into the SQL database.
                InsertDatabaseTable(orders);
                // fetchs the database records and prints them to the screen.
                Console.WriteLine("Displaying Database results... ");
                ConsoleTableBuilder.From(GetDatabaseTable()).ExportAndWriteLine();
                Console.ReadKey();
            }
        }

        private List<T> GetList<T>(ExcelWorksheet sheet)
        {
            Console.WriteLine("Reading from Excel Worksheet...");
            List<T> list = new List<T>();
            // first row is for knowing the properties of object
            var columnInfo = Enumerable.Range(1, sheet.Dimension.Columns).ToList().Select(n => new {Index=n,ColumnName=sheet.Cells[1,n].Value.ToString()});

            for(int row=2; row < sheet.Dimension.Rows; row++)
            {
                T obj = (T)Activator.CreateInstance(typeof(T)); //generic object
                foreach(var prop in typeof(T).GetProperties())
                {
                    int col = columnInfo.SingleOrDefault(c => c.ColumnName == prop.Name).Index;
                    var val = sheet.Cells[row, col].Value;
                    var propType = prop.PropertyType;
                    prop.SetValue(obj, Convert.ChangeType(val, propType));
                }
                list.Add(obj);
            }

            return list;
        }
        // puts all the data into the SQL database from the excel package.
        public static void InsertDatabaseTable(List<Order> orders)
        {
            Console.WriteLine("Seeding into the database...");
            using (var con = new SqlConnection(conString))
            {
                con.Open();

                foreach (var order in orders)
                {
                    using (var cmd = con.CreateCommand())
                    {
                        cmd.CommandText = "INSERT INTO ExcelSheet (Date, Product, Quantity, Customer) VALUES (@date, @product, @quantity, @customer)";
                        cmd.Parameters.AddWithValue("@date", order.Date);
                        cmd.Parameters.AddWithValue("@product", order.Product);
                        cmd.Parameters.AddWithValue("@quantity", order.Quantity);
                        cmd.Parameters.AddWithValue("@customer", order.Customer);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }
        // fetchs all the Database records.
        public static List<Order> GetDatabaseTable()
        {
            List<Order> tableData = new List<Order>();

            using (var con = new SqlConnection(conString))
            {
                using (var cmd = con.CreateCommand())
                {
                    con.Open();
                    cmd.CommandText = "SELECT * FROM ExcelSheet";

                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                tableData.Add(new Order
                                {
                                    Date = reader.GetString(0),
                                    Product = reader.GetString(1),
                                    Quantity = reader.GetInt32(2),
                                    Customer = reader.GetString(3)
                                });
                            }
                        }
                        else
                        {
                            Console.WriteLine(" No Rows to Display.");
                        }
                    }
                }
            }
            return tableData;
        }
    }
    // Model class
    public class Order
    {
        public string Date { get; set; }

        public string Product { get; set; }

        public int Quantity { get; set; }

        public string Customer { get; set; }
    }
}
