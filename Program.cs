using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.IO;
//using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
namespace ConsoleApplication13
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = Directory.GetCurrentDirectory();
            Console.WriteLine("Dates should be between 2000-12-01 and 2013-11-07");
            Console.WriteLine("Enter the start date:(yyyy-mm-dd)");//User enters the start and end dates
            string start = Console.ReadLine();
            Console.WriteLine("Enter The end date:(yyyy-mm-dd)");
            string end = Console.ReadLine();
            string startdate = DateTime.ParseExact(start,
                                   "yyyy-MM-dd",
                                    CultureInfo.InvariantCulture).ToString();
            string enddate = DateTime.ParseExact(end,
                                   "yyyy-MM-dd",
                                    CultureInfo.InvariantCulture).ToString(); ;
            double o = (DateTime.Parse(enddate) - DateTime.Parse(startdate)).TotalDays;//Difference of start and end dates
            DataTable d = s(path + "\\RAM.xls");//XLS file is taken as a data table
            if (o > 185 || o < 31)
            {
                Console.WriteLine("The days between start and end dates should be betwee 31 and 185 days");
                Environment.Exit(0);
            }
            var d1 = d.Rows[0];//Considering the required rows from the data table
            DataRow[] result = d.Select("enddate >= '" + startdate + "' AND startdate<='" + enddate + "' ");
            result = result.Distinct().ToArray();//An array list is taken for start and end dates
            var foos = new HashSet<DataRow>(result);
            int k = 0, j = result.Length - 1;
            HashSet<int> p = new HashSet<int>();


            result = foos.ToArray();
            double[] sum; // declare numbers as an int array of any size
            sum = new double[4];


            double numofdays = 0;//Initialize the number of days is zero
            for (int i1 = 0; i1 < result.Length; i1++)
            {
                var number = result[i1][3];//Here considering the expodays and absored dose values from the XLS file
                var dose = result[i1][5];
                DateTime start1;
                string loc = result[i1][2].ToString();
                DateTime s1 = DateTime.Parse(result[i1][0].ToString());
                DateTime s2 = DateTime.Parse(result[i1][1].ToString());
                float wt1 = float.Parse(dose.ToString()) / float.Parse(number.ToString());//Calculating the wt1=absored days/expodays
                int i = DateTime.Compare(DateTime.Parse(result[i1][1].ToString()), DateTime.Parse(enddate));
                if (i < 0)//Considering the user enter dates with each location from the XLS file and calculating the wt1 for 4 locations
                {
                    j = DateTime.Compare(DateTime.Parse(startdate), s1);
                    if (j == 0)
                    {

                        start1 = s1;
                    }
                    else if (j < 0)
                    {

                        start1 = s1;
                    }
                    else
                    {

                        start1 = DateTime.Parse(startdate);

                    }

                    numofdays = (start1 - s2).TotalDays;
                }
                else if (i > 0)
                {
                    j = DateTime.Compare(DateTime.Parse(startdate), s1);
                    if (j == 0)
                    {

                        start1 = s1;
                    }
                    else if (j < 0)
                    {

                        start1 = s1;
                    }
                    else
                    {

                        start1 = DateTime.Parse(startdate);

                    }

                    numofdays = (start1 - DateTime.Parse(enddate)).TotalDays;


                }
                else
                {
                    j = DateTime.Compare(DateTime.Parse(startdate), s1);
                    if (j == 0)
                    {

                        start1 = s1;
                    }
                    else if (j < 0)
                    {

                        start1 = s1;
                    }
                    else
                    {

                        start1 = DateTime.Parse(startdate);

                    }

                    numofdays = (start1 - s2).TotalDays;


                }
                if (loc == "(SM-1)")
                    sum[0] += (wt1 * (-numofdays));
                if (loc == "(SM-2)")
                    sum[1] += (wt1 * (-numofdays));
                if (loc == "(SM-3)")
                    sum[2] += (wt1 * (-numofdays));
                if (loc == "(SM-4)")
                    sum[3] += (wt1 * (-numofdays));
            }//Finally calculating the required dose value for given dates

            //foreach (double d12 in sum)
               // Console.WriteLine(d12);
            Console.WriteLine("The dose value in SM-1 is: " + sum[0]);//Displaying the dose values for each location
            Console.WriteLine("The dose value in SM-2 is: " + sum[1]);
            Console.WriteLine("The dose value in SM-3 is: " + sum[2]);
            Console.WriteLine("The dose value in SM-4 is: " + sum[3]);

            Console.ReadKey();

        }


        public static DataTable s(string path)
        {

            OleDbCommand cmd = new OleDbCommand();//This is the OleDB data base connection to the XLS file
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();

            String connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            String[] s = getsheet(path);
            for (int i = 0; i < s.Length; i++)
            {
                String query = "SELECT * FROM [Sheet2$]"; // You can use any different queries to get the data from the excel sheet
                OleDbConnection conn = new OleDbConnection(connString);
                if (conn.State == ConnectionState.Closed) conn.Open();
                try
                {
                    cmd = new OleDbCommand(query, conn);
                    da = new OleDbDataAdapter(cmd);
                    da.Fill(ds);
                    DataTable firstTable = ds.Tables[0];

                    return firstTable;
                }
                catch
                {
                    // Exception Msg 

                    return null;
                }
                finally
                {
                    da.Dispose();
                    conn.Close();

                }

            }
            return null;
        }
        private static String[] getsheet(string excelFile)
        {
            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;

            try
            {

                String connString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                  "Data Source=" + excelFile + ";Extended Properties=Excel 8.0;";

                objConn = new OleDbConnection(connString);

                objConn.Open();

                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null;
                }

                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;


                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }


                return excelSheets;
            }
            catch (Exception ex)
            {
                return null;
            }
            finally
            {
                // Clean up.
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }
    }
}

