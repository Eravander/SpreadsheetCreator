using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using SQL = System.Data;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace SpreadSheet_Creator
{
    public partial class Form1 : Form
    {
        
        List<string> list;
        public Form1()
        {
            //These check boxes would ideally be used to select Current and Future seasons
            InitializeComponent();
            list = new List<string>();
            listBox1.Items.AddRange(new string[] { "FA19", "SP20" });
            listBox1.ItemCheck += new ItemCheckEventHandler(ListBox1_ItemCheck);
        }

        private void GenBtn_Click(object sender, EventArgs e)
        {
            //Connect to Database
            string conString = "Data Source=localhost;Initial Catalog=master_show_file;Integrated Security=True";
            StringBuilder query = new StringBuilder();
            query.Append("SELECT [Show Code], [Date 1]");
            query.Append("FROM [Master_Show_File].[dbo].[DATA_F19] ");

            SQL.DataTable showData = new SQL.DataTable();
            using (SqlConnection cn = new SqlConnection(conString))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(query.ToString(), cn))
                {
                    da.Fill(showData);
                }
            }
           
            //configure excel, this will eventually create 2 sheets. One for Alpha and one for Chrono
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;

            oXL = new Excel.Application();
            oXL.Visible = true;

            oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;
           
            //Import data from SQL table to newly created excel spreadsheet
            try
            {
                SQL.DataTable dtShows = showData.DefaultView.ToTable(true, "Show Code");

                foreach (SQL.DataRow show in dtShows.Rows)
                {
                    oSheet = (Excel._Worksheet)oXL.Worksheets.Add();
                    oSheet.Name = show[0].ToString().Replace(" ", "").Replace("  ", "").Replace("/", "").Replace("\\", "").Replace("*", "");

                    string[] colNames = new string[showData.Columns.Count];

                    int col = 0;

                    foreach (SQL.DataColumn dc in showData.Columns)
                        colNames[col++] = dc.ColumnName;

                    char lastColumn = (char)(65 + showData.Columns.Count - 1);

                    oSheet.get_Range("A1", lastColumn + "1").Value2 = colNames;
                    oSheet.get_Range("A1", lastColumn + "1").Font.Bold = true;
                    oSheet.get_Range("A1", lastColumn + "1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                    SQL.DataRow[] dr = showData.Select(string.Format("Show Code='{0}'", show[0].ToString()));

                    string[,] rowData = new string[dr.Count<SQL.DataRow>(), showData.Columns.Count];

                    //Placeholder code for conditional formatting pulled from examples found online
                    int rowCnt = 0;
                    int redRows = 2;
                    foreach (SQL.DataRow row in dr)
                    {
                        for (col = 0; col < showData.Columns.Count; col++)
                        {
                            rowData[rowCnt, col] = row[col].ToString();
                        }
                        //CompareDate method should go here, if the Date is re moved then continue to next row.

                        //This if statement should begin the process of evaluating how far out the showdate is. Then set conditional formatting based on weeks out.
                        if (DateTime.Parse(row["Date 1"].ToString()) < DateTime.Parse(row["Date 1"].ToString()))
                        {
                            Range range = oSheet.get_Range("A" + redRows.ToString(), "J" + redRows.ToString());
                            range.Cells.Interior.Color = System.Drawing.Color.Red;
                        }

                        //Placeholder
                        redRows++;
                        rowCnt++;
                    }
                    oSheet.get_Range("A2", lastColumn + rowCnt.ToString()).Value2 = rowData;
                }

                oXL.Visible = true;
                oXL.UserControl = true;

                oWB.SaveAs("ShowData.xlsx",
                    AccessMode: Excel.XlSaveAsAccessMode.xlShared);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Marshal.ReleaseComObject(oWB);
            }

        }

        //Use this to compare Show Date to current date
        public static void CompareDate(DateTime t1, DateTime t2)
        {
            //The idea here is to figure out how to parse the DateTime then store it in a variable to be compared, then if it is less than the current date the row is omitted. 
            if(t1 < t2)
            {

            }
        }


        private void ListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            string item = listBox1.SelectedItem.ToString();
            if (e.NewValue == CheckState.Checked)
            {
                if (!list.Contains(item))
                    list.Add(item);
            }
            else
            {
                if (list.Contains(item))
                    list.Remove(item);
            }
        }

        private void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        //public void ReadData()
        //{
        //    Excel ex = new Excel(@"Test", 1);
        //    string[,] read = ex.ReadRange(1, 1, 2749, 3);
        //    ex.Close();
        //}
        //public void WriteData()
        //{
        //    Excel excel = new Excel(@"Test.xlsx", 1);
        //    excel.WriteToCell(0, 0, "Test2");
        //    excel.Save();
        //    excel.SaveAs(@"Test.xlsx");

        //    excel.Close();
        //}
    }
}
