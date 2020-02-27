using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ConsoleAPplication
{
    public partial class Form1 : Form
    {

        Microsoft.Office.Interop.Excel.Application excel;
        Microsoft.Office.Interop.Excel.Workbook worKbooK;
        Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
        Microsoft.Office.Interop.Excel.Range celLrangE;


        public Form1()
        {
            InitializeComponent();
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            string resultFilePath = AppDomain.CurrentDomain.BaseDirectory + "RESULT";
            var connectionString = @"Server=(Local)\SQLExpress;Database=Vinod;Integrated Security=true";
            List<string> lstRollNumberRecords = new List<string>();




            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            string query = "select Roll from Sheet1$ group by Roll order by Roll";
            SqlCommand cmd = new SqlCommand(query, connection);
            SqlDataAdapter sqlDataReader = new SqlDataAdapter(cmd);
            DataTable tblGetData = new DataTable();
            sqlDataReader.Fill(tblGetData);
            connection.Close();

            foreach (DataRow row in tblGetData.Rows)
            {
                lstRollNumberRecords.Add(row["Roll"].ToString());
            }

            //Create dictionay
            Dictionary<string, List<string>> dicRowData = new Dictionary<string, List<string>>();
            foreach (var item in lstRollNumberRecords)
            {
                string qq = "select distinct CLASS from Sheet1$ where ROLL = '" + item.ToString() + "'";
                SqlConnection con = new SqlConnection(connectionString);
                con.Open();
                SqlCommand cmdd = new SqlCommand(qq, con);
                SqlDataAdapter sqlDataReaderr = new SqlDataAdapter(cmdd);
                DataTable tblGetDataaa = new DataTable();
                sqlDataReaderr.Fill(tblGetDataaa);
                connection.Close();

                List<string> classRecod = new List<string>();

                foreach (DataRow itemRow in tblGetDataaa.Rows)
                {
                    classRecod.Add(itemRow["CLASS"].ToString());
                }
                dicRowData.Add(item.ToString(), classRecod);
            }



            DataSet ds = new DataSet();

            foreach (var item in dicRowData)
            {
                string folderName = item.Key.ToString();
                List<string> folderValue = item.Value.ToList();

                string NewFolder = resultFilePath + "\\" + folderName;
                //string fileNAme = "";

                if (!Directory.Exists(NewFolder))
                {
                    Directory.CreateDirectory(NewFolder);
                }

                foreach (var folderValues in folderValue)
                {
                    if (folderName.Equals("DL01"))
                    {
                        continue;
                    }
                    else
                    {
                        if (folderValues.Equals("#N/A"))
                        {
                            continue;
                        }
                        else
                        {
                            using (FileStream fs = File.Create(NewFolder + "\\" + folderValues + ".xlsx"))
                            {
                            }

                            string qry = "select [ROLL NUMBER],Name,SCORE,GRADE from Sheet1$ where ROLL = '" + folderName + "' and CLASS='" + folderValues + "' order by SCORE desc";
                            SqlConnection con = new SqlConnection(connectionString);
                            con.Open();
                            SqlCommand cmmd = new SqlCommand(qry, con);
                            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(cmmd);
                            DataTable tblData = new DataTable();
                            tblData.TableName = item.ToString();
                            sqlDataAdapter.Fill(tblData);

                            con.Close();


                            //Create Excel File 
                            try
                            {
                                excel = new Microsoft.Office.Interop.Excel.Application();
                                excel.Visible = false;
                                excel.DisplayAlerts = false;
                                worKbooK = excel.Workbooks.Add(Type.Missing);


                                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                                worKsheeT.Name = folderValues.ToString();

                                //worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[1, 8]].Merge();
                                //worKsheeT.Cells[1, 1] = "Student Report Card";
                                // worKsheeT.Cells.Font.Size = 12;

                                //Bold the First Row
                                celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[1, 4]];
                                celLrangE.EntireRow.Font.Bold = true;//.AutoFit();


                                int rowcount = 1;

                                foreach (DataRow datarow in tblData.Rows)
                                {
                                    rowcount += 1;
                                    for (int i = 1; i <= tblData.Columns.Count; i++)
                                    {

                                        if (rowcount == 2)
                                        {
                                            worKsheeT.Cells[1, i] = tblData.Columns[i - 1].ColumnName;
                                            // worKsheeT.Cells.Font.Color = System.Drawing.Color.Red;
                                            //worKsheeT.get_Range(1, i).Font.Bold = true;

                                        }

                                        worKsheeT.Cells[rowcount, i] = datarow[i - 1].ToString();

                                        if (rowcount > 3)
                                        {
                                            if (i == tblData.Columns.Count)
                                            {
                                                if (rowcount % 2 == 0)
                                                {
                                                    celLrangE = worKsheeT.Range[worKsheeT.Cells[rowcount, 1], worKsheeT.Cells[rowcount, tblData.Columns.Count]];
                                                }

                                            }
                                        }

                                    }

                                }

                                celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[rowcount, tblData.Columns.Count]];
                                celLrangE.EntireColumn.AutoFit();

                                celLrangE.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                                Microsoft.Office.Interop.Excel.Borders border = celLrangE.Borders;
                                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                border.Weight = 2d;

                                worKbooK.SaveAs(NewFolder + "\\" + folderValues + ".xlsx"); ;
                                worKbooK.Close();
                                excel.Quit();

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);

                            }
                            finally
                            {
                                worKsheeT = null;
                                celLrangE = null;
                                worKbooK = null;
                            }

                            //End Excel File



                            // tblData.WriteXml(NewFolder + "\\" + folderValues + ".xls");
                        }

                    }
                }
            }


        }
    }
}
