using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace InstaJsonToExcelConverter
{
        public partial class Form1 : Form
        {
     
            DataTable dtAll = new DataTable();
            public Form1()
            {
                InitializeComponent();
            }


            private void button_Click(object sender, EventArgs e)
            {
                OpenFileDialog file = new OpenFileDialog();

                if (file.ShowDialog() == DialogResult.OK)
                {
                    string filePath = file.FileName;
                    ProcessJsonFiles(filePath);
                }
            }

            private void ProcessJsonFiles(string filePath)
            {
                // Initialize the DataTable for storing all data
                dtAll.Columns.Add(new DataColumn("href", typeof(string)));
                dtAll.Columns.Add(new DataColumn("type", typeof(string)));
                dtAll.Columns.Add(new DataColumn("time", typeof(string)));

                // Get all JSON files in the specified directory
                DirectoryInfo directoryInfo = new DirectoryInfo(Path.GetDirectoryName(filePath));
                FileInfo[] files = directoryInfo.GetFiles("*.json");

                // Close any existing Excel processes
                CloseExcelProcesses();

                // Define the output file path
                string outputFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), directoryInfo.Name + ".xlsx");

                // Create a new Excel application and workbook
                var excelApp = new Excel.Application();
                var workbook = excelApp.Workbooks.Add();

                // Process each JSON file and add data to the workbook
                foreach (FileInfo fileInfo in files)
                {
                    string jsonContent = File.ReadAllText(fileInfo.FullName);
                    DataTable dataTable = ParseJson(jsonContent);
                    AddExcelSheet(dataTable, workbook, fileInfo.Name);
                }

                // Remove unwanted rows based on type
                RemoveUnwantedRows();

                // Add a final sheet containing all data
                AddExcelSheet(dtAll, workbook, "all");

                // Save and open the workbook
                SaveAndOpenWorkbook(outputFile, workbook);
            }

            private DataTable ParseJson(string jsonContent)
            {
                jsonContent = jsonContent.Replace("\"relationships_followers\":", "")
                      .Replace("\"relationships_following\":", "")
                      .Replace("\"relationships_following_hashtags\":", "")
                      .Replace("\"relationships_follow_requests_sent\":", "")
                      .Replace("\"relationships_permanent_follow_requests\":", "")
                      .Replace("\"relationships_unfollowed_users\":", "")
                      .Replace("\"relationships_dismissed_suggested_users\":", "")
                      .Replace("\"media_list_data\": [", "")
                      .Replace("\n", "")
                      .Replace("\"media_list_data\": [              ],", "")
                      .Replace("\"title\": \"\",", "")
                      .Replace("\"string_list_data\": [        {          ", "")
                      .Replace("    \"string_list_data\": [      {       ", "")
                      .Replace("[  {                  ", "")
                      .Replace("},  {", "},    {")
                      .Replace("{   [    {                  ", "")
                      .Replace("],", "")
                      .Replace("]", "");


                DataTable dt2 = new DataTable();
                string[] jsonStringArray = Regex.Split(jsonContent, "},    {");
                List<string> ColumnsName = new List<string>();
                foreach (string jSA in jsonStringArray)
                {
                    string[] jsonStringData = Regex.Split(jSA, ",");
                    foreach (string ColumnsNameData in jsonStringData)
                    {

                        try
                        {
                            int idx = ColumnsNameData.IndexOf(":");
                            string ColumnsNameString = ColumnsNameData.Substring(0, idx - 1).Replace("\"", "").Trim();
                            ColumnsName.Add(ColumnsNameString);

                        }
                        catch
                        {

                        }

                    }
                    break;
                }

                foreach (string AddColumnName in ColumnsName)
                {
                    dt2.Columns.Add(AddColumnName.Trim());

                }
                foreach (string jSA in jsonStringArray)
                {
                    if (jSA != " ")
                    {
                        string[] RowData = Regex.Split(jSA.Replace("{", "").Replace("}", ""), ",");
                        DataRow nr = dt2.NewRow();

                        foreach (string rowData in RowData)
                        {
                            try
                            {
                                int idx = rowData.IndexOf(":");
                                string RowColumns = rowData.Substring(0, idx - 1).Replace("\"", "").Trim();
                                string RowDataString = rowData.Substring(idx + 1).Replace("\\n", "").Replace("\"", "").Replace("\\", "").Trim();
                                if (RowColumns == "timestamp")
                                {
                                    DateTime dateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
                                    RowDataString = dateTime.AddSeconds(Convert.ToDouble(RowDataString)).ToLocalTime().ToString("MM/dd/yyyy HH:mm:ss");
                                }
                                nr[RowColumns] = RowDataString;
                            }
                            catch
                            {

                            }
                        }
                        dt2.Rows.Add(nr);
                    }
                }
                return dt2;
            }
            private void RemoveUnwantedRows()
            {
                RemoveRowsByType("followers_1");
                RemoveRowsByType("removed_suggestions");
                RemoveRowsByType("following_hashtags");
            }

            private void RemoveRowsByType(string type)
            {
                DataRow[] rowsToRemove = dtAll.Select($"type = '{type}'");

                foreach (DataRow rowToRemove in rowsToRemove)
                {
                    dtAll.Rows.Remove(rowToRemove);
                }

                dtAll.AcceptChanges();
            }

            private void CloseExcelProcesses()
            {
                Process[] processes = Process.GetProcessesByName("Excel");

                foreach (Process process in processes)
                {
                    if (process.MainWindowTitle.Length == 0)
                    {
                        process.Kill();
                    }
                }
            }

            private void SaveAndOpenWorkbook(string outputFile, Excel.Workbook workbook)
            {
                workbook.SaveAs(outputFile, 51, Type.Missing, Type.Missing, false, false,
                    Excel.XlSaveAsAccessMode.xlExclusive, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);

                workbook.Save();
                workbook.Close();

                Process.Start(outputFile);
            }
            private void AddExcelSheet(DataTable dt, Excel.Workbook wb, string name)
            {
                if (dt.Rows.Count > 0)
                {
                    Excel.Sheets sh = wb.Sheets;
                    Excel.Worksheet osheet = sh.Add();
                    osheet.Name = name.Replace(".json", "");
                    int colIndex = 0;
                    int rowIndex = 1;

                    foreach (DataColumn dc in dt.Columns)
                    {
                        colIndex++;
                        osheet.Cells[1, colIndex] = dc.ColumnName;
                    }
                    foreach (DataRow dr in dt.Rows)
                    {
                        rowIndex++;
                        colIndex = 0;
                        DataRow nr = dtAll.NewRow();
                        foreach (DataColumn dc in dt.Columns)
                        {
                            colIndex++;
                            osheet.Cells[rowIndex, colIndex] = dr[dc.ColumnName];

                            if (name == "all")
                            {

                            }
                            else if (dc.ColumnName == "href")
                            {

                                nr["href"] = dr[dc.ColumnName];
                                nr["type"] = name.Replace(".json", "");

                            }
                            else if (dc.ColumnName == "timestamp")
                            {
                                nr["time"] = dr[dc.ColumnName];
                                dtAll.Rows.Add(nr);
                            }

                        }
                    }
                    osheet.Columns.AutoFit();
                    osheet.Rows.AutoFit();
                }
            }
        }
    }


