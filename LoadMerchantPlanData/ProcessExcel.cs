using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace LoadMerchantPlanData
{
    public class ProcessExcel
    {

        SendSMTPMail mail = new SendSMTPMail();

        public void process()
        {
            try
            {
                //if (ConfigurationManager.AppSettings["IsMailSend"].ToString().Equals("true"))
                //    mail.Dosend(ConfigurationManager.AppSettings["To"].ToString(), "Start Plan and LE Program ", "Plan and LE Program Status");

                var finalproccall = false;
                string folderpath = ConfigurationManager.AppSettings["LocalFolderDirectoryPath"].ToString();
                DirectoryInfo d = new DirectoryInfo(folderpath);

                OleDbConnection con2 = new OleDbConnection(ConfigurationManager.AppSettings["connectionstring"].ToString());
                var loaddate = DateTime.Now.ToString("MM/dd/yyyy");
                //for truncate data using loaddate
                if (d.GetFiles("*.xlsx")?.Length > 0)
                {
                    finalproccall = true;
                    Logger.LogInsert("Total Numer of file for loading-" + d.GetFiles("*.xlsx")?.Length);
                }
                // Logger.LogInsert("Started Files Loop");
                foreach (FileInfo file in d.GetFiles("*.xlsx"))
                {
                    try
                    {
                        Logger.LogInsert("Started deleting existing LE data-" + file.FullName);
                        con2.Open();
                        OleDbCommand com = new OleDbCommand("usp_deleteSTGEcomPlanLE", con2);
                        com.CommandType = CommandType.StoredProcedure;
                        OleDbParameter sqlParam = com.Parameters.Add("loaddate", OleDbType.VarChar);
                        sqlParam.Value = loaddate;
                        com.ExecuteNonQuery();
                        con2.Close();
                        Logger.LogInsert("End deleting existing LE data-" + file.FullName);
                        break;


                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        Logger.LogError("main", ex);
                    }
                }
                if (d.GetFiles("*.xlsx")?.Length > 0)
                {
                    foreach (FileInfo file in d.GetFiles("*.xlsx"))
                    {
                        try
                        {
                            Logger.LogInsert("Started File Name :" + file.Name);
                            // Load the Excel file
                            System.Data.DataTable dt = new System.Data.DataTable();

                            dt = ImportExceltoDatatable(file.FullName, "Plan & LE", true);

                            //going to insert this data into the oracale DB

                            SaveUsingOracleBulkCopy("STG_ECOM_Plan", dt, con2, loaddate);


                            string moveTo = ConfigurationManager.AppSettings["ArchiveLocalFolderDirectoryPath"].ToString() + AppendTimeStamp(file.Name);
                            //moving file
                            File.Move(file.FullName, moveTo);
                            Logger.LogInsert("End File Name :" + file.Name);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            Logger.LogError("main", ex);
                        }
                    }
                    Logger.LogInsert("End Files Loop");
                    if (finalproccall)
                    {
                        //calling final proc for moving data stg table to main table
                        try
                        {
                            Logger.LogInsert("Started calling final proc for moving data stg table to main table :usp_LoadEcomPlan");
                            con2.Open();
                            OleDbCommand com = new OleDbCommand("usp_LoadEcomPlan", con2);
                            com.CommandType = CommandType.StoredProcedure;
                            OleDbParameter sqlParam = com.Parameters.Add("loaddate", OleDbType.VarChar);
                            sqlParam.Value = loaddate;

                            com.ExecuteNonQuery();
                            con2.Close();
                            Logger.LogInsert("End calling final proc for moving data stg table to main table :usp_LoadEcomPlan");

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            Logger.LogError("main", ex);
                        }
                    }
                    Logger.LogInsert("End Plan and LE Program ");
                    //if (ConfigurationManager.AppSettings["IsMailSend"].ToString().Equals("true"))
                    //    mail.Dosend(ConfigurationManager.AppSettings["To"].ToString(), "End Plan and LE Program ", "Plan and LE Program Status");
                }
                else
                {
                    Logger.LogInsert("No File In Folder");
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Logger.LogError("main", ex);
            }
            finally
            {
                //for sent mail after all execution done
                // if (failFilenames != "" && ConfigurationManager.AppSettings["IsMailSend"].ToString().Equals("true"))
                //    mail.Dosend(ConfigurationManager.AppSettings["To"].ToString(), "Below files are fails-" + failFilenames, "Plan and LE Fails File Status");
            }
        }

        public static string AppendTimeStamp(string fileName)
        {
            return string.Concat(
                Path.GetFileNameWithoutExtension(fileName),
                DateTime.Now.ToString("yyyyMMddHHmmssfff"),
                Path.GetExtension(fileName)
                );
        }

        public void SaveUsingOracleBulkCopy(string destTableName, DataTable dt, OleDbConnection con2, string loaddate)
        {
            try
            {


                string sqlStatement = "INSERT INTO " + destTableName + " (DAY,DEMAND_PLAN,SLS_RTL,ITEM_MRGN,ITEM_MRGN_PCT_TY,SHIPPED_ORDERS,SHIPPED_UNIT_VOLUME,LOCATION,PLN_VRSN,LOAD_DATE,NET_AUR_PLAN) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}')";

                foreach (DataRow row in dt.Rows)
                {
                    if (row[0] != "" && row[0] != null)
                    {
                        StringBuilder sqlBatch = new StringBuilder();
                        sqlBatch.AppendLine(string.Format(sqlStatement, Convert.ToDateTime(row[0]).ToString("MM/dd/yyyy"), row[1]?.ToString()?.Replace("$", "")?.Replace(",", "")?.Replace("-", "")?.Replace("(", "")?.Replace(")", "")?.Trim(), row[2]?.ToString()?.Replace("$", "")?.Replace(",", "")?.Replace("-", "")?.Replace("(", "")?.Replace(")", "")?.Trim(), row[3]?.ToString()?.Replace("$", "")?.Replace(",", "")?.Replace("-", "")?.Replace("(", "")?.Replace(")", "")?.Trim(), row[4]?.ToString()?.Replace("$", "")?.Replace(",", "")?.Replace("%", "")?.Replace("-", "")?.Replace("(", "")?.Replace(")", "")?.Trim(), row[5]?.ToString()?.Replace("$", "")?.Replace(",", "")?.Replace("-", "")?.Replace("(", "")?.Replace(")", "")?.Trim(), row[6]?.ToString()?.Replace("$", "")?.Replace(",", "")?.Replace("-", "")?.Replace("(", "")?.Replace(")", "")?.Trim(), row[7], row[8], loaddate, row[9]?.ToString()?.Replace("$", "")?.Replace(",", "")?.Replace("-", "")?.Replace("(", "")?.Replace(")", "")?.Trim()));
                        con2.Open();
                        OleDbCommand cmd1 = new OleDbCommand(sqlBatch.ToString(), con2);
                        cmd1.ExecuteNonQuery();
                        con2.Close();
                    }
                }

            }
            catch (Exception ex)
            {
                Logger.LogError("main", ex);
                throw ex;
            }
        }



        public System.Data.DataTable ImportExceltoDatatable(string filePath, string SheetName, bool hasHeaderFlag)
        {
            try
            {
                DataSet ds = new DataSet();

                string constring = "";
                string hdrOption = "Yes";

                // hasHeaderFlag = true;
                if (!hasHeaderFlag)
                {
                    hdrOption = "No";
                }
                string schemaFileName = Path.GetDirectoryName(filePath) + @"\Schema.ini";
                if (File.Exists(schemaFileName))
                {
                    File.Delete(schemaFileName);
                }
                string fileExtension = Path.GetExtension(filePath);
                switch (fileExtension.ToLower())
                {
                    case ".xls": //Excel 1997-2003  
                                 // constring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath
                                 //+ ";Extended Properties=\"Excel 8.0;HDR=" + hdrOption + ";IMEX=1\"";
                        return ReadExcel(filePath, SheetName, hdrOption);
                        break;
                    case ".xlsx": //Excel 2007-2010  
                                  // constring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath
                                  //+ ";Extended Properties=\"Excel 12.0 xml;HDR=" + hdrOption + ";IMEX=1\"";
                        return ReadExcel(filePath, SheetName, hdrOption);
                        break;
                    default:
                        constring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=" + hdrOption + ";IMEX=1;'";
                        // constring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=" + hdrOption + ";IMEX=1;'";

                        break;
                }

                OleDbDataAdapter da;
                OleDbConnection con = new OleDbConnection(constring + "");
                con.Open();
                System.Data.DataTable dtSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                String sheetName = "";

                if (fileExtension == ".csv" || fileExtension == ".txt")
                {

                    string sqlquery = "Select * from [" + Path.GetFileName(filePath) + "]";
                    da = new OleDbDataAdapter(sqlquery, con);
                    da.Fill(ds);
                }
                else
                {
                    //for (int i = 0; i < dtSchema.Rows.Count; i++)
                    //{
                    //    sheetName = dtSchema.Rows[i]["TABLE_NAME"].ToString();
                    //    break;
                    //}

                    if (SheetName.Contains("$"))
                    {
                        string sqlquery = "Select * from [" + SheetName + "]";
                        da = new OleDbDataAdapter(sqlquery, con);
                        da.Fill(ds);
                    }
                    else
                    {
                        string sqlquery = "Select * from [" + SheetName + "$]";
                        da = new OleDbDataAdapter(sqlquery, con);
                        da.Fill(ds);
                    }
                }

                System.Data.DataTable dtFirstSheetData = ds.Tables[0];


                con.Close();
                return dtFirstSheetData;
            }
            catch (Exception ex)
            {
                Logger.LogError("main", ex);
                return null;
            }
        }


        private System.Data.DataTable ReadExcel(string path, string SheetName, string hdrOption)
        {
            List<string> list = new List<string>();
            list.Add("DAY");
            list.Add("DEMAND_PLAN");
            list.Add("SLS_RTL");
            list.Add("ITEM_MRGN");
            list.Add("ITEM_MRGN_PCT_TY");
            list.Add("SHIPPED_ORDERS");
            list.Add("SHIPPED_UNIT_VOLUME");
            list.Add("LOCATION");
            list.Add("PLN_VRSN");
            list.Add("NET_AUR_PLAN");
            String[] str = list.ToArray();

            //Instance reference for Excel Application
            Microsoft.Office.Interop.Excel.Application objXL = null;
            //Workbook refrence
            Microsoft.Office.Interop.Excel.Workbook objWB = null;
            DataSet ds = new DataSet();
            try
            {
                DataTable dt = new DataTable();
                //Instancing Excel using COM services
                objXL = new Microsoft.Office.Interop.Excel.Application();
                //Adding WorkBook
                objWB = objXL.Workbooks.Open(path);
                if (objWB.Worksheets.Count > 0)
                {

                    foreach (Microsoft.Office.Interop.Excel.Worksheet objSHT in objWB.Worksheets)
                    {
                        if (objSHT.Name == SheetName)
                        {
                            int rows = objSHT.UsedRange.Rows.Count;
                            int cols = objSHT.UsedRange.Columns.Count;

                            int noofrow = 1;
                            //If 1st Row Contains unique Headers for datatable include this part else remove it
                            //Start
                            for (int c = 1; c <= cols; c++)
                            {
                                if (hdrOption == "Yes")
                                {
                                    string colname = objSHT.Cells[1, c].Text;
                                    dt.Columns.Add(str[c - 1]);
                                    noofrow = 2;
                                }
                                else
                                {
                                    dt.Columns.Add("F" + (c));
                                    noofrow = 1;
                                }

                            }
                            //END
                            for (int r = noofrow; r <= rows; r++)
                            {
                                DataRow dr = dt.NewRow();
                                for (int c = 1; c <= cols; c++)
                                {
                                    if (objSHT.Cells[r, c].Text.ToString().Contains("#"))
                                    {
                                        var abc = objSHT.Cells[r, c].Value;
                                        // double val = double.Parse(objSHT.Cells[r, c].Value.ToString());
                                        // DateTime requiredDate = DateTime.FromOADate(val);
                                        dr[c - 1] = abc;
                                    }
                                    else
                                    {
                                        dr[c - 1] = objSHT.Cells[r, c].Text;
                                    }
                                }
                                dt.Rows.Add(dr);
                            }
                            Marshal.ReleaseComObject(objSHT);
                            break;
                        }
                    }
                }
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //Closing workbook
                objWB.Close();
                Marshal.ReleaseComObject(objWB);
                //Closing excel application
                objXL.Quit();
                Marshal.ReleaseComObject(objXL);
                return dt;
            }

            catch (Exception ex)
            {
                objWB.Saved = true;
                //Closing work book
                objWB.Close();
                //Closing excel application
                objXL.Quit();
                //Response.Write("Illegal permission");
                return null;
            }

        }

    }
}
