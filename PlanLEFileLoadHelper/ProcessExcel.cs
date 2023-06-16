using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using log4net;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace PlanLEFileLoadHelper
{
    public class ProcessExcel
    {
        DataTable tblEcomStg = null;
        SendSMTPMail mail = new SendSMTPMail();
        
        const string sFilePattern = "Ecom Daily Plan*.xlsx";
        const string DOT = ".";
        const string SCHEMA = "AIMDB";

        private ILog log ;

        string[] arrHeader = { "date", "demand plan", "net sales plan", "item margin $ plan", "item margin % plan", "orders shipped", "units shipped", "version" };
        string[] arrSheetName = { "dkny", "klp", "wl", "bass", "am" };

        public string HeaderString
        {
            get
            { 
                return string.Join(",", arrHeader); 
            }
        }
       
        public void CreateDataTable()
        {

            string sDataTableName= "MyTable";
            tblEcomStg = new DataTable(sDataTableName);         

            tblEcomStg.Columns.Add(new DataColumn("DATE", typeof(DateTime)));
            tblEcomStg.Columns.Add(new DataColumn("DEMAND_PLAN", typeof(int)));
            tblEcomStg.Columns.Add(new DataColumn("SLS_RTL", typeof(int)));
            tblEcomStg.Columns.Add(new DataColumn("ITEM_MRGN", typeof(int)));
            tblEcomStg.Columns.Add(new DataColumn("ITEM_MRGN_PCT_TY", typeof(decimal)));
            tblEcomStg.Columns.Add(new DataColumn("SHIPPED_ORDERS", typeof(int)));
            tblEcomStg.Columns.Add(new DataColumn("SHIPPED_UNIT_VOLUME", typeof(int)));
            tblEcomStg.Columns.Add(new DataColumn("PLN_VRSN", typeof(String)));
            tblEcomStg.Columns.Add(new DataColumn("LOCATION", typeof(string)));
            tblEcomStg.Columns.Add(new DataColumn("LOAD_DATE", typeof(DateTime)));
        }

        public bool Validate(string Excelfilename, int minRowCount)
        {
           
                IWorkbook workbook = null;
                ISheet worksheet = null;
            using (FileStream FS = new FileStream(Excelfilename, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(FS);
                    {
                        //Validation include 2 parts file check and can data be loaded to data table?
                        PerformValidations(workbook);
                        AddDataToDatatable(workbook);

                        if (tblEcomStg.Rows.Count < minRowCount)
                        {
                            throw new Exception("Number of rows less than minRowCount");
                        }
                    }
                }
                return true;           
        }

        public bool LoadData(string sFileFullPath, string sArchivePath, int minRowCount)
        {
            try
            {

                string sLoaddate = DateTime.Now.ToString("MM/dd/yyyy");

                Logger.LogInsert("Truncating Stage Table");

                string sDelete = "TRUNCATE TABLE " + SCHEMA + DOT + "STG_ECOM_PLAN";
                DataHelper.ExecNonQueryOLEDB(sDelete);

                Logger.LogInsert("Started loading file");

                if (File.Exists(sFileFullPath))
                {
                    Logger.LogInsert("File Name :" + sFileFullPath);

                    Excel_To_StageTable(sFileFullPath);

                    Logger.LogInsert(tblEcomStg.Rows.Count + "Rows found in file");

                    //Check if rowcount is in desirable range
                    if (tblEcomStg.Rows.Count > minRowCount)
                    {
                        CopyStageTableToMainTable(sLoaddate);
                        //ArchiveFile(sFileFullPath);
                    }

                    Logger.LogInsert("File load completed");

                    //Archive File
                    Logger.LogInsert("End File Name :" + sFileFullPath);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Logger.LogError("main", ex);
                //mail.Dosend(sAdminEmail, "Plan LE job failed", ex.Message);
                return false;
            }
            finally
            {
                //for sent mail after all execution done
                // if (failFilenames != "" && ConfigurationManager.AppSettings["IsMailSend"].ToString().Equals("true"))
                //    mail.Dosend(ConfigurationManager.AppSettings["To"].ToString(), "Below files are fails-" + failFilenames, "Plan and LE Fails File Status");
            }
            return true;
        }

               

        private void CopyStageTableToMainTable(string sLoadDate)
        {
            string sDelCommand = "Delete from AIMDB.ECOM_PLAN where PLN_VRSN in (select PLN_VRSN FROM STG_ECOM_PLAN where rownum <2)";
            DataHelper.ExecNonQueryOLEDB(sDelCommand);

            string sCommand = "Insert into AIMDB.ecom_plan(PLN_VRSN, DAY, SLS_RTL, ITEM_MRGN, ITEM_MRGN_PCT_TY, DEMAND_PLAN" +
                    ", SHIPPED_ORDERS, SHIPPED_UNIT_VOLUME, LOCATION) " +
                    " select PLN_VRSN, DAY,SLS_RTL,ITEM_MRGN,ITEM_MRGN_PCT_TY,DEMAND_PLAN,SHIPPED_ORDERS" +
                    ",SHIPPED_UNIT_VOLUME,LOCATION " +
                    " FROM " +  "STG_ECOM_PLAN stg where stg.LOAD_DATE = to_date('" + sLoadDate + "', 'MM/DD/YYYY')";

            DataHelper.ExecNonQueryOLEDB(sCommand);


        }

        

        private void Excel_To_StageTable(string Excelfilename)
        {
            if (File.Exists(Excelfilename))
            {
                IWorkbook workbook = null;
                ISheet worksheet = null;
                string first_sheet_name = "";


                using (FileStream FS = new FileStream(Excelfilename, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(FS);
                    {                        

                        AddDataToDatatable(workbook);

                        SaveDataTotable();
                    }
                }
            }
        }

        private void PerformValidations(IWorkbook workbook)
        {
            if (!AreValidSheetNames(workbook))
            {
                throw new Exception("Workbook worksheet should be dkny,klp,wl,bass,am");
            }
            if (!AreHeadersValid(workbook, arrHeader.Length))
            {
                throw new Exception("There shoule be " + arrHeader.Length + " Columns in each spreadsheet, names should be " + string.Join(",", arrHeader));
            }

            if (!AreFieldsCorrect(workbook))
            {
                throw new Exception("Some values are not correct");
            }
        }

        private void SaveDataTotable()
        {
            //string sConn = ConfigurationManager.AppSettings["connectionstring"].ToString();

            


            var destTableName = "STG_ECOM_PLAN";

            foreach (DataRow row in tblEcomStg.Rows)
            {
                if (row["DATE"].ToString() != "" && row["DATE"] != null)
                {
                    StringBuilder sqlBatch = new StringBuilder();

                    DateTime Datedt = DateTime.Parse(row["DATE"].ToString());

                    string sDAY = " TO_DATE('" + Datedt.ToString("MM/dd/yyyy") + "','MM/dd/yyyy') ";


                    string sDEMAND_PLAN = row["DEMAND_PLAN"].ToString();
                    string sSLS_RTL = row["SLS_RTL"].ToString();

                    string sITEM_MRGN = "";
                    if (row["ITEM_MRGN"] == null)
                    {
                        sITEM_MRGN = "null";
                    }
                    else
                    {
                        sITEM_MRGN = "'" + row["ITEM_MRGN"].ToString() + "'";
                    }

                    string sITEM_MRGN_PCT_TY = "";
                    if (row["ITEM_MRGN_PCT_TY"] == null)
                    {
                        sITEM_MRGN_PCT_TY = null;
                    }
                    else if (row["ITEM_MRGN_PCT_TY"].ToString() == "")
                    {
                        sITEM_MRGN_PCT_TY = "null";
                    }
                    else
                    {
                        sITEM_MRGN_PCT_TY = "'" + row["ITEM_MRGN_PCT_TY"].ToString() + "'";
                    }

                    string sSHIPPED_ORDERS = "";
                    if (row["SHIPPED_ORDERS"] == null)
                    {
                        sSHIPPED_ORDERS = "null";
                    }
                    else
                    {
                        sSHIPPED_ORDERS = "'" + row["SHIPPED_ORDERS"].ToString() + "'";
                    }

                    string sSHIPPED_UNIT_VOLUME = "";
                    if (row["SHIPPED_UNIT_VOLUME"] == null)
                    {
                        sSHIPPED_UNIT_VOLUME = "null";
                    }
                    else
                    {
                        sSHIPPED_UNIT_VOLUME = "'" + row["SHIPPED_UNIT_VOLUME"].ToString() + "'";
                    }

                    string sLOCATION = row["LOCATION"].ToString();
                    string sPLN_VRSN = row["PLN_VRSN"].ToString();

                    DateTime loaddt = DateTime.Parse(row["LOAD_DATE"].ToString());


                    string sLOAD_DATE = " TO_DATE('" + loaddt.ToString("MM/dd/yyyy") + "','MM/dd/yyyy') ";
                    string sqlStatement = "INSERT INTO " + destTableName +
                    " (DAY,DEMAND_PLAN,SLS_RTL,ITEM_MRGN,ITEM_MRGN_PCT_TY,SHIPPED_ORDERS,SHIPPED_UNIT_VOLUME,LOCATION,PLN_VRSN,LOAD_DATE) " +
                        "VALUES (" + sDAY + ",'" + sDEMAND_PLAN + "','" + sSLS_RTL + "'," + sITEM_MRGN + "," + sITEM_MRGN_PCT_TY + "," + sSHIPPED_ORDERS + "," + sSHIPPED_UNIT_VOLUME
                        + ",'" + sLOCATION + "','" + sPLN_VRSN + "'," + sLOAD_DATE + ")";

                    DataHelper.ExecNonQueryOLEDB(sqlStatement);
                }
            }

        }

        private bool AddDataToDatatable(IWorkbook workbook)
        {
            CreateDataTable();
            ISheet worksheet = null;
            for (int fileIndex = 0; fileIndex < workbook.NumberOfSheets; fileIndex++)
            {
                if (!workbook.IsSheetHidden(fileIndex))
                {
                    worksheet = workbook.GetSheetAt(fileIndex);

                    DataRow NewReg = null;
                    int rowindex = 1;
                    IRow row = worksheet.GetRow(rowindex);//skip header

                    while (row != null)
                    {
                        DataRow rw = tblEcomStg.NewRow();
                        string sDate = "";
                        var d = row.GetCell(0);
                        if (d.CellType == CellType.Numeric)
                        {
                            //sDate = d.DateCellValue.ToString("MM/dd/yyyy");

                            rw["DATE"] = d.DateCellValue;
                        }



                        var demand = row.GetCell(1);
                        if (demand != null)
                        {
                            rw["DEMAND_PLAN"] = demand.NumericCellValue;
                        }
                        else
                        {
                            rw["DEMAND_PLAN"] = null;
                        }


                        var sls = row.GetCell(2);
                        if (sls != null)
                        {

                            rw["SLS_RTL"] = sls.NumericCellValue;
                        }
                        else
                        {
                            rw["SLS_RTL"] = null;
                        }

                        var itmmrgn = row.GetCell(3);
                        if (itmmrgn != null)
                        {

                            rw["ITEM_MRGN"] = itmmrgn.NumericCellValue;
                        }
                        //else
                        //{
                        //    rw["ITEM_MRGN"] = null;
                        //}

                        var itmmrgnpct = row.GetCell(4);
                        if (itmmrgnpct != null)
                        {

                            rw["ITEM_MRGN_PCT_TY"] = itmmrgnpct.NumericCellValue;
                        }
                        //else
                        //{
                        //    rw["ITEM_MRGN_PCT_TY"] = null;
                        //}


                        var ordshp = row.GetCell(5);
                        if (ordshp != null)
                        {
                            rw["SHIPPED_ORDERS"] = ordshp.NumericCellValue;
                        }
                        //else
                        //{
                        //    rw["SHIPPED_ORDERS"] = null;
                        //}

                        var shpvol = row.GetCell(6);
                        if (shpvol != null)
                        {

                            rw["SHIPPED_UNIT_VOLUME"] = shpvol.NumericCellValue;
                        }
                        //else
                        //{
                        //    rw["SHIPPED_UNIT_VOLUME"] = null;
                        //}

                        var plnver = row.GetCell(7);
                        if (plnver != null)
                        {
                            rw["PLN_VRSN"] = plnver.StringCellValue + "_ECDMD";
                        }
                        else
                        {
                            throw new Exception("unknown version");
                        }

                        if (worksheet.SheetName.ToLower().Trim() == "dkny")
                            rw["LOCATION"] = "DKNY Ecommerce";
                        else if (worksheet.SheetName.ToLower().Trim() == "klp")
                            rw["LOCATION"] = "Karl Lagerfeld Paris Ecommerce";
                        else if (worksheet.SheetName.ToLower().Trim() == "wl")
                            rw["LOCATION"] = "Wilsons Leather Ecommerce";
                        else if (worksheet.SheetName.ToLower().Trim() == "bass")
                            rw["LOCATION"] = "GH Bass Ecommerce";
                        else if (worksheet.SheetName.ToLower().Trim() == "am")
                            rw["LOCATION"] = "Andrew Marc Ecommerce";
                        else
                            throw new Exception("unknown tab");

                        rw["LOAD_DATE"] = DateTime.Now.ToString("MM/dd/yyyy");




                        tblEcomStg.Rows.Add(rw);
                        row = worksheet.GetRow(++rowindex);
                    }




                }
            }

            return true;
        }

        private bool AreFieldsCorrect(IWorkbook workbook)
        {
            ISheet worksheet = null;
            for (int fileIndex = 0; fileIndex < workbook.NumberOfSheets; fileIndex++)
            {
                if (!workbook.IsSheetHidden(fileIndex))
                {
                    worksheet = workbook.GetSheetAt(fileIndex);

                    DataRow NewReg = null;
                    int rowindex = 1;
                    IRow row = worksheet.GetRow(rowindex);//skip header

                    while (row != null)
                    {

                        string sDate = "";
                        var d = row.GetCell(0);
                        if (d.CellType == CellType.Numeric)
                        {
                            if (d.DateCellValue.ToString() != "")
                                sDate = d.DateCellValue.ToString("MM/dd/yyyy");
                        }



                        if (!IsDateValue(sDate))
                            return false;


                        for (int j = 1; j < arrHeader.Length - 2; j++)
                        {
                            var v = row.GetCell(j);
                            if (v != null)//nulls are allowe for cells, weekends will not have sale or shipped value
                            {
                                if (v.CellType == CellType.Formula || v.CellType == CellType.Numeric)
                                    if (!IsNumericValue(v.NumericCellValue.ToString()))
                                        return false;
                            }
                        }


                        if (!IsDateValue(sDate))
                            return false;


                        var c = row.GetCell(arrHeader.Length - 1);
                        string sVersion = c.StringCellValue;

                        if (!IsValidVersion(sVersion))
                            return false;

                        row = worksheet.GetRow(++rowindex);
                    }




                }
            }

            return true;
        }

        private bool IsNumericValue(string v)
        {
            if (v.Trim() == "")
                return false;

            double test;
            return double.TryParse(v, out test);




            return false;
        }

        private bool IsDateValue(string sDate)
        {
            DateTime dt;
            if (sDate.Trim() == "")
                return false;

            if (!DateTime.TryParse(sDate, out dt))
            {
                return false;
            }

            return true;
        }

        private bool IsValidVersion(string sVersion)
        {
            string tVer = sVersion.ToUpper().Trim();
            if (!(tVer == "LE01" || tVer == "LE02" || tVer == "LE03" || tVer == "LE04" || tVer == "LE05" || tVer == "LE06" || tVer == "LE07" || tVer == "LE08"
                || tVer == "LE09" || tVer == "LE10" || tVer == "LE11" || tVer == "LE12" || tVer == "PLAN"))
            {
                return false;
            }
            return true;
        }

        private bool EmptyRow(IRow row)
        {

            for (int i = 0; i < arrHeader.Length; i++)
            {
                if (row.GetCell(i).StringCellValue.Trim().Length == 0)
                    return false;
            }
            return true;
        }

        private bool AreValidSheetNames(IWorkbook workbook)
        {
            ISheet worksheet = null;
            for (int fileIndex = 0; fileIndex < workbook.NumberOfSheets; fileIndex++)
            {
                if (!workbook.IsSheetHidden(fileIndex))
                {
                    if (!IsValidSheetName(fileIndex, workbook.GetSheetAt(fileIndex).SheetName))
                        return false;
                }
            }

            return true;
        }

        private bool IsValidSheetName(int fileIndex, string sheetName)
        {

            foreach (string sname in arrSheetName)
            {
                if (sname.ToLower().Trim() == sheetName.ToLower().Trim())
                {
                    return true;
                }
            }


            return false;
        }

        private bool AreHeadersValid(IWorkbook workbook, int maxColCount)
        {
            ISheet worksheet = null;
            for (int fileIndex = 0; fileIndex < workbook.NumberOfSheets; fileIndex++)
            {
                if (!workbook.IsSheetHidden(fileIndex))
                {
                    worksheet = workbook.GetSheetAt(fileIndex);

                    DataRow NewReg = null;
                    IRow row = worksheet.GetRow(0);
                    if (row != null) //null is when the row only contains empty cells 
                    {

                        for (int colNum = 0; colNum < maxColCount; colNum++)
                        {
                            var c = row.GetCell(colNum);
                            string sheader = c.StringCellValue;

                            if (!CheckHeaderColumn(colNum, sheader))
                                return false;
                        }
                    }
                    else
                    {
                        throw new Exception("Header not found in worksheet" + worksheet.SheetName);
                        // return null;
                    }

                }
            }

            return true;
        }

        private bool CheckHeaderColumn(int i, string header)
        {
            if (arrHeader[i] == header.ToLower().Trim())
                return true;

            return false;
        }

        
    }
}
