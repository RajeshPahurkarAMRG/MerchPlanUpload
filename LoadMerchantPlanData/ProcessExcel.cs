using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace LoadMerchantPlanData
{
    public class ProcessExcel
    {
        DataTable Tabla = null;
        SendSMTPMail mail = new SendSMTPMail();

        string[] arrHeader = { "date", "demand plan", "net sales plan", "item margin $ plan", "item margin % plan", "orders shipped", "units shipped", "version" };
        string[] arrSheetName = { "dkny", "klp", "wl", "bass", "am" };
        public void process()
        {
            //try
            //{
                //if (ConfigurationManager.AppSettings["IsMailSend"].ToString().Equals("true"))
                //    mail.Dosend(ConfigurationManager.AppSettings["To"].ToString(), "Start Plan and LE Program ", "Plan and LE Program Status");

                var finalproccall = false;
                string folderpath = ConfigurationManager.AppSettings["LocalFolderDirectoryPath"].ToString();
                DirectoryInfo d = new DirectoryInfo(folderpath);

                OleDbConnection con2 = new OleDbConnection(ConfigurationManager.AppSettings["connectionstring"].ToString());
                var loaddate = DateTime.Now.ToString("MM/dd/yyyy");

                FileInfo[] allFiles = d.GetFiles("Ecom Daily Plan*.xlsx");

                //for truncate data using loaddate
                if (d.GetFiles("*.xlsx")?.Length > 0)
                {
                    finalproccall = true;
                    Logger.LogInsert("Total Numer of file for loading-" + allFiles.Length);
                }
                // Logger.LogInsert("Started Files Loop");
                foreach (FileInfo file in allFiles)
                {
                //try
                //{
                Logger.LogInsert("Started deleting existing LE data-" + file.FullName);
                con2.Open();
                OleDbCommand com = new OleDbCommand("truncate table stg_ecom_plan", con2);             
                com.ExecuteNonQuery();
                con2.Close();
                Logger.LogInsert("End deleting existing LE data-" + file.FullName);
                //    break;


                //}
                //catch (Exception ex)
                //{
                //    Console.WriteLine(ex.Message);
                //    Logger.LogError("main", ex);
                //}
            }
                if (allFiles.Length > 0)
                {
                    foreach (FileInfo file in allFiles)
                    {
                        //try
                        //{
                            Logger.LogInsert("Started File Name :" + file.Name);
                            // Load the Excel file
                            System.Data.DataTable dt = new System.Data.DataTable();

                            Excel_To_DataTable(file.FullName);
                            //dt = ImportExceltoDatatable(file.FullName, "Plan & LE", true,"");

                            //going to insert this data into the oracale DB
                            

                            //string moveTo = ConfigurationManager.AppSettings["ArchiveLocalFolderDirectoryPath"].ToString() + AppendTimeStamp(file.Name);
                            ////moving file
                            //File.Move(file.FullName, moveTo);
                            //Logger.LogInsert("End File Name :" + file.Name);
                        //}
                        //catch (Exception ex)
                        //{
                        //    Console.WriteLine(ex.Message);
                        //    Logger.LogError("main", ex);
                        //}
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
            //}

            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //    Logger.LogError("main", ex);
            //}
            //finally
            //{
            //    //for sent mail after all execution done
            //    // if (failFilenames != "" && ConfigurationManager.AppSettings["IsMailSend"].ToString().Equals("true"))
            //    //    mail.Dosend(ConfigurationManager.AppSettings["To"].ToString(), "Below files are fails-" + failFilenames, "Plan and LE Fails File Status");
            //}
        }

        public static string AppendTimeStamp(string fileName)
        {
            return string.Concat(
                Path.GetFileNameWithoutExtension(fileName),
                DateTime.Now.ToString("yyyyMMddHHmmssfff"),
                Path.GetExtension(fileName)
                );
        }

        
        public void ImportExceltoDatatableNPOI(string path)
        {
            IWorkbook book;
            FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            //book = new (fs);

        }

       
        private void Excel_To_DataTable(string Excelfilename)
        {
            
            
            if (System.IO.File.Exists(Excelfilename))
            {
                IWorkbook workbook = null;
                ISheet worksheet = null;
                string first_sheet_name = "";

                first_sheet_name = "MyTable";
                Tabla = new DataTable(first_sheet_name);
                Tabla.Rows.Clear();
                Tabla.Columns.Clear();

                Tabla.Columns.Add(new DataColumn("DATE", typeof(DateTime)));
                Tabla.Columns.Add(new DataColumn("DEMAND_PLAN", typeof(int)));
                Tabla.Columns.Add(new DataColumn("SLS_RTL", typeof(int)));
                Tabla.Columns.Add(new DataColumn("ITEM_MRGN", typeof(int)));
                Tabla.Columns.Add(new DataColumn("ITEM_MRGN_PCT_TY", typeof(decimal)));
                Tabla.Columns.Add(new DataColumn("SHIPPED_ORDERS", typeof(int)));
                Tabla.Columns.Add(new DataColumn("SHIPPED_UNIT_VOLUME", typeof(int)));
                Tabla.Columns.Add(new DataColumn("PLN_VRSN", typeof(String)));
                Tabla.Columns.Add(new DataColumn("LOCATION", typeof(string)));
                Tabla.Columns.Add(new DataColumn("LOAD_DATE", typeof(DateTime)));

                using (FileStream FS = new FileStream(Excelfilename, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(FS);

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

                        AddDataToDatatable(workbook);

                        SaveDataTotable();
                    }
                }
            }
        }

        private void SaveDataTotable()
        {
            string sConn = ConfigurationManager.AppSettings["connectionstring"].ToString();

            OleDbConnection con2 = new OleDbConnection(sConn);
      


            //foreach (DataRow row in dt.Rows)
            //{

            //    }
            //}
            //con2.Close();
            var destTableName = "STG_ECOM_PLAN";

            foreach (DataRow row in Tabla.Rows)
            {
                if (row["DATE"].ToString() != "" && row["DATE"] != null)
                {
                    StringBuilder sqlBatch = new StringBuilder();

                    DateTime Datedt = DateTime.Parse(row["DATE"].ToString());

                    string sDAY = " TO_DATE('" + Datedt.ToString("MM/dd/yyyy") + "','MM/dd/yyyy') ";

                    
                    string sDEMAND_PLAN = row["DEMAND_PLAN"].ToString();
                    string sSLS_RTL = row["SLS_RTL"].ToString();

                    string sITEM_MRGN = "";
                    if (row["ITEM_MRGN"]==null)
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
                    else if(row["ITEM_MRGN_PCT_TY"].ToString()=="")
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
                        sSHIPPED_UNIT_VOLUME = "'" + row["SHIPPED_UNIT_VOLUME"].ToString() +"'";
                    }
                    
                    string sLOCATION = row["LOCATION"].ToString();
                    string sPLN_VRSN = row["PLN_VRSN"].ToString();

                    DateTime loaddt = DateTime.Parse(row["LOAD_DATE"].ToString());

                    
                    string sLOAD_DATE = " TO_DATE('" + loaddt.ToString("MM/dd/yyyy") + "','DD-MMM-YY') ";
                    string sqlStatement = "INSERT INTO " + destTableName + 
                    " (DAY,DEMAND_PLAN,SLS_RTL,ITEM_MRGN,ITEM_MRGN_PCT_TY,SHIPPED_ORDERS,SHIPPED_UNIT_VOLUME,LOCATION,PLN_VRSN,LOAD_DATE) " +
                        "VALUES ("+ sDAY + ",'" + sDEMAND_PLAN + "','" + sSLS_RTL + "'," + sITEM_MRGN + "," + sITEM_MRGN_PCT_TY + "," + sSHIPPED_ORDERS + "," + sSHIPPED_UNIT_VOLUME 
                        + ",'" + sLOCATION + "','" + sPLN_VRSN + "'," + sLOAD_DATE + ")";

                    con2.Open();
                    OleDbCommand cmd1 = new OleDbCommand(sqlStatement, con2);
                    cmd1.ExecuteNonQuery();
                    con2.Close();
                }
            }
          
        }

        private bool AddDataToDatatable(IWorkbook workbook)
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
                        DataRow rw= Tabla.NewRow();
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
                        
                        

                        
                        Tabla.Rows.Add(rw);
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
                    IRow row = worksheet.GetRow( rowindex );//skip header

                    while (row != null)
                    {

                        string sDate = "";
                        var d = row.GetCell(0);
                        if (d.CellType== CellType.Numeric)
                        {
                            if (d.DateCellValue.ToString()!="")
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


                        var c = row.GetCell(arrHeader.Length-1);
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

            if (!DateTime.TryParse(sDate,out dt))
            {
                return false;
            }

            return true;
        }

        private bool IsValidVersion(string sVersion)
        {
            string tVer = sVersion.ToUpper().Trim();
            if (!(tVer=="LE01" || tVer == "LE02" || tVer == "LE03" || tVer == "LE04" || tVer == "LE05" || tVer == "LE06" || tVer == "LE07" || tVer == "LE08" 
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

            foreach(string sname in arrSheetName)
            {
                if (sname.ToLower().Trim()==sheetName.ToLower().Trim())
                {
                    return true;
                }
            }
           

            return false;
        }

        private bool AreHeadersValid(IWorkbook workbook,int maxColCount )
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
        private DataTable Excel_To_DataTable1(string Excelfilename)
        {            
            DataTable Tabla = null;
            try
            {
                if (System.IO.File.Exists(Excelfilename))
                {

                    IWorkbook workbook = null;  
                    ISheet worksheet = null;
                    string first_sheet_name = "";
                    
                    first_sheet_name = "MyTable";
                    Tabla = new DataTable(first_sheet_name);
                    Tabla.Rows.Clear();
                    Tabla.Columns.Clear();

                    using (FileStream FS = new FileStream(Excelfilename, FileMode.Open, FileAccess.Read))
                    {
                        workbook = WorkbookFactory.Create(FS);

                        for (int fileIndex = 0; fileIndex < workbook.NumberOfSheets; fileIndex++)
                        {
                            if (!workbook.IsSheetHidden(fileIndex))
                            {

                                worksheet = workbook.GetSheetAt(fileIndex);                       


                                for (int rowIndex = 0; rowIndex <= worksheet.LastRowNum; rowIndex++)
                                {
                                    DataRow NewReg = null;
                                    IRow row = worksheet.GetRow(rowIndex);
                                    IRow row2 = null;
                                    IRow row3 = null;

                                    if (rowIndex == 0)
                                    {
                                        row2 = worksheet.GetRow(rowIndex + 1);
                                        row3 = worksheet.GetRow(rowIndex + 2);
                                    }

                                    if (row != null) //null is when the row only contains empty cells 
                                    {
                                        if (rowIndex > 0)
                                        {
                                            NewReg = Tabla.NewRow();
                                        }
                                        int colIndex = 0;

                                        foreach (ICell cell in row.Cells)
                                        {
                                            object valorCell = null;
                                            string cellType = "";
                                            string[] cellType2 = new string[2];

                                            if (rowIndex == 0)
                                            {
                                                for (int i = 0; i < 2; i++)
                                                {
                                                    ICell cell2 = null;
                                                    if (i == 0) 
                                                    { 
                                                        cell2 = row2.GetCell(cell.ColumnIndex); 
                                                    }
                                                    else 
                                                    { 
                                                        cell2 = row3.GetCell(cell.ColumnIndex); 
                                                    }

                                                    if (cell2 != null)
                                                    {
                                                        switch (cell2.CellType)
                                                        {
                                                            case CellType.Blank: break;
                                                            case CellType.Boolean: cellType2[i] = "System.Boolean"; break;
                                                            case CellType.String: cellType2[i] = "System.String"; break;
                                                            case CellType.Numeric:
                                                                if (HSSFDateUtil.IsCellDateFormatted(cell2)) { cellType2[i] = "System.DateTime"; }
                                                                else
                                                                {
                                                                    cellType2[i] = "System.Double";
                                                                }
                                                                break;

                                                            case CellType.Formula:
                                                                bool continuar = true;
                                                                switch (cell2.CachedFormulaResultType)
                                                                {
                                                                    case CellType.Boolean: cellType2[i] = "System.Boolean"; break;
                                                                    case CellType.String: cellType2[i] = "System.String"; break;
                                                                    case CellType.Numeric:
                                                                        if (HSSFDateUtil.IsCellDateFormatted(cell2)) { cellType2[i] = "System.DateTime"; }
                                                                        else
                                                                        {
                                                                            try
                                                                            {
                                                                                //DETERMINAR SI ES BOOLEANO
                                                                                if (cell2.CellFormula == "TRUE()") { cellType2[i] = "System.Boolean"; continuar = false; }
                                                                                if (continuar && cell2.CellFormula == "FALSE()") { cellType2[i] = "System.Boolean"; continuar = false; }
                                                                                if (continuar) { cellType2[i] = "System.Double"; continuar = false; }
                                                                            }
                                                                            catch { }
                                                                        }
                                                                        break;
                                                                }
                                                                break;
                                                            default:
                                                                cellType2[i] = "System.String"; break;
                                                        }
                                                    }
                                                }

                                                if (cellType2[0] == cellType2[1]) 
                                                { 
                                                    cellType = cellType2[0]; 
                                                }
                                                else
                                                {
                                                    if (cellType2[0] == null) cellType = cellType2[1];
                                                    if (cellType2[1] == null) cellType = cellType2[0];
                                                    if (cellType == "") cellType = "System.String";
                                                }

                                                string colName = "Column_{0}";
                                                try 
                                                { 
                                                    colName = cell.StringCellValue; 
                                                }
                                                catch 
                                                { 
                                                    colName = string.Format(colName, colIndex); 
                                                }

                                                foreach (DataColumn col in Tabla.Columns)
                                                {
                                                    if (col.ColumnName == colName) colName = string.Format("{0}_{1}", colName, colIndex);
                                                }


                                                DataColumn codigo = new DataColumn(colName, System.Type.GetType(cellType));
                                                Tabla.Columns.Add(codigo); 
                                                colIndex++;
                                            }
                                            else
                                            {
                                                switch (cell.CellType)
                                                {
                                                    case CellType.Blank: valorCell = DBNull.Value; break;
                                                    case CellType.Boolean: valorCell = cell.BooleanCellValue; break;
                                                    case CellType.String: valorCell = cell.StringCellValue; break;
                                                    case CellType.Numeric:
                                                        if (HSSFDateUtil.IsCellDateFormatted(cell)) { valorCell = cell.DateCellValue; }
                                                        else { valorCell = cell.NumericCellValue; }
                                                        break;
                                                    case CellType.Formula:
                                                        switch (cell.CachedFormulaResultType)
                                                        {
                                                            case CellType.Blank: valorCell = DBNull.Value; break;
                                                            case CellType.String: valorCell = cell.StringCellValue; break;
                                                            case CellType.Boolean: valorCell = cell.BooleanCellValue; break;
                                                            case CellType.Numeric:
                                                                if (HSSFDateUtil.IsCellDateFormatted(cell)) { valorCell = cell.DateCellValue; }
                                                                else { valorCell = cell.NumericCellValue; }
                                                                break;
                                                        }
                                                        break;
                                                    default: valorCell = cell.StringCellValue; break;
                                                }
                                                if (cell.ColumnIndex <= Tabla.Columns.Count - 1) NewReg[cell.ColumnIndex] = valorCell;
                                            }
                                        }
                                    }
                                    if (rowIndex > 0) Tabla.Rows.Add(NewReg);
                                }
                            }
                            Tabla.AcceptChanges();
                        }
                    }
                }
                else
                {
                    throw new Exception("ERROR 404");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return Tabla;
        }

    }
}
