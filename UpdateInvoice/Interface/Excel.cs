using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace UpdateInvoice.Interface
{
    public class Excel
    {

        #region  临时表(销售发票使用)

        private string _SearTable = @"
                                        
                                                  SELECT TOP 1 b.FBillNo,a.FEntryID,a.FSourceEntryID,a.FSourceBillNo,a.FSourceInterId,a.FSourceTranType
                                                  from dbo.ICSaleEntry a
                                                  INNER JOIN dbo.ICSale b ON a.FInterID=b.FInterID;
                                           
                                     ";

        #endregion

        #region  更新语句(销售发票使用)

        private string _UpdateInvoice = @"
                                                    UPDATE ice set ice.fsourceEntryID=@FsourceEntryid,ice.fsourcebillNo=@Fsourcebillno,ice.FSourceInterId=@Fsourceinterid,ice.FSourceTranType=@FSourceTranType
                                                    FROM icsaleentry ice
                                                    INNER JOIN icsale ic ON ic.finterid=ice.finterid
                                                    WHERE ice.fentryid=@Fentryid      --行号
                                                    AND ic.fbillno=@Fbillno;          --发票号码
                                            ";


        #endregion

        #region   临时表(采购发票使用)

        private string _SearTable1 = @"

                                                SELECT TOP 1 b.FBillNo,a.FEntryID,a.FSourceEntryID,a.FSourceBillNo,a.FSourceInterId,a.FSourceTranType
                                                from dbo.ICPurchaseEntry a
                                                INNER JOIN dbo.ICPurchase b ON a.FInterID=b.FInterID;
                                               
                                        ";

        #endregion

        #region  更新语句(采购发票使用)

        private string _UpdatePurchase = @"
                                                
                                                 UPDATE ice set ice.fsourceEntryID=@FsourceEntryid,ice.fsourcebillNo=@Fsourcebillno,ice.FSourceInterId=@Fsourceinterid,ice.FSourceTranType=@FSourceTranType
                                                 FROM ICPurchaseEntry ice
                                                 INNER JOIN ICPurchase ic ON ic.finterid=ice.finterid
                                                 WHERE ice.fentryid=@Fentryid      --行号
                                                 AND ic.fbillno=@Fbillno;          --发票号码
                                                        
                                          ";
        #endregion

        #region  更新前先检查记录是否在销售发票中存在

        private string _SearErrorRecord_Sales = @"
                                                    SELECT  count(*) as tcount
                                                    FROM icsaleentry ice
                                                    inner join icsale ic on ic.finterid=ice.finterid
                                                    where ice.fentryid={0}
                                                    and ic.fbillno='{1}';
                                                 ";

        #endregion

        #region   更新前先检查记录是否在采购发票中存在

        private string _SearchErrorRecord_PO = @"
                                                   select count(*) as tcount
                                                   from ICPurchaseEntry ice
                                                   inner join ICPurchase ic on ic.finterid=ice.finterid
                                                   where ice.fentryid={0}
                                                   and ic.fbillno='{1}';
                                                ";

        #endregion

        private DataTable _errorTable;

        public DataTable ErrorTable
        {
            get { return _errorTable; }
        }

        /// <summary>
        /// 打开EXCEL并返回结果为DataSet
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public DataTable OpenExcel(string filename)
        {
            var dt = new DataTable();
            var importExcelDt = new DataTable();

            try
            {
                //var pubs = ConfigurationManager.ConnectionStrings["OleDbConStr"];  //读取配置文件
                //var conSplit = pubs.ConnectionString.Split(';');
                //var strcon = conSplit[0] + ";" + string.Format(conSplit[1], filename) + ";" + conSplit[2] + ";" + conSplit[3];

                //var con = new OleDbConnection(strcon);          //建立连接
                //const string strSql = "select * from [Sheet1$]";//表名的写法也应注意不同，对应的excel表为sheet1，在这里要在其后加美元符号$，并用中括号
                //var cmd = new OleDbCommand(strSql, con);        //建立要执行的命令
                //var da = new OleDbDataAdapter(cmd);             //建立数据适配器
                //da.Fill(ds);

                //change date:2018-07-03 使用NPOI技术进行导入EXCEL至DATATABLE
                importExcelDt = OpenExcelToDataTable(filename);
            }
            catch (Exception ex)
            {
                throw (new Exception(ex.Message));
            }
            //将从EXCEL里导入的整行空白行清除
            dt = RemoveEmptyRows(importExcelDt);
            return dt;
        }

        /// <summary>
        /// 读取EXCEL内容到DATATABLE内
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        private DataTable OpenExcelToDataTable(string filename)
        {
            IWorkbook wk;
            var dt = new DataTable();
            using (var fsRead = File.OpenRead(filename))
            {
                wk = new XSSFWorkbook(fsRead);
                //获取第一个sheet
                var sheet = wk.GetSheetAt(0);
                //获取第一行
                var hearRow = sheet.GetRow(0);
                //创建列标题
                for (int i = hearRow.FirstCellNum; i < hearRow.Cells.Count; i++)
                {
                    var dataColumn = new DataColumn();
                    switch (i)
                    {
                        case 0:
                            dataColumn.ColumnName = "行号";
                            break;
                        case 1:
                            dataColumn.ColumnName = "发票号码";
                            break;
                        case 2:
                            dataColumn.ColumnName = "源单单号";
                            break;
                        case 3:
                            dataColumn.ColumnName = "源单类型";
                            break;
                        case 4:
                            dataColumn.ColumnName = "源单内码";
                            break;
                        case 5:
                            dataColumn.ColumnName = "源单分录";
                            break;
                    }
                    dt.Columns.Add(dataColumn);
                }

                //创建完标题后,开始从第二行起读取对应列的值
                for (int r = 1; r <= sheet.LastRowNum; r++)
                {
                    bool result = false;
                    var dr = dt.NewRow();
                    //获取当前行
                    var row = sheet.GetRow(r);
                    //读取每列
                    for (int j = 0; j < row.Cells.Count; j++)
                    {
                        //循环获取行中的单元格
                        var cell = row.GetCell(j);
                        //循环获取行中的单元格的值
                        //dr[j] = cell.ToString();
                        dr[j] = GetCellValue(cell);
                        //全为空就不取
                        if (dr[j].ToString() != "")
                        {
                            result = true;
                        }
                    }
                    if (result == true)
                    {
                        //把每行增加到DataTable
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }

        //检查单元格的值
        private static string GetCellValue(ICell cell)
        {
            if (cell == null)
                return string.Empty;
            switch (cell.CellType)
            {
                case CellType.Blank: //空数据类型 这里类型注意一下，不同版本NPOI大小写可能不一样,有的版本是Blank（首字母大写)
                    return string.Empty;
                case CellType.Boolean: //bool类型
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    return cell.ErrorCellValue.ToString();
                case CellType.Numeric: //数字类型
                    if (HSSFDateUtil.IsCellDateFormatted(cell))//日期类型
                    {
                        return cell.DateCellValue.ToString();
                    }
                    else //其它数字
                    {
                        return cell.NumericCellValue.ToString();
                    }
                case CellType.Unknown: //无法识别类型
                default: //默认类型                    
                    return cell.ToString();//
                case CellType.String: //string 类型
                    return cell.StringCellValue;
                case CellType.Formula: //带公式类型
                    try
                    {
                        var e = new XSSFFormulaEvaluator(cell.Sheet.Workbook);
                        e.EvaluateInCell(cell);
                        return cell.ToString();
                    }
                    catch
                    {
                        return cell.NumericCellValue.ToString();
                    }
            }
        }

        /// <summary>
        /// 将从EXCEL导入的DATATABLE的空白行清空
        /// </summary>
        /// <param name="dt"></param>
        protected DataTable RemoveEmptyRows(DataTable dt)
        {
            var removeList = new List<DataRow>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                bool isNull = true;
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    //将不为空的行标记为False
                    if (!string.IsNullOrEmpty(dt.Rows[i][j].ToString().Trim()))
                    {
                        isNull = false;
                    }
                }
                //将整行都为空白的记录
                if (isNull)
                {
                    removeList.Add(dt.Rows[i]);
                }
            }

            //将整理出来的所有空白行通过循环进行删除
            for (int i = 0; i < removeList.Count; i++)
            {
                dt.Rows.Remove(removeList[i]);
            }
            return dt;
        }


        /// <summary>
        /// 验证能成功更新的记录并生成一个新的DataTable
        /// </summary>
        /// <param name="resourceTable"></param>
        /// <param name="clickid"></param>
        /// <returns></returns>
        public DataTable VaildCanImportTable(DataTable resourceTable, int clickid)
        {
            var dt = new DataTable();
            var failddt = new DataTable();

            //对DT创建自定义列(包括列名及数据类型) 接收能进行更新
            dt.Columns.Add("行号", Type.GetType("System.Int32"));
            dt.Columns.Add("发票号码", Type.GetType("System.String")); 
            dt.Columns.Add("源单单号", Type.GetType("System.String")); 
            dt.Columns.Add("源单类型", Type.GetType("System.Int32"));
            dt.Columns.Add("源单内码", Type.GetType("System.Int32"));
            dt.Columns.Add("源单分录", Type.GetType("System.Int32"));

            //接收不能更新的信息
            failddt.Columns.Add("行号", Type.GetType("System.Int32"));
            failddt.Columns.Add("发票号码", Type.GetType("System.String"));
            failddt.Columns.Add("源单单号", Type.GetType("System.String"));
            failddt.Columns.Add("源单类型", Type.GetType("System.Int32"));
            failddt.Columns.Add("源单内码", Type.GetType("System.Int32"));
            failddt.Columns.Add("源单分录", Type.GetType("System.Int32"));

            var sqlScript = clickid == 0 ? _SearErrorRecord_Sales : _SearchErrorRecord_PO;

            try
            {
                var pubs = ConfigurationManager.ConnectionStrings["Connstring"];        //读取配置文件 
                using (var conn = new SqlConnection(pubs.ConnectionString)) //读取配置文件中的连接字符串并使用
                {
                    for (int i = 0; i <= resourceTable.Rows.Count - 1; i++)
                    {
                        conn.Open();
                        var sqlcommand = new SqlCommand(string.Format(sqlScript, resourceTable.Rows[i][0], resourceTable.Rows[i][1]), conn);
                        var sqlcom = sqlcommand.ExecuteReader();
                        sqlcom.Read();
                        var tcount = sqlcom.GetInt32(0);
                        if (tcount != 0)
                        {
                            var newrow = dt.NewRow();
                            newrow["行号"] = resourceTable.Rows[i][0];
                            newrow["发票号码"] = resourceTable.Rows[i][1];
                            newrow["源单单号"] = resourceTable.Rows[i][2];
                            newrow["源单类型"] = resourceTable.Rows[i][3];
                            newrow["源单内码"] = resourceTable.Rows[i][4];
                            newrow["源单分录"] = resourceTable.Rows[i][5];
                            dt.Rows.Add(newrow);
                        }
                        else
                        {
                            var newrow = failddt.NewRow();
                            newrow["行号"] = resourceTable.Rows[i][0];
                            newrow["发票号码"] = resourceTable.Rows[i][1];
                            newrow["源单单号"] = resourceTable.Rows[i][2];
                            newrow["源单类型"] = resourceTable.Rows[i][3];
                            newrow["源单内码"] = resourceTable.Rows[i][4];
                            newrow["源单分录"] = resourceTable.Rows[i][5];
                            failddt.Rows.Add(newrow);
                        }
                        conn.Close();
                    }
                    _errorTable = failddt;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return dt;
        }

        /// <summary>
        /// 将从DataGridView获取的数据进行更新
        /// </summary>
        /// <param name="dtTable"></param>
        /// <param name="clickid"></param>
        public void ImportExcel(DataTable dtTable, int clickid)
        {
            //var result = new Hashtable();
            String sqlScript, updatescript;

            try
            {
                var pubs = ConfigurationManager.ConnectionStrings["Connstring"];        //读取配置文件 
                using (var conn = new SqlConnection(pubs.ConnectionString))             //读取配置文件中的连接字符串并使用
                {
                    var sqlDataAdapter = new SqlDataAdapter();
                    var ds = new DataSet();

                    conn.Open(); //打开连接

                    //先建一个临时表(用于将从DataGridview的数据源放到此表并最后完成更新) 注:临时表里要显示的字段是包括了后面更新语句要用到的所有字段,包括更新字段及查询字段

                    sqlScript = clickid == 0 ? _SearTable : _SearTable1;

                    sqlDataAdapter.SelectCommand = new SqlCommand(sqlScript, conn);
                    //将记录集填充至Dataset内
                    sqlDataAdapter.Fill(ds);

                    //建立更新相关设置
                    updatescript = clickid == 0 ? _UpdateInvoice : _UpdatePurchase;

                    sqlDataAdapter.UpdateCommand = new SqlCommand(updatescript, conn);

                    sqlDataAdapter.UpdateCommand.Parameters.Add("@Fentryid", SqlDbType.Int, 8, "Fentryid");                   //行号
                    sqlDataAdapter.UpdateCommand.Parameters.Add("@fbillno", SqlDbType.NVarChar, 255, "fbillno");              //发票号码
                    sqlDataAdapter.UpdateCommand.Parameters.Add("@fsourcebillNo", SqlDbType.NVarChar, 255, "fsourcebillNo");  //源单单号
                    sqlDataAdapter.UpdateCommand.Parameters.Add("@FSourceTranType", SqlDbType.Int, 8, "FSourceTranType");     //源单类型
                    sqlDataAdapter.UpdateCommand.Parameters.Add("@FSourceInterId", SqlDbType.Int, 8, "FSourceInterId");       //源单内码
                    sqlDataAdapter.UpdateCommand.Parameters.Add("@fsourceEntryID", SqlDbType.Int, 8, "fsourceEntryID");       //源单分录
                    sqlDataAdapter.UpdateCommand.UpdatedRowSource = UpdateRowSource.None;
                    sqlDataAdapter.UpdateBatchSize = 0;

                    //开始进行更新(注:这样的操作所更新的行会受临时表的行数限制) 其实就是对要插入的表内的内容进行更改
                    for (int i = 0; i <= dtTable.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < 1; j++)
                        {
                            ds.Tables[0].Rows[j].BeginEdit();
                            ds.Tables[0].Rows[j]["Fentryid"] = dtTable.Rows[i][0];         //行号
                            ds.Tables[0].Rows[j]["fbillno"] = dtTable.Rows[i][1];          //发票号码
                            ds.Tables[0].Rows[j]["fsourcebillNo"] = dtTable.Rows[i][2];    //源单单号
                            ds.Tables[0].Rows[j]["FSourceTranType"] = dtTable.Rows[i][3];  //源单类型
                            ds.Tables[0].Rows[j]["FSourceInterId"] = dtTable.Rows[i][4];   //源单内码
                            ds.Tables[0].Rows[j]["fsourceEntryID"] = dtTable.Rows[i][5];   //源单分录
                            ds.Tables[0].Rows[j].EndEdit();
                        }
                        sqlDataAdapter.Update(ds.Tables[0]);
                    }

                    //完成更新后将相关内容清空
                    ds.Tables[0].Clear();
                    sqlDataAdapter.Dispose();
                    ds.Dispose();
                    //关闭连接 
                    conn.Close();
                }
               // result["Code"] = "0";    //当正常更新后的结果为0
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
                //result["Code"] = ex.Message;
            }
           // return result;
        }

        /// <summary>
        /// 导出EXCEL
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        public Hashtable ExportExcel(string filename, DataTable dt)
        {
            var result = new Hashtable();

            try
            {
                //声明一个WorkBook(XSSFWorkbook类是用于导入导出07版本EXCEL的 而HSSFWorkbook类是用于导入导出03版本的EXCEL)
                var xssfWorkbook = new XSSFWorkbook();

                //为WorkBook创建work(创建工作表)
                var sheet = xssfWorkbook.CreateSheet("Sheet1");
                //创建第一行并对第一行内各列值进行设置
                var row = sheet.CreateRow(0);

                for (int l = 0; l < dt.Columns.Count; l++)
                {
                    //设置列宽
                    sheet.SetColumnWidth(l, (int)((20 + 0.72) * 256));

                    switch (l)
                    {
                        case 0:
                            row.CreateCell(l).SetCellValue("行号");
                            break;
                        case 1:
                            row.CreateCell(l).SetCellValue("发票号码");
                            break;
                        case 2:
                            row.CreateCell(l).SetCellValue("源单单号");
                            break;
                        case 3:
                            row.CreateCell(l).SetCellValue("源单类型");
                            break;
                        case 4:
                            row.CreateCell(l).SetCellValue("源单内码");
                            break;
                        case 5:
                            row.CreateCell(l).SetCellValue("源单分录");
                            break;
                    }
                }

                //创建每一行及对每行内的列进行赋值
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    //从第二行开始,因为第一行已作标题使用
                    row = sheet.CreateRow(j + 1);
                    //row.CreateCell(0).SetCellValue(j + 1);

                    for (int k = 0; k < dt.Columns.Count; k++)
                    {
                        row.CreateCell(k).SetCellValue(dt.Rows[j][k].ToString());
                        // row.CreateCell(k + 1).SetCellValue(dt.Rows[i * 10000 + j][k].ToString());
                    }
                }
                //写入数据
                var file = new FileStream(filename, FileMode.Create);
                xssfWorkbook.Write(file);
                file.Close();
                result["Code"] = 0;
            }
            catch (Exception ex)
            {
                result["Code"] = ex.Message;
            }

            return result;
        }

        //对表进行创建新行操作(注:这样的操作就是所更新的行不会受临时表的行数限制)插入功能使用;需配合InsertCommand属性
        //var newrow =dt.NewRow();
        //newrow["Fentryid"] = dtTable.Rows[i][0];
        //newrow["fbillno"] = dtTable.Rows[i][1];
        //newrow["fsourcebillNo"] = dtTable.Rows[i][2];
        //newrow["FSourceTranType"] = dtTable.Rows[i][3];
        //newrow["FSourceInterId"] = dtTable.Rows[i][4];
        //newrow["fsourceEntryID"] = dtTable.Rows[i][5];

        ////将新增的行添加至指定的DataTable内
        //dt.Rows.Add(newrow);
    }
}
