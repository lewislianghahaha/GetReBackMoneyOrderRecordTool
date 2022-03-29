using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace GetReBackMoneyOrderRecordTool
{
    public class Import
    {
        Tempdt tempdt=new Tempdt();

        /// <summary>
        /// EXcel导入
        /// </summary>
        /// <param name="fileAddress"></param>
        /// <returns></returns>
        public DataTable ImportExcelToDt(string fileAddress)
        {
            var dt = new DataTable();
            try
            {
                //使用NPOI技术进行导入EXCEL至DATATABLE
                var importExcelDt = OpenExcelToDataTable(fileAddress);
                //将从EXCEL过来的记录集为空的行清除
                dt = RemoveEmptyRows(importExcelDt);
            }
            catch (Exception)
            {
                dt.Columns.Clear();
                dt.Rows.Clear();
            }
            return dt;
        }

        /// <summary>
        /// 打开EXCEL并获取其内容
        /// </summary>
        /// <param name="fileAddress">EXCEL导入地址</param>
        /// <returns></returns>
        private DataTable OpenExcelToDataTable(string fileAddress)
        {
            IWorkbook wk;

            //获取动态生成绑定字段临时表
            var dt = tempdt.MakeImportTemp();
            using (var fsRead = File.OpenRead(fileAddress))
            {
                wk = new XSSFWorkbook(fsRead);
                //获取第一个sheet
                var sheet = wk.GetSheetAt(0);
                //定义列数(注:此总列次必须与EXCEL中的总列数一致) 同过dt获取总列数
                var colnum = dt.Columns.Count;

                //创建完标题后,开始从第二行起读取对应列的值(因为EXCEL模板前两行都是标题信息,从第三行开始才是内容)
                for (var r = 1; r <= sheet.LastRowNum; r++)
                {
                    var result = false;
                    var dr = dt.NewRow();
                    var row = sheet.GetRow(r);
                    if (row == null) continue;

                    //循环读取EXCEL中各行的单元格
                    for (var j = 0; j < colnum; j++)
                    {
                        var cell = row.GetCell(j);
                        var cellValue = GetCellValue(cell);
                        //当为空的就不获取
                        if (cellValue == string.Empty)
                        {
                            continue;
                        }
                        else
                        {
                            dr[j] = cellValue;
                        }

                        //全为空的就不取
                        if (dr[j].ToString() != "")
                        {
                            result = true;
                        }
                    }
                    if (result)
                    {
                        //将每行增加至Dt结果集内
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }

        /// <summary>
        /// 检查单元格的数据类型并获其中的值
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
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
                    if (DateUtil.IsCellDateFormatted(cell))//日期类型
                    {
                        return cell.DateCellValue.ToString();
                    }
                    else //其它数字
                    {
                        return cell.NumericCellValue.ToString();

                    }

                case CellType.Unknown: //无法识别类型
                default: //默认类型                    
                    return cell.ToString();
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
        ///  将从EXCEL导入的DATATABLE的空白行清空
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        protected DataTable RemoveEmptyRows(DataTable dt)
        {
            var removeList = new List<DataRow>();
            for (var i = 0; i < dt.Rows.Count; i++)
            {
                var isNull = true;
                for (var j = 0; j < dt.Columns.Count; j++)
                {
                    //将不为空的行标记为False
                    if (!string.IsNullOrEmpty(dt.Rows[i][j].ToString().Trim()))
                    {
                        isNull = false;
                    }
                }
                //将整行都为空白的记录进行记录
                if (isNull)
                {
                    removeList.Add(dt.Rows[i]);
                }
            }

            //将整理出来的所有空白行通过循环进行删除
            for (var i = 0; i < removeList.Count; i++)
            {
                dt.Rows.Remove(removeList[i]);
            }
            return dt;
        }
    }
}
