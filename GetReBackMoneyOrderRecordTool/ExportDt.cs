using System;
using System.Data;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace GetReBackMoneyOrderRecordTool
{
    public class ExportDt
    {
        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="fileAdd"></param>
        /// <param name="sourcedt"></param>
        /// <returns></returns>
        public bool ExportDtToExcel(string fileAdd, DataTable sourcedt)
        {
            var result = true;
            var sheetcount = 0;  //记录所需的sheet页总数
            var rownum = 1;

            try
            {
                //声明一个WorkBook
                var xssfWorkbook = new XSSFWorkbook();

                //执行sheet页(注:1)先列表temp行数判断需拆分多少个sheet表进行填充; 以一个sheet表有100W行记录填充为基准)
                sheetcount = sourcedt.Rows.Count % 1000000 == 0 ? sourcedt.Rows.Count / 1000000 : sourcedt.Rows.Count / 1000000 + 1;

                //i为EXCEL的Sheet页数ID
                for (var i = 1; i <= sheetcount; i++)
                {
                    //创建sheet页
                    var sheet = xssfWorkbook.CreateSheet("Sheet" + i);
                    //创建"标题行"
                    var row = sheet.CreateRow(0);

                    //创建sheet页各列标题
                    for (var j = 0; j < sourcedt.Columns.Count; j++)
                    {
                        //设置列宽度
                        sheet.SetColumnWidth(j, (int)((20 + 0.72) * 256));
                        //创建标题
                        switch (j)
                        {
                            case 0:
                                row.CreateCell(j).SetCellValue("U订货订单号");
                                break;
                            case 1:
                                row.CreateCell(j).SetCellValue("U订货返利单号");
                                break;
                            case 2:
                                row.CreateCell(j).SetCellValue("返利使用金额");
                                break;
                            case 3:
                                row.CreateCell(j).SetCellValue("使用状态编码");
                                break;
                            case 4:
                                row.CreateCell(j).SetCellValue("使用状态名称");
                                break;
                            case 5:
                                row.CreateCell(j).SetCellValue("创建时间");
                                break;
                            case 6:
                                row.CreateCell(j).SetCellValue("创建人");
                                break;
                        }
                    }

                    //计算进行循环的起始行
                    var startrow = (i - 1) * 1000000;
                    //计算进行循环的结束行
                    var endrow = i == sheetcount ? sourcedt.Rows.Count : i * 1000000;

                    //每一个sheet表显示100000行  
                    for (var j = startrow; j < endrow; j++)
                    {
                        //创建行
                        row = sheet.CreateRow(rownum);
                        //循环获取DT内的列值记录
                        for (var k = 0; k < sourcedt.Columns.Count; k++)
                        {
                            if (Convert.ToString(sourcedt.Rows[j][k]) == "") continue;
                            else
                            {
                                    row.CreateCell(k, CellType.String).SetCellValue(Convert.ToString(sourcedt.Rows[j][k]));
                            }
                        }
                        rownum++;
                    }
                    //当一个SHEET页填充完毕后,需将变量初始化
                    rownum = 1;
                }

                //写入数据
                var file = new FileStream(fileAdd, FileMode.Create);
                xssfWorkbook.Write(file);
                file.Close();           //关闭文件流
                xssfWorkbook.Close();   //关闭工作簿
                file.Dispose();         //释放文件流
            }
            catch (Exception)
            {
                result = false;
            }
            return result;
        }
    }
}
