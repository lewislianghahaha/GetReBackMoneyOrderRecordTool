using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetReBackMoneyOrderRecordTool
{
    public class Tempdt
    {
        /// <summary>
        /// 返利使用记录临时表
        /// </summary>
        /// <returns></returns>
        public DataTable MakeRebateRecordTemp()
        {
            var dt = new DataTable();
            for (var i = 0; i < 7; i++)
            {
                var dc = new DataColumn();
                switch (i)
                {
                    //U订货订单号
                    case 0:
                        dc.ColumnName = "cOrderNo";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    //U订货返利单号
                    case 1:
                        dc.ColumnName = "cRebateNo";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    //返利使用金额
                    case 2:
                        dc.ColumnName = "fOrderRebateMoney";
                        dc.DataType = Type.GetType("System.String"); 
                        break;
                    //使用状态编码
                    case 3:
                        dc.ColumnName = "cRecordStatus";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    //使用状态名称
                    case 4:
                        dc.ColumnName = "cRecordStatusName";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    //创建时间
                    case 5:
                        dc.ColumnName = "dCreateDate";
                        dc.DataType = Type.GetType("System.String"); 
                        break;
                    //创建人
                    case 6:
                        dc.ColumnName = "iSubmiterId";
                        dc.DataType = Type.GetType("System.String");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

        /// <summary>
        /// 导入EXCEL临时表
        /// </summary>
        /// <returns></returns>
        public DataTable MakeImportTemp()
        {
            var dt = new DataTable();
            for (var i = 0; i < 1; i++)
            {
                var dc = new DataColumn();
                switch (i)
                {
                    //U订货返利单号
                    case 0:
                        dc.ColumnName = "cRebateNo";
                        dc.DataType = Type.GetType("System.String");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

    }
}
