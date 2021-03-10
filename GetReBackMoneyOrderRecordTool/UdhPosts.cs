using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Xml;

namespace GetReBackMoneyOrderRecordTool
{
    public class UdhPosts
    {
        Tempdt tempdt=new Tempdt();
        ExportDt exportDt=new ExportDt();

        /// <summary>
        /// 返利单号获取返利单
        /// </summary>
        /// <param name="order"></param>
        /// <returns></returns>
        public static string GetOrderList(string order)
        {
            string ret = "";
            var param = new Dictionary<string, string>();
            param.Add("rebateno", order);
            ret = UWeb.Get("/rs/Rebates/getRebate",param);
            return ret;
        }

        /// <summary>
        /// 获取所有返利单
        /// </summary>
        /// <returns></returns>
        public static string GetAllOrder()
        {
            string result = "";
            var param = new Dictionary<string, string>();
            param.Add("pageindex","1");
            param.Add("pagesize","10");
            result = UWeb.Get("/rs/Rebates/getSummaryRebates", param);
            return result;
        }

        /// <summary>
        /// 根据返利单号获取返利使用记录
        /// </summary>
        /// <param name="order"></param>
        /// <returns></returns>
        public string GetUseOrderList(string order)
        {
            var result = "";
            var param = new Dictionary<string, string>();
            var resultdt = tempdt.MakeRebateRecordTemp().Clone();

            try
            {
                var orderlist = order.Split(',');

                for (var i = 0; i < orderlist.Length; i++)
                {
                    var rebateno = orderlist[i].Trim();
                    param.Remove("rebateno");
                    param.Add("rebateno", rebateno);
                    result = UWeb.Get("/rs/Rebates/getRebateRecord", param);

                    var dt = GetXmLtoDb(result).Copy();
                    resultdt.Merge(dt.Rows.Count == 0 ? MakeEmptyDt(rebateno) : dt);
                   // var a = result;
                }
                //当完成后将相关记录导出至EXCEL
                ExportExcel(resultdt);
            }
            catch (Exception ex)
            {
                result = ex.Message;
            }
            return result;
        }

        /// <summary>
        /// 根据时间范围获取返利单使用记录
        /// </summary>
        /// <returns></returns>
        public void GeteUseTimeRebateNoList()
        {
            var resultdt = tempdt.MakeRebateRecordTemp().Clone();

            var param = new Dictionary<string, string>();
            param.Remove("pageindex");
            param.Add("pageindex","1");           //页码
            param.Remove("pagesize");
            param.Add("pagesize","50");           //一页显示条数
            param.Remove("startdate");
            param.Add("startdate","2019-02-10");  //开始时间
            param.Remove("enddate");
            param.Add("enddate","2021-03-10");    //结束时间
            var result = UWeb.Get("/rs/Rebates/getRebateRecordsByDate", param);

            resultdt.Merge(GetXmlList(result));

            //当完成后将相关记录导出至EXCEL
            ExportExcel(resultdt);
        }

        /// <summary>
        /// 循环获取XML子节点内的指定节点信息
        /// </summary>
        /// <param name="xmlstring"></param>
        private DataTable GetXmlList(string xmlstring)
        {
            var dt = tempdt.MakeRebateRecordTemp().Clone();
            //创建xmlDocment对像并读取传输过来的xml字符串
            var xmldoc = new XmlDocument();
            xmldoc.LoadXml(xmlstring);
            
            //循环层级获取XML节点记录
            var xmlNode = xmldoc.DocumentElement;
            if (xmlNode != null)
                foreach (XmlNode node in xmlNode)
                {
                    if (node.Name == "data")
                    {
                        var xmlNode1 = node.ChildNodes;

                        foreach (XmlNode node1 in xmlNode1)
                        {
                            if (node1.Name == "rebateRecords")
                            {
                                var pXmlNodeList = node1.ChildNodes;
                                 
                                foreach (XmlNode p2 in pXmlNodeList)
                                {
                                    if (p2.Name == "item")
                                    {
                                        var pp = p2.ChildNodes;
                                        //根据对应的字段名称,获取其对应的节点值至DT.newrow内
                                        var newrow = dt.NewRow();
                                        for (var i = 0; i < dt.Columns.Count; i++)
                                        {
                                            foreach (XmlNode p3 in pp)
                                            {
                                                if (p3.Name != "cOrderNo" && p3.Name != "cRebateNo" && p3.Name != "fOrderRebateMoney" &&
                                                    p3.Name != "cRecordStatus" && p3.Name != "cRecordStatusName" && p3.Name != "dCreateDate" &&
                                                    p3.Name != "iSubmiterId") continue;

                                                if (p3.Name == dt.Columns[i].ColumnName)
                                                {
                                                    newrow[i] = p3.InnerText;
                                                }
                                            }
                                        }
                                        dt.Rows.Add(newrow);
                                    }
                                }
                            }
                        }
                    }
                }
            return dt;
        }

        /// <summary>
        /// 获取XML记录并生成DT
        /// </summary>
        /// <returns></returns>
        private DataTable GetXmLtoDb(string xmlstring)
        {
            var dt = tempdt.MakeRebateRecordTemp().Clone();

            try
            {
                var xmldoc = new XmlDocument();
                xmldoc.LoadXml(xmlstring);

                //注:SelectSingleNode("//response");  使用此函数才需要在节点前加//
                //var nodestring = "//cOrderNo|//cRebateNo|//fOrderRebateMoney|//cRecordStatus|//cRecordStatusName|//dCreateDate|//iSubmiterId";

                //var nodesname = "cOrderNo|cRebateNo|fOrderRebateMoney|cRecordStatus|cRecordStatusName|dCreateDate|iSubmiterId";

                XmlNode xmlNode = xmldoc.DocumentElement;
                if (xmlNode != null)
                {
                    foreach (XmlNode nodes in xmlNode)
                    {
                        if (nodes.Name == "data")
                        {
                            XmlNodeList xmlnode = nodes.ChildNodes;
                            foreach (XmlNode node1 in xmlnode)
                            {
                                if (node1.Name == "item")
                                {
                                    XmlNodeList p1 = node1.ChildNodes;

                                    var newrow = dt.NewRow();
                                    for (var i = 0; i < dt.Columns.Count; i++)
                                    {
                                        foreach (XmlNode p2 in p1)
                                        {
                                            if (p2.Name != "cOrderNo" && p2.Name != "cRebateNo" && p2.Name != "fOrderRebateMoney" &&
                                               p2.Name != "cRecordStatus" && p2.Name != "cRecordStatusName" && p2.Name != "dCreateDate" &&
                                               p2.Name != "iSubmiterId") continue;

                                            if (p2.Name == dt.Columns[i].ColumnName)
                                            {
                                                newrow[i] = p2.InnerText;
                                            }
                                        }
                                    }
                                    dt.Rows.Add(newrow);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var a = ex.Message;
            }
            return dt;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="order"></param>
        /// <returns></returns>
        private DataTable MakeEmptyDt(string order)
        {
            var dt = tempdt.MakeRebateRecordTemp().Clone();
            //只记录返利单号,其余为空
            var newrow = dt.NewRow();
            newrow[1] = order;
            dt.Rows.Add(newrow);
            return dt;
        }

        /// <summary>
        /// 导出EXCEL
        /// </summary>
        /// <param name="resultdt"></param>
        private void ExportExcel(DataTable resultdt)
        {
            //当完成后将相关记录导出至EXCEL
            var saveFileDialog = new SaveFileDialog { Filter = $"Xlsx文件|*.xlsx" };
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                var fileAdd = saveFileDialog.FileName;
                exportDt.ExportDtToExcel(fileAdd, resultdt);
            }
        }
    }
}
