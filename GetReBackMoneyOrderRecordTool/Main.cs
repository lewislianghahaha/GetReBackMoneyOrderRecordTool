using System;
using System.Data;
using System.Windows.Forms;

namespace GetReBackMoneyOrderRecordTool
{
    public partial class Main : Form
    {
        UdhPosts udh=new UdhPosts();
        Import import=new Import();

        public Main()
        {
            InitializeComponent();
            OnShow();
        }

        private void OnShow()
        {
            var result = string.Empty;
            var uorderlist = string.Empty;
            //中转判断值
            var tempstring = string.Empty;

            try
            {
                //导入所需的‘返利单号'信息
                var openFileDialog = new OpenFileDialog { Filter = $"Xlsx文件|*.xlsx" };
                if (openFileDialog.ShowDialog() != DialogResult.OK) return;

                var dt = import.ImportExcelToDt(openFileDialog.FileName);

                //将获取到的DT记录整理并最终赋值给uorderlist
                foreach (DataRow rows in dt.Rows)
                {
                    if (string.IsNullOrEmpty(uorderlist))
                    {
                        uorderlist = Convert.ToString(rows[0]);
                        tempstring = Convert.ToString(rows[0]);
                    }
                    else
                    {
                        if (tempstring != Convert.ToString(rows[0]))
                        {
                            uorderlist += "," + Convert.ToString(rows[0]);
                            tempstring = Convert.ToString(rows[0]);
                        }
                    }
                }
                result = udh.GetUseOrderList(uorderlist);
                //result = UdhPosts.GetOrderList("UF-232c768161622003180001");
                //result = UdhPosts.GetAllOrder();
                //"UF-23d6e16171162103050001"  "UF-23d6e0d1716e2102010001,"  uorderlist  "UF-24fe3ee1401a2102010001"
                //udh.GeteUseTimeRebateNoList();
                //udh.GetOrderMessage("UO-4da68d0190392104110001");
            }
            catch (Exception ex)
            {
                result = ex.Message;
            }
        }
    }
}
