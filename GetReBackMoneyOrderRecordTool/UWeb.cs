using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;

namespace GetReBackMoneyOrderRecordTool
{
    public class UWeb
    {
        #region 
        //public static string appkey = "62ff2134ba1f6916017aa1e312b74fbfac1cac8a";
        //public static string secret = "bd07e6f69217e6ecbc0118088158987061c7c1e5";
        //public static string token = "!*LsDMzUIFD8un@jNdNg~yPX3zsGR~WEzgJ4cpaMcTFJTRdtD65m05psKGbU5l0x~o-";
        #endregion

        public static string appkey = "e8bd3c320a926ba47fdae7c39f95b39c97ae5451";  //
        public static string secret = "1ec3cdfffa96d2922944897b27ee77d13f72ba80";
        public static string token = "!*iuKaR7Kgc7JBkAoB7~Jnw@e@GinZ~mzdWZZyWZZBGKs*-";//"!*iuKaR7Kgc7JBkAoB7~Jnw@zwVnbvq0golWNtIJREsUDRdtD65m05psKGbU5l0x~o-";
        public static string url = "https://api.udinghuo.cn";  

        public static string CreateSign(IDictionary<string, string> parameters, string secret)
        {
            parameters.Remove("sign");
            IDictionary<string, string> dictionary = new SortedDictionary<string, string>(parameters);
            IEnumerator<KeyValuePair<string, string>> enumerator = dictionary.GetEnumerator();
            StringBuilder builder = new StringBuilder(secret);
            while (enumerator.MoveNext())
            {
                KeyValuePair<string, string> current = enumerator.Current;
                string key = current.Key;
                KeyValuePair<string, string> pair2 = enumerator.Current;
                string str2 = pair2.Value;
                if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(str2))
                {
                    builder.Append(key).Append(str2);
                }
            }
            builder.Append(secret);
            byte[] buffer = System.Security.Cryptography.MD5.Create().ComputeHash(Encoding.UTF8.GetBytes(builder.ToString()));
            StringBuilder builder2 = new StringBuilder();
            for (int i = 0; i < buffer.Length; i++)
            {
                string str3 = buffer[i].ToString("X");
                if (str3.Length == 1)
                {
                    builder2.Append("0");
                }
                builder2.Append(str3);
            }
            return builder2.ToString();
        }


        public static string Get(string method, IDictionary<string, string> param)
        {
            string cururl = url;
            string str3;
            try
            {
                new XmlDocument();

                param.Remove("appkey");
                param.Add("appkey", appkey);
                param.Remove("token");
                param.Add("token", token);
                param.Remove("format");
                param.Add("format", "xml");
                param.Remove("sign");
                param.Add("sign", CreateSign(param, secret));

                string input = string.Empty;
                GC.Collect();
                cururl += method + "?";
                StringBuilder builder = new StringBuilder();
                bool flag = false;
                foreach (string str2 in param.Keys)
                {
                    if (flag)
                    {
                        builder.Append("&");
                    }
                    builder.Append(str2);
                    builder.Append("=");
                    builder.Append(Uri.EscapeDataString(param[str2]));
                    flag = true;
                }
                cururl += builder.ToString();
                HttpWebRequest request = null;
                HttpWebResponse response = null;
                try
                {
                    ServicePointManager.DefaultConnectionLimit = 50;
                    request = (HttpWebRequest)WebRequest.Create(cururl);
                    request.Method = "GET";
                    response = (HttpWebResponse)request.GetResponse();
                    StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
                    input = reader.ReadToEnd().Trim();
                    reader.Close();

                    if (reader != null)
                    {
                        reader.Close();
                    }
                }
                catch (ThreadAbortException exception)
                {

                    Thread.ResetAbort();
                }
                catch (WebException exception2)
                {

                    throw new Exception(exception2.ToString());
                }
                catch (Exception exception3)
                {

                    throw new Exception(exception3.ToString());
                }
                finally
                {
                    if (response != null)
                    {
                        response.Close();
                    }
                    if (request != null)
                    {
                        request.Abort();
                    }
                }

                str3 = Regex.Replace(input, @"[\x00-\x08\x0b-\x0c\x0e-\x1f]", "");
            }
            catch (WebException exception4)
            {

                throw exception4;
            }
            return str3;
        }


        public static string Post(string method, IDictionary<string, string> param)
        {
            string str3;
            string cururl = url;
            try
            {
                new XmlDocument();
                param.Remove("appkey");
                param.Add("appkey", appkey);
                param.Remove("token");
                param.Add("token", token);
                param.Remove("format");
                param.Add("format", "xml");
                param.Remove("sign");
                param.Add("sign", CreateSign(param, secret));
                string input = string.Empty;
                GC.Collect();
                cururl += method + "?";
                StringBuilder builder = new StringBuilder();
                bool flag = false;
                foreach (string str2 in param.Keys)
                {
                    if (flag)
                    {
                        builder.Append("&");
                    }
                    builder.Append(str2);
                    builder.Append("=");
                    builder.Append(Uri.EscapeDataString(param[str2]));
                    flag = true;
                }
                cururl += builder.ToString();
                HttpWebRequest request = null;
                HttpWebResponse response = null;
                try
                {
                    ServicePointManager.DefaultConnectionLimit = 50;
                    request = (HttpWebRequest)WebRequest.Create(cururl);
                    request.Method = "POST";
                    response = (HttpWebResponse)request.GetResponse();
                    StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
                    input = reader.ReadToEnd().Trim();
                    reader.Close();

                    if (reader != null)
                    {
                        reader.Close();
                    }
                }
                catch (ThreadAbortException exception)
                {

                    Thread.ResetAbort();
                }
                catch (WebException exception2)
                {

                    throw new Exception(exception2.ToString());
                }
                catch (Exception exception3)
                {

                    throw new Exception(exception3.ToString());
                }
                finally
                {
                    if (response != null)
                    {
                        response.Close();
                    }
                    if (request != null)
                    {
                        request.Abort();
                    }
                }

                str3 = Regex.Replace(input, @"[\x00-\x08\x0b-\x0c\x0e-\x1f]", "");
            }
            catch (WebException exception4)
            {
                throw exception4;
            }
            return str3;
        }

        public static string FGet()
        {
            string url = "https://api.udinghuo.cn/rs/Orders/getSummaryOrders?";

            Dictionary<string, string> param = new Dictionary<string, string>();

            string str3;
            try
            {
                //param.Remove("appkey");
                param.Add("appkey", appkey);
                //param.Remove("token");
                param.Add("token", token);
                //param.Remove("format");
                param.Add("format", "xml");
                //param.Remove("sign");
                param.Add("sign", CreateSign(param, secret));
                string input = string.Empty;
                // this.TransDetail.dCreated = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                GC.Collect();
                //  url = url + method + "?";
                StringBuilder builder = new StringBuilder();
                bool flag = false;
                foreach (string str2 in param.Keys)
                {
                    if (flag)
                    {
                        builder.Append("&");
                    }
                    builder.Append(str2);
                    builder.Append("=");
                    builder.Append(Uri.EscapeDataString(param[str2]));
                    flag = true;
                }
                url = url + builder.ToString();

                HttpWebRequest request = null;
                HttpWebResponse response = null;
                try
                {
                    ServicePointManager.DefaultConnectionLimit = 50;
                    request = (HttpWebRequest)WebRequest.Create(url);
                    request.Method = "GET";
                    response = (HttpWebResponse)request.GetResponse();
                    StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
                    input = reader.ReadToEnd().Trim();
                    reader.Close();

                    if (reader != null)
                    {
                        reader.Close();
                    }
                }
                catch (ThreadAbortException exception)
                {

                    Thread.ResetAbort();
                }
                catch (WebException exception2)
                {

                    throw new Exception(exception2.ToString());
                }
                catch (Exception exception3)
                {

                    throw new Exception(exception3.ToString());
                }
                finally
                {
                    if (response != null)
                    {
                        response.Close();
                    }
                    if (request != null)
                    {
                        request.Abort();
                    }
                }

                str3 = Regex.Replace(input, @"[\x00-\x08\x0b-\x0c\x0e-\x1f]", "");
            }
            catch (WebException exception4)
            {

                throw exception4;
            }
            return str3;
        }
     }
}
