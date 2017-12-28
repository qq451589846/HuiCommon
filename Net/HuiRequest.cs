using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Net.Http;

namespace Common.Net
{
    public class HuiRequest
    {
        string _strempty = string.Empty;
        public int timeout { get; set; } = 5000;
        private static readonly HttpClient _client = new Lazy<HttpClient>(() => new HttpClient()).Value;


        public string ExecuteAsGet(string url)
        {
            if (CheckArgument(url))
            {
                _client.Timeout = new TimeSpan(5000);
                return _client.GetStringAsync(url).Result;
            }
            return _strempty;
        }

        /// <summary>
        /// Post方式访问地址
        /// </summary>
        /// <param name="url">地址</param>
        /// <param name="strparama">参数</param>
        /// <param name="dict">（可选）Headers字典，默认添加Content-Type application/x-www-form-urlencoded;charset=utf-8</param>
        /// <returns></returns>
        public string ExecuteAsPost(string url, string strparama, Dictionary<string, string> dict = null)
        {
            if (CheckArgument(url))
            {
                using (var content = new StringContent(strparama))
                {
                    content.Headers.Clear();
                    if (dict != null)
                    {
                        foreach (var item in dict)
                        {
                            content.Headers.Add(item.Key, item.Value);
                        }
                    }
                    else
                    {
                        content.Headers.Add("Content-Type", "application/x-www-form-urlencoded;charset=utf-8");
                    }

                    _client.Timeout = new TimeSpan(5000);
                    var resp = _client.PostAsync(url, content).Result;
                    return resp.Content.ReadAsStringAsync().Result;
                }
            }
            return _strempty;
        }

        /// <summary>
        /// 获取请求的反馈信息
        /// </summary>
        /// <param name="url"></param>
        /// <param name="bData">参数字节数组</param>
        /// <returns></returns>
        public string doPostRequest(string url, string param, out string status)
        {
            string strResult = string.Empty;
            status = string.Empty;
            try
            {
                ServicePointManager.Expect100Continue = false;//远程服务器返回错误: (417) Expectation failed 异常源自HTTP1.1协议的一个规范： 100(Continue)
                HttpWebRequest hwRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
                hwRequest.Timeout = timeout;
                hwRequest.Method = "POST";
                hwRequest.ContentType = "application/x-www-form-urlencoded;charset=utf-8";
                byte[] bData = Encoding.UTF8.GetBytes(param);
                hwRequest.ContentLength = bData.Length;
                using (var smWrite = hwRequest.GetRequestStream())
                {
                    smWrite.Write(bData, 0, bData.Length);
                }

                using (var hwResponse = (HttpWebResponse)hwRequest.GetResponse())
                {
                    status = hwResponse.StatusCode.ToString();
                    using (StreamReader srReader = new StreamReader(hwResponse.GetResponseStream(), Encoding.UTF8))
                    {
                        strResult = srReader.ReadToEnd();
                    }
                }
                hwRequest.Abort();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            return strResult;
        }

        private bool CheckArgument(string url)
        {
            return !url.Equals(_strempty);
        }
    }
}
