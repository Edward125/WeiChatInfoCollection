using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Renci.SshNet;

namespace WeiChatInfoCollection
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        #region 參數

        
        static string StartTime = "1556467200000";
        static string EndTime = "1556553600000";
        static string Plant = "F721";
        static string WebLink = @"http://webap01.wks.wistron.com.cn:3010/api/DmcEvents?filter[where][STime][between][0]=" + @StartTime + @"&filter[where][STime][between][1]=" + @EndTime + @"&filter[where][toDMC]=1&filter[where][plant]=" + @Plant;

        #endregion

        private void LoadUI()
        {
            txtWebLink.Text = WebLink;
            dtpStart.Value = Convert.ToDateTime(DateTime.Now.AddDays(-7).ToShortDateString() + " 00:00:00");
            dtpEnd.Value = Convert.ToDateTime(DateTime.Now.ToShortDateString() + " 00:00:00");
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            LoadUI();
        }


        #region other



        public   enum MessageType
        {
            Begin,
            Success,
            Failure,
        }



       public  void ShowMessageInternal(MessageType status, string message)
        {

            if (message == null)
                message = status.ToString();
            this.Invoke((EventHandler)(delegate
            {
                switch (status)
                {
                    case MessageType.Begin:
                        this.richMessage.SelectionColor = Color.Blue;
                        this.richMessage.AppendText(DateTime.Now.ToString("HH:mm:ss") + " " + message + "\n");
                        this.richMessage.Update();
                        break;
                    case MessageType.Success:
                        this.richMessage.SelectionColor = Color.Green;
                        this.richMessage.AppendText(DateTime.Now.ToString("HH:mm:ss") + " " + message + "\n");
                        this.richMessage.Update();
                        break;
                    case MessageType.Failure:
                        this.richMessage.SelectionColor = Color.Red;
                        this.richMessage.AppendText(DateTime.Now.ToString("HH:mm:ss") + " " + message + "\n");
                        this.richMessage.Update();
                        break;
                    default:
                        break;
                }
                if (richMessage.TextLength > richMessage.MaxLength - 1000)
                    richMessage.Clear();
                richMessage.SelectionStart = richMessage.TextLength;
                richMessage.ScrollToCaret();
            }));

        }



        #endregion

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (!CheckValue())
                return;



            ShowMessageInternal(MessageType.Begin, "Start loading info from web link...");
            string result = HttpGet(WebLink);

           // List<MyStok> myDeserializedObjList = (List<MyStok>)Newtonsoft.Json.JsonConvert.DeserializeObject(sc), typeof(List<MyStok>));


            List<WebInfo >  myWebInfoList = (List<WebInfo>) Newtonsoft.Json.JsonConvert.DeserializeObject(result,typeof(List<WebInfo>));
            WebInfo wi  = JsonConvert.DeserializeObject<WebInfo>(result);

            MessageBox.Show("OK");


        }



        private bool CheckValue()
        {
            txtStart.Text = Local2Utc(dtpStart.Value);
            StartTime = txtStart.Text.Trim();

            if (string.IsNullOrEmpty(StartTime))
            {
                ShowMessageInternal(MessageType.Failure, "Start time is null,pls retry...");
                dtpStart.Focus();
                return false;
            }


            txtEnd.Text = Local2Utc(dtpEnd.Value);
            EndTime = txtEnd.Text.Trim();
            if (string.IsNullOrEmpty(EndTime))
            {
                ShowMessageInternal(MessageType.Failure, "End time is null,pls retry...");
                dtpEnd.Focus();
                return false;
            }

              


            Plant = comboPlant.Text.Trim();
            if (string.IsNullOrEmpty(Plant ))
            {
                ShowMessageInternal(MessageType.Failure, "Plant is null,pls retry...");
                comboPlant.Focus();
                return false;
            }



            return true;
        }


        #region TransferTime



        private string  Local2Utc(DateTime dt)
        {
            return ((dt.ToUniversalTime().Ticks - 0x89f7ff5f7b58000L )/ 0x989680L).ToString().PadRight(13, '0');

        }

        public string Transfer2UTC(DateTime dt)
        {
            //本地时间(北京时间)
            //DateTime dt = Convert.ToDateTime("2019-01-01 00:00:00");

            //将北京时间转换成utc时间 （北京时间是utc时间+8小时，所以此时utc时间应该是 2016-06-11 15:59:59）
            DateTime utcNow = dt.ToUniversalTime();

            //将utc时间转换成秒 (即将1970-01-01 00:00:00 到 2016-06-11 15:59:59的时间转换成秒)
            double utc = ConvertDateTimeInt(utcNow);

            //将秒数转换成北京时间 (其实就是将utc时间转换成北京时间),所以又得到2016-06-11 23:59:59
            //DateTime dtime = ConvertIntDatetime(utc);

            //return View();
            return utc.ToString();
            //return utc.ToString().PadRight(13, '0');
        }

        /// <summary>
        /// 将时间转换成秒(这个秒是指1970-1-1 00:00:00 到你指定的时间之间的秒数)
        /// </summary>
        /// <param name="time">指定时间</param>
        /// <returns>秒数</returns>
        public double ConvertDateTimeInt(System.DateTime time)
        {
            double intResult = 0;
            System.DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1));
            intResult = (time - startTime).TotalSeconds;
            return intResult;
        }

        /// <summary>
        /// 将秒数转换成北京时间
        /// </summary>
        /// <param name="utc">秒数</param>
        /// <returns>北京时间</returns>
        public DateTime ConvertIntDatetime(double utc)
        {
            System.DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1));
            startTime = startTime.AddSeconds(utc);
            startTime = startTime.AddHours(8);//转化为北京时间(北京时间=UTC时间+8小时 )
            return startTime;
        }

        
    
 
 






        #endregion

        private void dtpStart_ValueChanged(object sender, EventArgs e)
        {
            txtStart.Text = Local2Utc(dtpStart.Value);
            StartTime = txtStart.Text.Trim();
            ShowMessageInternal(MessageType.Begin, "Start Local Time:" + dtpStart.Value.ToString("yyyy-MM-dd HH:mm:ss") + ", UTC Time:" + StartTime);
        }

        private void dtpEnd_ValueChanged(object sender, EventArgs e)
        {
            txtEnd.Text = Local2Utc(dtpEnd.Value);
            EndTime = txtEnd.Text.Trim();
            ShowMessageInternal(MessageType.Begin, "End Local Time:" + dtpEnd.Value.ToString("yyyy-MM-dd HH:mm:ss") + ", UTC Time:" + EndTime);
        }

        private void comboPlant_SelectedIndexChanged(object sender, EventArgs e)
        {
            Plant = comboPlant.Text;
        }

        #region JsonInfo


        public class WebInfo
        {
            //
            public string System { set; get; }
            public string plant { set; get; }
            public string eventId { set; get; }
            public int eventType { set; get; }
            public int alertType { set; get; }
            public int ActionOK { set; get; }
            public int alertItem { set; get; }
            public int IssueType { set; get; }
            public string syncId { set; get; }
            public UInt64 STime { set; get; }
            public UInt64 ETime { set; get; }
            public UInt64 PTime { set; get; }
            public UInt64 EndingTime { set; get; }
            public UInt64 L1Time { set; get; }
            public UInt64 L2Time { set; get; }
            public UInt64 L3Time { set; get; }
            public string uId { set; get; }
            public int status { set; get; }
            public string level { set; get; }
            public string shortMessage { set; get; }
            public string eventTime { set; get; }
            public string evtvalue1 { set; get; }
            public string evtvalue2 { set; get; }
            public string evtvalue3 { set; get; }
            public string evtvalue4 { set; get; }
            public string evtvalue5 { set; get; }
            public string evtvalue6 { set; get; }
            public string evtvalue7 { set; get; }
            public string evtvalue8 { set; get; }
            public string evtvalue9 { set; get; }
            public string evtvalue10 { set; get; }
            public string evtvalue11 { set; get; }
            public string evtvalue12 { set; get; }
            public string evtvalue13 { set; get; }
            public string evtvalue14 { set; get; }
            public string evtvalue15 { set; get; }
            public string pic { set; get; }
            public string mPic { set; get; }
            public string mPicPhone { set; get; }
            public string userId { set; get; }
            public string actionId { set; get; }
            public string actionName { set; get; }
            public string comment { set; get; }
            public string extenDate { set; get; }
            public string replyUserId { set; get; }
            public string replyUserName { set; get; }
            public string replyDate { set; get; }
            public int toDMC { set; get; }
            public int toNitify { set; get; }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>

        public static string HttpGet(string url)
        {
            StreamReader reader = null;
            try
            {
                //ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);
                Encoding encoding = Encoding.UTF8;
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "GET";
                request.Accept = "text/html, application/xhtml+xml, */*";
                request.ContentType = "application/json";

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                using (reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                
               
                return reader.ReadToEnd();
            }

        }
        

        #endregion
    }
}
