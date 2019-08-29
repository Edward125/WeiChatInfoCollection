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
using System.Threading;
using NPOI.XSSF.UserModel;
using NPOI.XSSF;
using NPOI.SS;
using NPOI.SS.UserModel;

namespace WeiChatInfoCollection
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        #region 參數

        
       public  static string StartTime = "1556467200000";
       public   static string EndTime = "1556553600000";
       public   static string Plant = "F721";
       public  static string WebLink = @"http://webap01.wks.wistron.com.cn:3010/api/DmcEvents?filter[where][STime][between][0]=" + @StartTime + @"&filter[where][STime][between][1]=" + @EndTime + @"&filter[where][toDMC]=1&filter[where][plant]=" + @Plant;

        Dictionary<string, string> ErrCode = new Dictionary<string, string>();

        List<string> eventIdList = new List<string>();
        List<string> eventNameList = new List<string>();

        //Thread 


        #endregion

        private void LoadUI()
        {
            this.Text = "Collect WeiChat Info." + "Ver.: " + Application.ProductVersion + ",Author:Edward Song";
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

            //DateTime ddt = ConvertIntDatetime("1566088633192");
            //ShowMessageInternal(MessageType.Begin,ddt.ToString ("yyyy-MM-dd HH:mm:ss"));
            //return;

            if (!CheckValue())
                return;

            this.btnStart.Cursor = Cursors.WaitCursor;
            string result = "";
            List<WebInfo> myWebInfoList ;
            try
            {
                ShowMessageInternal(MessageType.Begin, "Start loading info from web link...");
                result  = HttpGet(WebLink);
                ShowMessageInternal(MessageType.Success , "Load info frm web sucessful..");
            }
            catch (Exception e1)
            {

                ShowMessageInternal(MessageType.Failure, "Load info frm web error...");
                ShowMessageInternal(MessageType.Failure, e1.Message);
                this.btnStart.Cursor = Cursors.Default;
                return;
            }

            try
            {
                ShowMessageInternal(MessageType.Begin, "Analysis Json data...");
                JsonSerializerSettings jsetting = new JsonSerializerSettings();
                jsetting.NullValueHandling = NullValueHandling.Ignore;
                jsetting.Formatting = Formatting.None;
                jsetting.MissingMemberHandling = MissingMemberHandling.Ignore;
                 myWebInfoList = (List<WebInfo>)Newtonsoft.Json.JsonConvert.DeserializeObject(result, typeof(List<WebInfo>), jsetting);
                ShowMessageInternal(MessageType.Success, "Analysis Json data sucessful,total " + myWebInfoList.Count + " record(s)");
            }
            catch (Exception e2)
            {
                ShowMessageInternal(MessageType.Failure, "Analysis Json data fail...");
                ShowMessageInternal(MessageType.Failure, e2.Message);
                this.btnStart.Cursor = Cursors.Default;
                return;
            }


            try
            {
                ShowMessageInternal(MessageType.Begin, "Loading eventId & eventName...");
                string file = "EventId.txt";
                GetErrCode(file);
                ShowMessageInternal(MessageType.Success, "Total eventId:" + ErrCode.Keys.Count);
            }
            catch (Exception e3)
            {

                ShowMessageInternal(MessageType.Failure, "Loading eventId & eventName error...");
                ShowMessageInternal(MessageType.Failure, e3.Message);
                this.btnStart.Cursor = Cursors.Default;
                return;
            }

            ShowMessageInternal(MessageType.Begin, "Start to create excel...");
            IWorkbook wb = new XSSFWorkbook(new FileStream("sample.xlsx", FileMode.Open));
            try
            {
                //设定要使用的Sheet为第0个Sheet
                ISheet TempSheet = wb.GetSheetAt(0);
                int UsedRows = myWebInfoList.Count;
                for (int i = 1; i <= UsedRows; i++)
                {

                    DateTime dt = ConvertIntDatetime(myWebInfoList[i - 1].STime);
                    TempSheet.CreateRow(i).CreateCell(0).SetCellValue(dt.ToString("yyyy-MM-dd HH:mm:ss"));
                    //第一个Row要用Create的
                    TempSheet.GetRow(i).CreateCell(1).SetCellValue(myWebInfoList[i - 1].System);
                    //第二个Row之后直接用Get的
                    TempSheet.GetRow(i).CreateCell(2).SetCellValue(myWebInfoList[i - 1].plant);
                    string eventId = myWebInfoList[i - 1].eventId;
                    TempSheet.GetRow(i).CreateCell(3).SetCellValue(myWebInfoList[i - 1].eventId);

                    string eventName = "";
                    ErrCode.TryGetValue(eventId, out eventName);
                    TempSheet.GetRow(i).CreateCell(4).SetCellValue(eventName);
                    TempSheet.GetRow(i).CreateCell(5).SetCellValue(myWebInfoList[i - 1].eventType);
                    TempSheet.GetRow(i).CreateCell(6).SetCellValue(myWebInfoList[i - 1].alertType);
                    TempSheet.GetRow(i).CreateCell(7).SetCellValue(myWebInfoList[i - 1].ActionOK);
                    TempSheet.GetRow(i).CreateCell(8).SetCellValue(myWebInfoList[i - 1].alertItem);
                    TempSheet.GetRow(i).CreateCell(9).SetCellValue(myWebInfoList[i - 1].IssueType);
                    TempSheet.GetRow(i).CreateCell(10).SetCellValue(myWebInfoList[i - 1].syncId);

                    string it = myWebInfoList[i - 1].STime;
                    if (!string .IsNullOrEmpty (it)  && it != "null")
                    {
                        DateTime dtt = ConvertIntDatetime (it );
                        TempSheet.GetRow(i).CreateCell(11).SetCellValue(dtt.ToString ("yyyy-MM-dd"));
                    }
                    
                    //TempSheet.GetRow(i).CreateCell(11).SetCellValue(myWebInfoList[i - 1].STime);
                    it = myWebInfoList[i - 1].ETime.Trim();
                    if (!string .IsNullOrEmpty (it)  && it != "null")
                    {
                        DateTime dtt = ConvertIntDatetime (it );
                        TempSheet.GetRow(i).CreateCell(12).SetCellValue(dtt.ToString("yyyy-MM-dd HH:mm:ss"));
                    }
                    //TempSheet.GetRow(i).CreateCell(12).SetCellValue(myWebInfoList[i - 1].ETime);
                    it = myWebInfoList[i - 1].PTime.Trim();
                    if (!string .IsNullOrEmpty (it)  && it != "null")
                    {
                        DateTime dtt = ConvertIntDatetime (it );
                        TempSheet.GetRow(i).CreateCell(13).SetCellValue(dtt.ToString ("yyyy-MM-dd HH:mm:ss"));
                    }
                    //TempSheet.GetRow(i).CreateCell(13).SetCellValue(myWebInfoList[i - 1].PTime);
                    it = myWebInfoList[i - 1].STime.Trim();
                    if (!string .IsNullOrEmpty (it)  && it != "null")
                    {
                        DateTime dtt = ConvertIntDatetime (it );
                        TempSheet.GetRow(i).CreateCell(14).SetCellValue(dtt.ToString ("yyyy-MM-dd HH:mm:ss"));
                    }
                    //TempSheet.GetRow(i).CreateCell(14).SetCellValue(myWebInfoList[i - 1].EndingTime);


                    TempSheet.GetRow(i).CreateCell(15).SetCellValue(myWebInfoList[i - 1].L1Time);
                    TempSheet.GetRow(i).CreateCell(16).SetCellValue(myWebInfoList[i - 1].L2Time);
                    TempSheet.GetRow(i).CreateCell(17).SetCellValue(myWebInfoList[i - 1].L3Time);
                    TempSheet.GetRow(i).CreateCell(18).SetCellValue(myWebInfoList[i - 1].uId);
                    TempSheet.GetRow(i).CreateCell(19).SetCellValue(myWebInfoList[i - 1].status);
                    TempSheet.GetRow(i).CreateCell(20).SetCellValue(myWebInfoList[i - 1].level);
                    TempSheet.GetRow(i).CreateCell(21).SetCellValue(myWebInfoList[i - 1].shortMessage);
                    TempSheet.GetRow(i).CreateCell(22).SetCellValue(myWebInfoList[i - 1].eventTime);
                    TempSheet.GetRow(i).CreateCell(23).SetCellValue(myWebInfoList[i - 1].evtvalue1);
                    TempSheet.GetRow(i).CreateCell(24).SetCellValue(myWebInfoList[i - 1].evtvalue2);
                    TempSheet.GetRow(i).CreateCell(25).SetCellValue(myWebInfoList[i - 1].evtvalue3);
                    TempSheet.GetRow(i).CreateCell(26).SetCellValue(myWebInfoList[i - 1].evtvalue4);
                    TempSheet.GetRow(i).CreateCell(27).SetCellValue(myWebInfoList[i - 1].evtvalue5);
                    TempSheet.GetRow(i).CreateCell(28).SetCellValue(myWebInfoList[i - 1].evtvalue6);
                    TempSheet.GetRow(i).CreateCell(29).SetCellValue(myWebInfoList[i - 1].evtvalue7);
                    TempSheet.GetRow(i).CreateCell(30).SetCellValue(myWebInfoList[i - 1].evtvalue8);
                    TempSheet.GetRow(i).CreateCell(31).SetCellValue(myWebInfoList[i - 1].evtvalue9);
                    TempSheet.GetRow(i).CreateCell(32).SetCellValue(myWebInfoList[i - 1].evtvalue10);
                    TempSheet.GetRow(i).CreateCell(33).SetCellValue(myWebInfoList[i - 1].evtvalue11);
                    TempSheet.GetRow(i).CreateCell(34).SetCellValue(myWebInfoList[i - 1].evtvalue12);
                    TempSheet.GetRow(i).CreateCell(35).SetCellValue(myWebInfoList[i - 1].evtvalue13);
                    TempSheet.GetRow(i).CreateCell(36).SetCellValue(myWebInfoList[i - 1].evtvalue14);
                    TempSheet.GetRow(i).CreateCell(37).SetCellValue(myWebInfoList[i - 1].evtvalue15);
                    TempSheet.GetRow(i).CreateCell(38).SetCellValue(myWebInfoList[i - 1].pic);
                    TempSheet.GetRow(i).CreateCell(39).SetCellValue(myWebInfoList[i - 1].mPic);
                    TempSheet.GetRow(i).CreateCell(40).SetCellValue(myWebInfoList[i - 1].mPicPhone);
                    TempSheet.GetRow(i).CreateCell(41).SetCellValue(myWebInfoList[i - 1].userId);
                    TempSheet.GetRow(i).CreateCell(42).SetCellValue(myWebInfoList[i - 1].actionId);
                    TempSheet.GetRow(i).CreateCell(43).SetCellValue(myWebInfoList[i - 1].actionName );
                    TempSheet.GetRow(i).CreateCell(44).SetCellValue(myWebInfoList[i - 1].comment);
                    TempSheet.GetRow(i).CreateCell(45).SetCellValue(myWebInfoList[i - 1].extenDate);
                    TempSheet.GetRow(i).CreateCell(46).SetCellValue(myWebInfoList[i - 1].replyUserId);
                    TempSheet.GetRow(i).CreateCell(47).SetCellValue(myWebInfoList[i - 1].replyUserName);
                    TempSheet.GetRow(i).CreateCell(48).SetCellValue(myWebInfoList[i - 1].replyDate);
                    TempSheet.GetRow(i).CreateCell(49).SetCellValue(myWebInfoList[i - 1].toDMC);
                    TempSheet.GetRow(i).CreateCell(50).SetCellValue(myWebInfoList[i - 1].toNitify);

                }

                string newfile = "Summary_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
                if (File.Exists(newfile))
                    File.Delete(newfile);
                FileStream  fs = new FileStream("Summary_" + DateTime .Now.ToString  ("yyyyMMdd") +".xlsx", FileMode.Create);
                wb.Write(fs);
                fs.Close();
                fs.Dispose();

   
            }
            catch (Exception e4)
            {
                ShowMessageInternal(MessageType.Failure, "Create excel file error...");
                ShowMessageInternal(MessageType.Failure, e4.Message);
                this.btnStart.Cursor = Cursors.Default;
                return;
            }
            ShowMessageInternal(MessageType.Success, "Create " + "Summary_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx" +" file sucessful...");
            try
            {
                System.Diagnostics.Process.Start("explorer.exe", Application.StartupPath);
            }
            catch (Exception)
            {
                
            }
            
            this.btnStart.Cursor = Cursors.Default;


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
            UpdateWebLink();

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
        public DateTime ConvertIntDatetime(string utc)
        {
            System.DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1));
            double ltime = Convert.ToDouble(utc) / 1000;
            startTime = startTime.AddSeconds(ltime);
            //startTime = startTime.AddHours(8);//转化为北京时间(北京时间=UTC时间+8小时 )
            return startTime;
        }


    
        #endregion

        private void dtpStart_ValueChanged(object sender, EventArgs e)
        {
            txtStart.Text = Local2Utc(dtpStart.Value);
            StartTime = txtStart.Text.Trim();

            ShowMessageInternal(MessageType.Begin, "Start Local Time:" + dtpStart.Value.ToString("yyyy-MM-dd HH:mm:ss") + ", UTC Time:" + StartTime);
            UpdateWebLink();
        }

        private void dtpEnd_ValueChanged(object sender, EventArgs e)
        {
            txtEnd.Text = Local2Utc(dtpEnd.Value);
            EndTime = txtEnd.Text.Trim();
            ShowMessageInternal(MessageType.Begin, "End Local Time:" + dtpEnd.Value.ToString("yyyy-MM-dd HH:mm:ss") + ", UTC Time:" + EndTime);
            UpdateWebLink();
        }

        private void comboPlant_SelectedIndexChanged(object sender, EventArgs e)
        {
            Plant = comboPlant.Text;
            UpdateWebLink();

        }


        private void UpdateWebLink()
        {
            WebLink = @"http://webap01.wks.wistron.com.cn:3010/api/DmcEvents?filter[where][STime][between][0]=" + @StartTime + @"&filter[where][STime][between][1]=" + @EndTime + @"&filter[where][toDMC]=1&filter[where][plant]=" + @Plant;
            txtWebLink.Text = string.Empty;
            txtWebLink.Text = WebLink;
        }
        private void GetErrCode(string errcodefile)
        {

            ShowMessageInternal(MessageType.Begin, "Start to load eventId & eventMessage...");
            ErrCode = new Dictionary<string, string>();
            StreamReader sr = new StreamReader(errcodefile);
            while (!sr.EndOfStream)
            {

                string linestr = sr.ReadLine().Trim ();
                if (!string.IsNullOrEmpty (linestr ) && linestr.Contains (","))
                {

                    string eventId = linestr.Split(',')[0].Trim();
                    string eventMsg = linestr.Split(',')[1].Trim();
                    if (!ErrCode.Keys.Contains(eventId))
                        ErrCode.Add(eventId, eventMsg);
                }

            }
            sr.Close();
            ShowMessageInternal(MessageType.Success , "Load eventId & eventMessage sucessful...");
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
            public string STime { set; get; }
            public string ETime { set; get; }
            public string PTime { set; get; }
            public string EndingTime { set; get; }
            public string L1Time { set; get; }
            public string L2Time { set; get; }
            public string L3Time { set; get; }
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
            public string  evtvalue9 { set; get; }
            public string  evtvalue10 { set; get; }
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
