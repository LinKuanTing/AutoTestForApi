using System;
using System.Collections;
using System.Configuration;
using System.IO;
using System.Text.RegularExpressions;

namespace NewmanCMD
{
    class Program
    {
        static string strMyDate, strCmdText, strCollectionPath, strEnvironmentPath, strResultDirPath, strResultPath, strFileName;
        static Boolean isPassForTest = false;


        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern bool FreeConsole();

        static void Main(string[] args)
        {
            //將視窗隱藏起來
            FreeConsole();
            
            Console.WriteLine("********** Setting file path... **********");
            try
            {
                //設定檔案的位置等        
                DateTime dateTime = DateTime.Now;
                strMyDate = dateTime.ToString("yyyyMMddHHmmss");
                FileModel.SetDate(strMyDate);
                strFileName = "result_" + strMyDate + ".html";            
                strCollectionPath = ConfigurationManager.AppSettings["collectionDirPath"] + ConfigurationManager.AppSettings["collectionFileName"];
                strEnvironmentPath = ConfigurationManager.AppSettings["environmentDirPath"] + ConfigurationManager.AppSettings["environmentFileName"];
                strResultDirPath = ConfigurationManager.AppSettings["resultDirPath"];
                strResultPath = strResultDirPath + strFileName;
            }
            catch (Exception e)
            {
                FileModel.WriteLog(e.Message);
                FileModel.WriteLog(e.StackTrace);
            }
            Console.WriteLine("********** Finish **********");



            Console.WriteLine("********** Check file path is exist... **********");
            try
            {
                //檢查路徑檔案是否存在
                if (!File.Exists(strCollectionPath) || !File.Exists(strEnvironmentPath))
                {
                    //此兩檔案是測試必要條件 因此缺少其一無法測試則離開
                    if (!File.Exists(strCollectionPath))
                    {
                        FileModel.WriteLog("No find path with " + strCollectionPath);
                    }
                    if (!File.Exists(strEnvironmentPath))
                    {
                        FileModel.WriteLog("No find path with " + strEnvironmentPath);
                    }
                    return;
                }
                if (!Directory.Exists(strResultDirPath))
                {
                    //用來存放報告日誌的目錄
                    Directory.CreateDirectory(strResultDirPath);
                }
            }
            catch(Exception e)
            {
                FileModel.WriteLog(e.Message);
                FileModel.WriteLog(e.StackTrace);
            }
            Console.WriteLine("********** Finish **********");



            Console.WriteLine("********** Execute newman command to get test report... **********");
            try
            {
                //command 指令: 產生測試報告
                strCmdText = "newman run " + strCollectionPath + " -e " + strEnvironmentPath + " -r html --reporter-html-export " + strResultPath;
                CommandLineModel.excuteCMD(strCmdText);
            }
            catch (Exception e)
            {
                FileModel.WriteLog(e.Message);
                FileModel.WriteLog(e.StackTrace);
            }
            Console.WriteLine("********** Finish **********");



            Console.WriteLine("********** Analyze report content... **********");
            //處理報告內容
            ArrayList alMainContent = new ArrayList();
            ArrayList alDetailContent = new ArrayList();
            ArrayList alFailDetailContent = new ArrayList();
            try
            {
                string strReadFile = File.ReadAllText(strResultPath);             
                int start = strReadFile.IndexOf("<body>");
                int end = strReadFile.IndexOf("</body>");
                String strHtmlBody = strReadFile.Substring(start, end - start + 7);
                //Console.WriteLine(Regex.Replace(strHtmlBody, @"<[^>]*>", String.Empty));

                String[] strCheckPoint;
                if (strHtmlBody.IndexOf("<br/><h4>Failures</h4>")>0)
                {
                    strCheckPoint= new String[4] { "<h3>Newman Report</h3>", "<br/><h4>Requests</h4>", "<br/><h4>Failures</h4>", "</body>" };
                }
                else
                {
                    strCheckPoint = new String[3] { "<h3>Newman Report</h3>", "<br/><h4>Requests</h4>", "</body>" };
                }

                ArrayList alHtmlContent = new ArrayList();
                //內容取分段
                for (int i=1; i<strCheckPoint.Length; i++)
                {
                    int indexStart = strHtmlBody.IndexOf(strCheckPoint[i - 1]);
                    int indexEnd = strHtmlBody.IndexOf(strCheckPoint[i]);                 
                    String s = strHtmlBody.Substring(indexStart,indexEnd-indexStart);
                    alHtmlContent.Add(s);
                }


               
                for (int i=0; i<alHtmlContent.Count; i++)
                {
                    string str;
                    string[] strArray, strPick;
                    switch (i)
                    {
                        case 0:                            
                            str = alHtmlContent[i].ToString();
                            str = Regex.Replace(str, @"<[^>]*>", " ").Replace("&nbsp;", "");
                            strArray = str.Split("\n");

                            strPick = new string[] {"Collection","Time", "Exported with", "Test Scripts", "Total run duration", "Total Failures" };
                            foreach (string s in strArray)
                            {
                                foreach (string ss in strPick)
                                {
                                    if (s.Contains(ss))
                                    {                                      
                                        
                                        if (ss.Equals("Time"))
                                        {
                                            string datetime = s.Replace("Time", "").Trim();
                                            string[] strAry = datetime.Split(" ");

                                            DateTime dateObject = DateTime.Parse(strAry[1] + " " + strAry[2] + " " + strAry[3]);

                                            alMainContent.Add("Date " + dateObject.ToString("yyyy-MM-dd"));
                                            alMainContent.Add("Time " + strAry[4]);

                                            continue;
                                        }

                                        if (ss.Equals("Total Failures"))
                                        {                                                                                  
                                            isPassForTest = (Convert.ToInt32(s.Replace("Total Failures", "").Trim()) == 0);
                                        }

                                        alMainContent.Add(s.Trim());

                                    }
                                }
                            }

                            break;
                        case 1:
                            str = alHtmlContent[i].ToString();
                            strArray = str.Split("\n");

                            strPick = new string[] { "panel-title", "Method", "URL", "Mean time per request", "Total passed tests", "Total failed tests", "<tbody>" };
                            foreach (string s in strArray)
                            {
                                foreach (string ss in strPick)
                                {
                                    if (s.Contains(ss))
                                    {
                                        if (ss.Equals("panel-title"))
                                        {
                                            alDetailContent.Add("Api Name "+ Regex.Replace(s, @"<[^>]*>", " ").Trim());
                                        }else if (ss.Equals("<tbody>"))
                                        {
                                            alDetailContent.Add("Test Case " + Regex.Replace(s, @"<[^>]*>", " ").Trim());
                                        }
                                        else
                                        {
                                            alDetailContent.Add(Regex.Replace(s, @"<[^>]*>", " ").Trim());
                                        }
                                        
                                        
                                    }
                                }
                            }

                            break;
                        case 2:
                            str = alHtmlContent[i].ToString();
                            str = str.Replace("&#x27;", "'");
                            strArray = str.Split("\n");

                            strPick = new string[] {"Description" , "#collapse-request" };
                            foreach (string s in strArray)
                            {
                                foreach (string ss in strPick)
                                {
                                    if (!s.Trim().Equals("") && !s.Contains("<"))
                                    {
                                        alFailDetailContent.Add(Regex.Replace(s, @"<[^>]*>", " ").Trim());
                                        break;
                                    }
                                    else if (s.Contains(ss))
                                    {                                        
                                        alFailDetailContent.Add(Regex.Replace(s, @"<[^>]*>", " ").Trim());
                                    }
                                }
                            }

                            break;
                    }
                }
            }
            catch (Exception e)
            {
                FileModel.WriteLog(e.Message);
                FileModel.WriteLog(e.StackTrace);
            }
            Console.WriteLine("********** Finish **********");



            Console.WriteLine("********** Upload test data to server... **********");
            try
            {
                //資料儲存至Server
                SqlModel sqlModel = new SqlModel();
                sqlModel.ConnectToServer();
                sqlModel.MainDataUpload(alMainContent);
                sqlModel.DetailDataUpload(alDetailContent,alFailDetailContent);
            }
            catch (Exception e)
            {
                FileModel.WriteLog(e.Message);
                FileModel.WriteLog(e.StackTrace);
            }
            Console.WriteLine("********** Finish **********");



            Console.WriteLine("********** Send report mail... **********");
            try
            {
                //寄送報告至收件者(有錯誤才寄送報告)
                if (!isPassForTest)
                {
                    string strAllReceivers = ConfigurationManager.AppSettings["receivers"];

                    string[] receivers = strAllReceivers.Split(",");

                    for (int i = 0; i < receivers.Length; i++)
                    {
                        String strReceiver = receivers[i];
                        MailModel.SendMail(strReceiver, strResultPath);
                    }
                }
 
            }
            catch (Exception e)
            {
                FileModel.WriteLog(e.Message);
                FileModel.WriteLog(e.StackTrace);
            }
            Console.WriteLine("********** Finish **********");

    


            Console.WriteLine("********** Delete exist over 90 days file... **********");
            try
            {
                //刪除超過90天資料
                FileModel.DeleteOldFiles(ConfigurationManager.AppSettings["resultDirPath"], 90);
            }
            catch (Exception e)
            {
                FileModel.WriteLog(e.Message);
                FileModel.WriteLog(e.StackTrace);
            }
            Console.WriteLine("********** Finish **********");



            Console.WriteLine("********** Check Requirement of Create weekly report ... **********");
            try
            {
                DateTime nowTime = DateTime.Now;
                DateTime lastTime = DateTime.Parse(ConfigurationManager.AppSettings["lastCreateReportDateTime"]);

                if (nowTime.DayOfWeek == DayOfWeek.Monday && (nowTime-lastTime).Days >= 7)
                {
                    SqlModel sqlModel = new SqlModel();
                    sqlModel.ConnectToServer();
                    FileModel.CreateWeeklyReport(sqlModel.GetWeeklyReportContent());
                    sqlModel.DisconnectToServer();

                    Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                    config.AppSettings.Settings["lastCreateReportDateTime"].Value = nowTime.ToString("yyyy/MM/dd");

                }
            }
            catch (Exception e)
            {
                FileModel.WriteLog(e.Message);
                FileModel.WriteLog(e.StackTrace);
            }
            Console.WriteLine("********** Finish **********");

           // Console.ReadKey();


        }






    }
}
