using System;
using System.Collections;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Text;

namespace NewmanCMD
{
    class SqlModel
    {
        private string strServerName, strDataBaseName, strUserName, strPassword;
        string apiName;
        int mainID, detailID;
        SqlConnection sqlConnection;
        
        public void ConnectToServer()
        {
            strServerName = ConfigurationManager.AppSettings["dataSource"];
            strDataBaseName = ConfigurationManager.AppSettings["initialCatalog"];
            strUserName = ConfigurationManager.AppSettings["sqlUserId"];
            strPassword = ConfigurationManager.AppSettings["sqlPassword"];
            //使用SqlConnection建立資料庫連線
            sqlConnection = new SqlConnection("data source = "+ strServerName + "; " + "initial catalog = "+ strDataBaseName +
                "; user id = "+ strUserName + "; password = "+ strPassword);

            //啟用連線
            sqlConnection.Open();          

        }

        public void DisconnectToServer()
        {
            sqlConnection.Close();
            sqlConnection.Dispose();
        }


        public void MainDataUpload(ArrayList mainData)
        {
            //將SQL查詢指令寫在這  ex.(select * from table) 
            string strCmdText = "INSERT INTO NM_MainResult([CollectionName],[Date],[Time],[ExportedWith],[TotalRequests],[TotalRunDuration],[TotalFailures]) VALUES(@value1, @value2, @value3, @value4, @value5, @value6, @value7);select @@IDENTITY";

            SqlCommand myCmd = new SqlCommand(strCmdText, sqlConnection);

            //處理main result report儲存內容
            foreach (string s in mainData)
            {
                if (s.Contains("Collection"))
                {
                    myCmd.Parameters.AddWithValue("@value1", s.Replace("Collection","").Trim());
                }else if (s.Contains("Date"))
                {
                    myCmd.Parameters.AddWithValue("@value2", s.Replace("Date","").Trim());
                }else if (s.Contains("Time"))
                {
                    myCmd.Parameters.AddWithValue("@value3", s.Replace("Time","").Trim());
                }else if (s.Contains("Exported"))
                {
                    myCmd.Parameters.AddWithValue("@value4", s.Replace("Exported with", "").Trim());
                }else if (s.Contains("Test Scripts"))
                {
                    string[] str = s.Replace("Test Scripts", "").Trim().Split(" ");
                    myCmd.Parameters.AddWithValue("@value5", str[0]);                    
                }else if (s.Contains("Total run duration"))
                {
                    string str = s.Replace("Total run duration", "").Trim();
                    if (str.IndexOf(" ") != -1)
                    {
                        string[] strAry = str.Split(" ");
                        str = (float.Parse(strAry[0].Replace("m", "")) * 60 + float.Parse(strAry[1].Replace("s", ""))).ToString();
                    }
                    else
                    {
                        if (str.Contains("ms"))
                        {
                            str = (float.Parse(str.Replace("ms", "")) / 1000).ToString();
                        }
                        else
                        {
                            str = str.Replace("s", "");
                        }
                    }
                    myCmd.Parameters.AddWithValue("@value6", str);
                }else if (s.Contains("Total Failures"))
                {
                    myCmd.Parameters.AddWithValue("@value7", s.Replace("Total Failures", "").Trim());
                }
            }
            try
            {
                //執行SQL命令                
                mainID = Convert.ToInt32(myCmd.ExecuteScalar()) ;
               
            }
            catch(SqlException ex)
            {
                StringBuilder errorMessages = new StringBuilder();
                for (int i = 0; i < ex.Errors.Count; i++)
                {
                    errorMessages.Append("MainDataUpload failed \n" +
                        "Index #" + i + "\n" +
                        "Message: " + ex.Errors[i].Message + "\n" +
                        "LineNumber: " + ex.Errors[i].LineNumber + "\n" +
                        "Source: " + ex.Errors[i].Source + "\n" +
                        "Procedure: " + ex.Errors[i].Procedure + "\n");
                }
                //Console.WriteLine(errorMessages.ToString());             
                //Console.WriteLine(ex.StackTrace);
                FileModel.WriteLog(ex.Message);
                FileModel.WriteLog(ex.StackTrace);
                FileModel.WriteLog("Error: \n"+errorMessages.ToString());
            }
                
        }


        public void DetailDataUpload(ArrayList detailData , ArrayList failData)
        {
            //將SQL查詢指令寫在這  ex.(select * from table) 
            string strCmdText = "INSERT INTO NM_Detail([MainResultID],[ApiName],[Method],[URL],[RunDuration],[TestScript],[PassCount],[FailCount],[CDT]) VALUES(@value1, @value2, @value3, @value4, @value5, @value6, @value7, @value8, @value9);select @@IDENTITY";

            SqlCommand myCmd = new SqlCommand(strCmdText, sqlConnection);

            try
            {
                //處理main result report儲存內容         
                int c = 0;
                Boolean isExistFailure = false;
                foreach (string s in detailData)
                {             

                    if (s.Contains("Api Name"))
                    {
                        apiName = s.Replace("Api Name", "").Trim();
                        myCmd.Parameters.AddWithValue("@value2", apiName);
                    }
                    else if (s.Contains("Method"))
                    {   
                        myCmd.Parameters.AddWithValue("@value3", s.Replace("Method", "").Trim());                       
                    }
                    else if (s.Contains("URL"))
                    {
                        myCmd.Parameters.AddWithValue("@value4", s.Replace("URL", "").Trim());
                    }
                    else if (s.Contains("Mean time per request"))
                    {
                        string str = s.Replace("Mean time per request", "").Trim();

                        if (str.IndexOf(" ") != -1)
                        {
                            string[] strAry = str.Split(" ");
                            str = (float.Parse(strAry[0].Replace("m", ""))*60 + float.Parse(strAry[1].Replace("s",""))).ToString();
                        }
                        else
                        {
                            if (str.Contains("ms"))
                            {
                                str = (float.Parse(str.Replace("ms", "")) / 1000).ToString();
                            }
                            else
                            {
                                str = str.Replace("s", "");
                            }
                        }
                        
                        myCmd.Parameters.AddWithValue("@value5", str);
                    }
                    else if (s.Contains("Test Case"))
                    {
                        string[] str = s.Replace("Test Case", "").Trim().Split("  ");
                    
                        myCmd.Parameters.AddWithValue("@value6", (int)str.Length/3);
                    }
                    else if (s.Contains("Total passed tests"))
                    {
                        myCmd.Parameters.AddWithValue("@value7", s.Replace("Total passed tests", "").Trim());
                    }
                    else if (s.Contains("Total failed tests"))
                    {
                        String str = s.Replace("Total failed tests", "").Trim();
                        if (int.Parse(str) >= 1)
                        {
                            isExistFailure = true;
                        }
                        myCmd.Parameters.AddWithValue("@value8", str);
                        
                    }

                    c++;
                    if (c == 7)
                    {
                        myCmd.Parameters.AddWithValue("@value1", mainID);
                        myCmd.Parameters.AddWithValue("@value9", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                        
                        //執行SQL命令
                        if (isExistFailure)
                        {
                           
                            detailID = Convert.ToInt32(myCmd.ExecuteScalar());
                            FailDataUpload(failData);
                        }
                        else
                        {
                            myCmd.ExecuteNonQuery();
                        }
                        isExistFailure = false;
                        c = 0;
                        myCmd.Parameters.Clear();
                    }
                }
                
            }
            catch (SqlException ex)
            {
                StringBuilder errorMessages = new StringBuilder();
                for (int i = 0; i < ex.Errors.Count; i++)
                {
                    errorMessages.Append("DetailDataUpload failed \n" + 
                        "Index #" + i + "\n" +
                        "Message: " + ex.Errors[i].Message + "\n" +
                        "LineNumber: " + ex.Errors[i].LineNumber + "\n" +
                        "Source: " + ex.Errors[i].Source + "\n" +
                        "Procedure: " + ex.Errors[i].Procedure + "\n");
                }
                //Console.WriteLine(errorMessages.ToString());
                //Console.WriteLine(ex.StackTrace);
                FileModel.WriteLog(ex.Message);
                FileModel.WriteLog(ex.StackTrace);
                FileModel.WriteLog("Error: \n" + errorMessages.ToString());
            }

        }


        public void FailDataUpload(ArrayList failData)
        {

            //將SQL查詢指令寫在這  ex.(select * from table) 
            string strCmdText = "INSERT INTO NM_FailureDetail([DetailID],[FailureName],[Description],[CDT]) VALUES(@valueA, @valueB, @valueC, @valueD)";

            SqlCommand myCmd = new SqlCommand(strCmdText, sqlConnection);

            try
            {
                //格式修正儲存格式 (DetailID、ErrorName、Description)
                ArrayList newFailData = new ArrayList();
                for (int i = 0; i < failData.Count; i++)
                {   
                    String s1 = failData[i].ToString();

                    newFailData.Add(s1);

                    if (s1.Contains("Error"))
                    {                      
                       if (!failData[i+1].ToString().Contains("Description"))
                        {
                            newFailData.Add("Description");
                        }
                    }
                }



                //處理Fail detail report儲存內容 
                string getApiName = "";
                int c = 0;              
                for (int i=0; i<newFailData.Count; i++)
                {
                    string s = newFailData[i].ToString();

                    if (s.Contains("Error"))
                    {
                        myCmd.Parameters.AddWithValue("@valueB", s.Trim());
                    }
                    else if (s.Contains("Description"))
                    {                      
                        myCmd.Parameters.AddWithValue("@valueC", s.Replace("Description", "").Trim());
                    }
                    else
                    {
                        getApiName = s.Trim();
                    }

                    c++;
                    if (c%3 == 0)
                    {
                        if (getApiName.Equals(apiName))
                        {
                            myCmd.Parameters.AddWithValue("@valueA", detailID);
                            myCmd.Parameters.AddWithValue("@valueD", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));

                            //執行SQL命令
                            myCmd.ExecuteNonQuery();
                        }                       
                         //清空資料
                         c = 0;
                         myCmd.Parameters.Clear();
                    }     
                }
               
            }
            catch (SqlException ex)
            {
                StringBuilder errorMessages = new StringBuilder();
                for (int i = 0; i < ex.Errors.Count; i++)
                {
                    errorMessages.Append("FailDataUpload failed \n"+
                        "Index #" + i + "\n" +
                        "Message: " + ex.Errors[i].Message + "\n" +
                        "LineNumber: " + ex.Errors[i].LineNumber + "\n" +
                        "Source: " + ex.Errors[i].Source + "\n" +
                        "Procedure: " + ex.Errors[i].Procedure + "\n");
                }
                //Console.WriteLine(errorMessages.ToString());
                //Console.WriteLine(ex.StackTrace);
                FileModel.WriteLog(ex.Message);
                FileModel.WriteLog(ex.StackTrace);
                FileModel.WriteLog("Error: \n" + errorMessages.ToString());
            }

        }




      

        public String GetWeeklyReportContent()
        {
            string strSqlCommand = "";
            string strReportContent = "";
            SqlDataReader reader;

            /*處理內容*/
            //標題與創建時間
            strReportContent += "WGH API 測試統計週報" + ","  + "\t CDt:" + DateTime.Now.ToString("yyyy/MM/dd") + "\n";
            //集合API測試名稱
            strReportContent += "Collection 名稱" + "," + "WGH_APIs_Test" + "\n";
            //統計日期期間
            strReportContent += "日期" + "," + DateTime.Now.AddDays(-7).ToString("yyyy/MM/dd")+" ~ "+ DateTime.Now.AddDays(-1).ToString("yyyy/MM/dd") + "\n";
            //計算collection執行次數
            int collectionExecTime = 0;
            strSqlCommand = @"select Count(*) from NM_MainResult where Date >= DATEADD(DAY, -8, GETDATE()) AND Date <= DATEADD(DAY, -1, GETDATE())";
            reader = new SqlCommand(strSqlCommand, sqlConnection).ExecuteReader();
            while (reader.Read()) { collectionExecTime = reader.GetInt32(0); }
            reader.Close();
            strReportContent += "Collection 執行次數" + "," + collectionExecTime + "\n";


            //個別API執行測試腳本數量
            ArrayList alApiTestScripts = new ArrayList();
            strSqlCommand = @"select sum(TestScript), ApiName from NM_Detail inner join NM_MainResult on NM_Detail.MainResultID=NM_MainResult.ID where Date >= DATEADD(DAY, -8, GETDATE()) AND Date <= DATEADD(DAY, -1, GETDATE()) group by ApiName ORDER BY apiName ASC";
            reader = new SqlCommand(strSqlCommand,sqlConnection).ExecuteReader();
            while (reader.Read())
            {
                alApiTestScripts.Add(reader.GetInt32(0) + "," + reader.GetString(1));  //格式為 腳本數,API名稱
            }
            reader.Close();
            //統計總共腳本請求次數
            int totalScriptCount = 0;
            for (int i=0; i<alApiTestScripts.Count; i++)
            {
                string[] str = alApiTestScripts[i].ToString().Split(",");
                totalScriptCount += Convert.ToInt32(str[0]);

            }
            strReportContent += "總共請求次數" + "," + totalScriptCount + "\n";


            //個別API執行測試腳本失敗數量
            ArrayList alApiTestFailures = new ArrayList();
            strSqlCommand = @"SELECT SUM(FailCount), ApiName FROM NM_Detail INNER JOIN NM_MainResult ON NM_Detail.MainResultID=NM_MainResult.ID WHERE Date >= DATEADD(DAY, -8, GETDATE()) AND Date <= DATEADD(DAY, -1, GETDATE()) GROUP BY ApiName ORDER BY apiName ASC";
            reader = new SqlCommand(strSqlCommand, sqlConnection).ExecuteReader();
            while (reader.Read())
            {
                alApiTestFailures.Add(reader.GetInt32(0) + "," + reader.GetString(1));//格式為 腳本數,API名稱
            }
            reader.Close();
            //計算成功通過率
            int totalFailureCount = 0;
            for (int i=0; i<alApiTestFailures.Count; i++)
            {
                string[] str = alApiTestFailures[i].ToString().Split(",");
                totalFailureCount += Convert.ToInt32(str[0]);
            }
            strReportContent += "成功率(%)" + "," + String.Format("{0:0.0%}", (float) (totalScriptCount-totalFailureCount)/totalScriptCount).Replace("%","") + "\n";

            //執行collection測試的平均、最短、最長花費時間
            strSqlCommand = @"select AVG(TotalRunDuration), MAX(TotalRunDuration), MIN(TotalRunDuration) from NM_MainResult where Date >= DATEADD(DAY, -8, GETDATE()) AND Date <= DATEADD(DAY, -1, GETDATE())";
            reader = new SqlCommand(strSqlCommand, sqlConnection).ExecuteReader();
            while (reader.Read())
            {
                strReportContent += "平均花費時間(秒)" + "," + reader.GetDouble(0).ToString("f1") + "\n";
                strReportContent += "最長花費時間(秒)" + "," + reader.GetDouble(1) + "\n";
                strReportContent += "最短花費時間(秒)" + "," + reader.GetDouble(2) + "\n";
            }
            reader.Close();

            //紀錄Error發生的類型與次數
            ArrayList alErrorData = new ArrayList();
            strSqlCommand = @"select FailureName from NM_FailureDetail inner join NM_Detail on NM_FailureDetail.DetailID=NM_Detail.ID inner join NM_MainResult on NM_Detail.MainResultID=NM_MainResult.ID where Date >= DATEADD(DAY, -8, GETDATE()) AND Date <= DATEADD(DAY, -1, GETDATE())";
            reader = new SqlCommand(strSqlCommand, sqlConnection).ExecuteReader();
            while (reader.Read())
            {
                if (alErrorData.Count == 0)
                {
                    alErrorData.Add(reader.GetString(0).Split(":")[0] + "," + "1"); //格式為 錯誤類型 次數
                }
                else
                {
                    string getType = reader.GetString(0).Split(":")[0];
                    for (int i = 0; i < alErrorData.Count; ++i)
                    {
                        string nowType = alErrorData[i].ToString().Split(",")[0];
                        int errorTime = Convert.ToInt32(alErrorData[i].ToString().Split(",")[1]);
                        if (getType.Equals(nowType))
                        {                           
                            alErrorData[i] = nowType + "," + ++errorTime ;
                            break;
                        }
                        if (i == alErrorData.Count - 1)
                        {
                            alErrorData.Add(reader.GetString(0).Split(":")[0] + "," + "1");
                            break;
                        }
                    }         
                }
            }
            reader.Close();        
            for (int i = 0; i < alErrorData.Count; ++i)
            {
                string[] str = alErrorData[i].ToString().Split(",");
                strReportContent += str[0] + "出現次數" + "," + str[1] + "\n";
            }


            strReportContent += '\n';
            //紀錄個別API測試結果
            if (alApiTestScripts.Count == alApiTestFailures.Count)
            {
                //紀錄API名稱  按字母順序作排序
                strReportContent += "Api 名稱" + ",";
                for (int i = 0; i < alApiTestScripts.Count; ++i)
                {
                    string[] strApiData1 = alApiTestScripts[i].ToString().Split(",");
                    string[] strApiData2 = alApiTestFailures[i].ToString().Split(",");
                    if (strApiData1[1].Equals(strApiData2[1]))
                    {
                        if (i == alApiTestScripts.Count - 1)
                        {
                            strReportContent += strApiData1[1] + "\n";
                        }
                        else
                        {
                            strReportContent += strApiData2[1] + ",";
                        }
                    }
                }

                //紀錄API總共測試腳本數
                strReportContent += "總共測試次數" + ",";
                for (int i = 0; i < alApiTestScripts.Count; ++i)
                {
                    string[] strApiData = alApiTestScripts[i].ToString().Split(",");
                    if (i == alApiTestScripts.Count - 1)
                    {
                        strReportContent += strApiData[0] + "\n";
                    }
                    else
                    {
                        strReportContent += strApiData[0] + ",";
                    }
                }

                //紀錄API測試失敗次數
                strReportContent += "失敗次數" + ",";
                for (int i = 0; i < alApiTestFailures.Count; ++i)
                {
                    string[] strApiData = alApiTestFailures[i].ToString().Split(",");
                    if (i == alApiTestFailures.Count - 1)
                    {
                        strReportContent += strApiData[0] + "\n";
                    }
                    else
                    {
                        strReportContent += strApiData[0] + ",";
                    }
                }

                //紀錄API成功率
                strReportContent += "成功率(%)" + ",";
                for (int i = 0; i < alApiTestScripts.Count; ++i)
                {
                    int scriptCount = Convert.ToInt32(alApiTestScripts[i].ToString().Split(",")[0]);
                    int failureCount = Convert.ToInt32(alApiTestFailures[i].ToString().Split(",")[0]);

                    if (i == alApiTestScripts.Count - 1)
                    {
                        strReportContent += String.Format("{0:0.0%}", (float)(scriptCount - failureCount) / scriptCount).Replace("%", "").Replace("100.0", "100") + "\n";
                    }
                    else
                    {
                        strReportContent += String.Format("{0:0.0%}", (float)(scriptCount - failureCount) / scriptCount).Replace("%", "").Replace("100.0", "100") + ",";
                    }
                }

                //取得個別Api的平均、最大、最小花費時間
                ArrayList alApiRunDuration = new ArrayList();
                strSqlCommand = @"SELECT ApiName, AVG(RunDuration), MAX(RunDuration), MIN(RunDuration) FROM NM_Detail INNER JOIN NM_MainResult ON NM_Detail.MainResultID=NM_MainResult.ID WHERE Date >= DATEADD(DAY, -8, GETDATE()) AND Date <= DATEADD(DAY, -1, GETDATE()) GROUP BY ApiName ORDER BY apiName ASC";
                reader = new SqlCommand(strSqlCommand, sqlConnection).ExecuteReader();
                while (reader.Read())
                {
                    alApiRunDuration.Add(reader.GetString(0) + "," + reader.GetDouble(1) + "," + reader.GetDouble(2) + "," + reader.GetDouble(3));
                }
                reader.Close();
                //紀錄平均時間
                strReportContent += "平均時間" + ",";
                for (int i = 0; i < alApiRunDuration.Count; i++)
                {
                    string[] str = alApiRunDuration[i].ToString().Split(",");
                    if (i == alApiRunDuration.Count - 1)
                    {
                        strReportContent += str[1] + "\n";
                    }
                    else
                    {
                        strReportContent += str[1] + ",";
                    }
                }
                //紀錄最長時間
                strReportContent += "最長時間" + ",";
                for (int i = 0; i < alApiRunDuration.Count; i++)
                {
                    string[] str = alApiRunDuration[i].ToString().Split(",");
                    if (i == alApiRunDuration.Count - 1)
                    {
                        strReportContent += str[2] + "\n";
                    }
                    else
                    {
                        strReportContent += str[2] + ",";
                    }
                }
                //紀錄最短時間
                strReportContent += "最短時間" + ",";
                for (int i = 0; i < alApiRunDuration.Count; i++)
                {
                    string[] str = alApiRunDuration[i].ToString().Split(",");
                    if (i == alApiRunDuration.Count - 1)
                    {
                        strReportContent += str[3] + "\n";
                    }
                    else
                    {
                        strReportContent += str[3] + ",";
                    }
                }

            }
            else
            {
                FileModel.WriteLog("Weekly Report Create Fail");
                FileModel.WriteLog("Can't create repory because data length is different");
            }

            return strReportContent;


            

        }



    }
}
