using System;
using System.Configuration;
using System.IO;
using System.Text;

namespace NewmanCMD
{   
    class FileModel
    {
        static string myDate;

        public static void SetDate(string date)
        {
            myDate = date;
        }


        public static void DeleteOldFiles(string dirPath, int days)
        {
            try
            {
                if (!Directory.Exists(dirPath) || days < 1) return;

                DirectoryInfo dir = new DirectoryInfo(dirPath);
                FileInfo[] files = dir.GetFiles();
                foreach (FileInfo file in files)
                {
                    if (file.LastWriteTime < DateTime.Now.AddDays(-days))
                    {
                        file.Delete();
                    }
                }
            }
            catch (Exception) { }
        }


        public static void WriteLog(string strLogMessage)
        {

            string strLogFileName = "Log_" + myDate + ".txt";
            string strLogDirPath = ConfigurationManager.AppSettings["logDirPath"];
            string strLogPath = strLogDirPath + "/" + strLogFileName;

            if (!Directory.Exists(strLogDirPath))
            {
                //建立資料夾
                Directory.CreateDirectory(strLogDirPath);
            }

            if (!File.Exists(strLogPath))
            {
                //建立檔案
                var myfile = File.Create(strLogPath);
                myfile.Close();
            }



            DateTime dateTime = DateTime.Now;
            String strNowTime = dateTime.ToString("yyyy-MM-dd HH:mm:ss");

            StreamWriter sw = new StreamWriter(strLogPath,true);

            sw.WriteLine("--執行時間 " + strNowTime + "--");
            sw.WriteLine(strLogMessage);
            sw.WriteLine();
            sw.Close();

        }


        public static void CreateWeeklyReport(string strReportContent)
        {
            DateTime dateTime = DateTime.Now;
            string start = dateTime.AddDays(-8).ToString("yyyyMMdd");
            string end = dateTime.AddDays(-1).ToString("yyyyMMdd");
            string strFileName = "WeeklyReport" + start + "To" + end + ".csv";
            string strReportPath = ConfigurationManager.AppSettings["weeklyReportDirPath"] + "/" + strFileName;
            Console.WriteLine(strReportPath);

            //儲存帶有BOM表的檔案
            var enc = new UTF8Encoding(true);
            Byte[] bytes = enc.GetBytes(strReportContent);
            var fs = new FileStream(strReportPath, FileMode.Create);
            Byte[] preamble = enc.GetPreamble();
            fs.Write(preamble, 0, preamble.Length);
            fs.Write(bytes, 0, bytes.Length);
            fs.Close();
        }

    }
}
