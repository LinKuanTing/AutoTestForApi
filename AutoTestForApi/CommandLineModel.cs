using System;
using System.Diagnostics;


namespace NewmanCMD
{
    class CommandLineModel
    {
        public static void excuteCMD(string strCmdText)
        {
            //執行command.exe
            Process p = new Process();
            string str = null;

            p.StartInfo.FileName = "cmd.exe";

            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardInput = true;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.CreateNoWindow = true; //不跳出cmd視窗

            p.Start();
            p.StandardInput.WriteLine(@strCmdText);

            p.StandardInput.WriteLine("exit");

            str = p.StandardOutput.ReadToEnd();
            p.WaitForExit();
            p.Close();


            //System.Diagnostics.Trace.WriteLine(str);
        }
    }
}
