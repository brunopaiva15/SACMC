using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SACMC
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void BtnSwitch_Click(object sender, EventArgs e)
        {
            try
            {
                btnSwitch.Enabled = false;

                Process proc1 = new Process();
                proc1.StartInfo.FileName = "cmd";
                proc1.StartInfo.Verb = "runas";
                proc1.StartInfo.WorkingDirectory = Directory.GetCurrentDirectory();
                proc1.StartInfo.Arguments = "/k setlocal & exit";
                proc1.StartInfo.ErrorDialog = true;
                proc1.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                proc1.Start();
                proc1.WaitForExit();

                Process proc2 = new Process();
                proc2.StartInfo.FileName = "cmd";
                proc2.StartInfo.Verb = "runas";
                proc2.StartInfo.WorkingDirectory = Directory.GetCurrentDirectory();
                proc2.StartInfo.Arguments = "/k reg query HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration\\ /v CDNBaseUrl & exit";
                proc2.StartInfo.ErrorDialog = true;
                proc2.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                proc2.Start();
                proc2.WaitForExit();

                Process proc3 = new Process();
                proc3.StartInfo.FileName = "cmd";
                proc3.StartInfo.Verb = "runas";
                proc3.StartInfo.WorkingDirectory = Directory.GetCurrentDirectory();
                proc3.StartInfo.Arguments = "/k reg add HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration /v CDNBaseUrl /t REG_SZ /d \"http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60\" /f & exit";
                proc3.StartInfo.ErrorDialog = true;
                proc3.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                proc3.Start();
                proc3.WaitForExit();

                Process proc4 = new Process();
                proc4.StartInfo.FileName = "cmd";
                proc4.StartInfo.Verb = "runas";
                proc4.StartInfo.WorkingDirectory = Directory.GetCurrentDirectory();
                proc4.StartInfo.Arguments = "/k reg delete HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration /v UpdateUrl /f & exit";
                proc4.StartInfo.ErrorDialog = true;
                proc4.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                proc4.Start();
                proc4.WaitForExit();

                Process proc5 = new Process();
                proc5.StartInfo.FileName = "cmd";
                proc5.StartInfo.Verb = "runas";
                proc5.StartInfo.WorkingDirectory = Directory.GetCurrentDirectory();
                proc5.StartInfo.Arguments = "/k reg delete HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration /v UpdateToVersion /f & exit";
                proc5.StartInfo.ErrorDialog = true;
                proc5.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                proc5.Start();
                proc5.WaitForExit();

                Process proc6 = new Process();
                proc6.StartInfo.FileName = "cmd";
                proc6.StartInfo.Verb = "runas";
                proc6.StartInfo.WorkingDirectory = Directory.GetCurrentDirectory();
                proc6.StartInfo.Arguments = "/k reg delete HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Updates /v UpdateToVersion /f & exit";
                proc6.StartInfo.ErrorDialog = true;
                proc6.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                proc6.Start();
                proc6.WaitForExit();

                Process proc7 = new Process();
                proc7.StartInfo.FileName = "cmd";
                proc7.StartInfo.Verb = "runas";
                proc7.StartInfo.WorkingDirectory = Directory.GetCurrentDirectory();
                proc7.StartInfo.Arguments = "/k reg delete HKEY_LOCAL_MACHINE\\SOFTWARE\\Policies\\Microsoft\\Office\\16.0\\Common\\OfficeUpdate\\ /f & exit";
                proc7.StartInfo.ErrorDialog = true;
                proc7.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                proc7.Start();
                proc7.WaitForExit();

                Process proc8 = new Process();
                proc8.StartInfo.FileName = "cmd";
                proc8.StartInfo.Verb = "runas";
                proc8.StartInfo.WorkingDirectory = Directory.GetCurrentDirectory();
                proc8.StartInfo.Arguments = "/k \"%CommonProgramFiles%\\microsoft shared\\ClickToRun\\OfficeC2RClient.exe\" /update user & exit";
                proc8.StartInfo.ErrorDialog = true;
                proc8.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                proc8.Start();
                proc8.WaitForExit();

                Process proc9 = new Process();
                proc9.StartInfo.FileName = "cmd";
                proc9.StartInfo.Verb = "runas";
                proc9.StartInfo.WorkingDirectory = Directory.GetCurrentDirectory();
                proc9.StartInfo.Arguments = "/k Endlocal & exit";
                proc9.StartInfo.ErrorDialog = true;
                proc9.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                proc9.Start();
                proc9.WaitForExit();

                btnSwitch.Text = "SWITCHED";
            }
            catch
            {
                MessageBox.Show("Operation canceled.");
            }
        }
    }
}
