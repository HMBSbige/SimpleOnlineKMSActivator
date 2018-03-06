using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace SimpleOnlineKMSActivator
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            
        }

        private static string GetCMDOutput(string cmd)
        {
            const string exit = @"exit";
            var p = new Process
            {
                StartInfo =
                {
                    FileName = @"cmd.exe",
                    UseShellExecute = false,
                    RedirectStandardInput = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                }
            };
            p.Start();
            p.StandardInput.AutoFlush = true;
            p.StandardInput.WriteLine(cmd);
            p.StandardInput.WriteLine(exit);
            var outStr = p.StandardOutput.ReadToEnd();
            p.WaitForExit();
            p.Close();

            var str1 = Application.StartupPath + @">" + cmd + Environment.NewLine;
            var str2 = Application.StartupPath + @">" + exit + Environment.NewLine;
            var ret = outStr.Substring(outStr.IndexOf(str1, StringComparison.Ordinal) + str1.Length,
                outStr.IndexOf(str2, StringComparison.Ordinal) - (outStr.IndexOf(str1, StringComparison.Ordinal) + str1.Length));
            var ret2 = ret.Trim().Split('\n');
            return ret2[ret2.Length - 1].Trim();
        }

        private static string GetWindowsVersion()
        {
            const string cmd = @"wmic os get caption";
            return GetCMDOutput(cmd);
        }

        private static string GetWindowsLicense()
        {
            const string cmd = @"cscript %SystemRoot%\System32\slmgr.vbs /xpr";
            return GetCMDOutput(cmd);
        }

        private static string SetLicenseKey(string key)
        {
            var cmd = @"cscript %SystemRoot%\System32\slmgr.vbs /ipk " + key;
            return GetCMDOutput(cmd);
        }

        private static string SetKMSServer(string serveraddr)
        {
            var cmd = @"cscript %SystemRoot%\System32\slmgr.vbs /skms " + serveraddr;
            return GetCMDOutput(cmd);
        }

        private static string ActivateWindows()
        {
            const string cmd = @"cscript %SystemRoot%\System32\slmgr.vbs /ato";
            return GetCMDOutput(cmd);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            textBox1.Text = string.Empty;
            textBox1.Text += GetWindowsVersion() + Environment.NewLine;
            progressBar1.Value = 0;
            textBox1.Text += SetLicenseKey(@"W269N-WFGWX-YVC9B-4J6C9-T83GX") + Environment.NewLine;
            progressBar1.PerformStep();
            textBox1.Text += SetKMSServer(@"kms.bige0.com") + Environment.NewLine;
            progressBar1.PerformStep();
            textBox1.Text += ActivateWindows() + Environment.NewLine;
            progressBar1.PerformStep();
            textBox1.Text += GetWindowsLicense() + Environment.NewLine;
            button1.Enabled = true;
        }
    }
}
