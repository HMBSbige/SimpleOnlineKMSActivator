using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimpleOnlineKMSActivator
{
    public partial class MainForm : Form
    {
        #region Form

        public MainForm()
        {
            InitializeComponent();
            _label3CallBack = ChangeLabel3;
            _textBox1Add1CallBack = TextBox1Add;
            _getcombobox1 = GetComboBox1Text;
            _getcombobox2 = GetComboBox2Text;
            _progressBarPerformStepCallBack = UpdateProgressBar1;
            _textBox1ChangeCallBack = TextBox1Set;
            _progressBarResetCallBack = ResetProgressBar1;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            var t = new Task(() =>
            {
                label_version.Invoke(_label3CallBack, GetWindowsVersion());
            });
            t.Start();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                DisableControl();
                var t = new Task(() =>
                {
                    var key = Invoke(_getcombobox1) as string;
                    var server = Invoke(_getcombobox2) as string;
                    textBox1.Invoke(_textBox1ChangeCallBack, string.Empty);
                    progressBar1.Invoke(_progressBarResetCallBack);
                    textBox1.Invoke(_textBox1Add1CallBack, SetLicenseKey(key) + Environment.NewLine);
                    progressBar1.Invoke(_progressBarPerformStepCallBack);
                    textBox1.Invoke(_textBox1Add1CallBack, SetKMSServer(server) + Environment.NewLine);
                    progressBar1.Invoke(_progressBarPerformStepCallBack);
                    textBox1.Invoke(_textBox1Add1CallBack, ActivateWindows() + Environment.NewLine);
                    progressBar1.Invoke(_progressBarPerformStepCallBack);
                    textBox1.Invoke(_textBox1Add1CallBack, GetWindowsLicense() + Environment.NewLine);
                });
                t.Start();
                t.ContinueWith(task =>
                {
                    BeginInvoke(new VoidMethodDelegate(EnableControl));
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"出错了", MessageBoxButtons.OK, MessageBoxIcon.Error);
                EnableControl();
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(@"https://github.com/HMBSbige/kms-server#windows-gvlk-%E5%AF%86%E9%92%A5%E5%AF%B9%E7%85%A7%E8%A1%A8kms-%E6%BF%80%E6%B4%BB%E4%B8%93%E7%94%A8");
        }

        #endregion

        #region delegate_define

        private delegate void LabelCallBack(string str);
        private readonly LabelCallBack _label3CallBack;

        private delegate void VoidMethodDelegate();
        private readonly VoidMethodDelegate _progressBarPerformStepCallBack;
        private readonly VoidMethodDelegate _progressBarResetCallBack;

        private delegate void TextBoxTextCallBack(string str);
        private readonly TextBoxTextCallBack _textBox1Add1CallBack;
        private readonly TextBoxTextCallBack _textBox1ChangeCallBack;

        private delegate string GetTextCallBack();
        private readonly GetTextCallBack _getcombobox1;
        private readonly GetTextCallBack _getcombobox2;

        #endregion

        #region delegate_function

        private void ChangeLabel3(string str)
        {
            label_version.Text += str;
        }

        private void TextBox1Add(string str)
        {
            textBox1.Text += str;
        }

        private void TextBox1Set(string str)
        {
            textBox1.Text = str;
        }

        private string GetComboBox1Text()
        {
            return comboBox1.Text;
        }

        private string GetComboBox2Text()
        {
            return comboBox2.Text;
        }

        private void ResetProgressBar1()
        {
            progressBar1.Value = 0;
        }

        private void UpdateProgressBar1()
        {
            progressBar1.PerformStep();
        }

        private void DisableControl()
        {
            button1.Enabled = false;
            comboBox1.Enabled = false;
            comboBox2.Enabled = false;
            comboBox3.Enabled = false;
        }

        private void EnableControl()
        {
            button1.Enabled = true;
            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
            comboBox3.Enabled = true;
        }

        #endregion
        
        #region basefunction

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

        #endregion

    }
}