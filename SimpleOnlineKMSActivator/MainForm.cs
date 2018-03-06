using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
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
            _progressBar1PerformStepCallBack = UpdateProgressBar1;
            _progressBar1ResetCallBack = ResetProgressBar1;
            _progressBar2PerformStepCallBack = UpdateProgressBar2;
            _progressBar2ResetCallBack = ResetProgressBar2;
            _textBox1ChangeCallBack = TextBox1Set;
            _gettextbox2 = GetTextbox2Text;
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
                    progressBar1.Invoke(_progressBar1ResetCallBack);
                    textBox1.Invoke(_textBox1Add1CallBack, SetLicenseKey(key) + Environment.NewLine);
                    progressBar1.Invoke(_progressBar1PerformStepCallBack);
                    textBox1.Invoke(_textBox1Add1CallBack, SetWindowsKMSServer(server) + Environment.NewLine);
                    progressBar1.Invoke(_progressBar1PerformStepCallBack);
                    textBox1.Invoke(_textBox1Add1CallBack, ActivateWindows() + Environment.NewLine);
                    progressBar1.Invoke(_progressBar1PerformStepCallBack);
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

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                DisableControl();
                var t = new Task(() =>
                {
                    textBox1.Invoke(_textBox1ChangeCallBack, string.Empty);
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

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                DisableControl();
                var t = new Task(() =>
                {
                    var key = Invoke(_getcombobox1) as string;
                    var server = Invoke(_getcombobox2) as string;
                    textBox1.Invoke(_textBox1ChangeCallBack, string.Empty);
                    textBox1.Invoke(_textBox1Add1CallBack,
                        @"cscript %SystemRoot%\System32\slmgr.vbs /ipk " + key + Environment.NewLine +
                        @"cscript %SystemRoot%\System32\slmgr.vbs /skms " + server + Environment.NewLine +
                        @"cscript %SystemRoot%\System32\slmgr.vbs /ato");
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
        
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (label3.ForeColor == Color.Red)
                {
                    MessageBox.Show(@"目录不正确！请选择正确的 Office 目录", @"错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                DisableControl();
                var t = new Task(() =>
                {
                    var officePath = Invoke(_gettextbox2) as string;
                    var server = Invoke(_getcombobox2) as string;
                    textBox1.Invoke(_textBox1ChangeCallBack, string.Empty);
                    progressBar2.Invoke(_progressBar2ResetCallBack);
                    textBox1.Invoke(_textBox1Add1CallBack, SetOfficeKMSServer(server, officePath) + Environment.NewLine);
                    progressBar2.Invoke(_progressBar2PerformStepCallBack);
                    textBox1.Invoke(_textBox1Add1CallBack, ActivateOffice(officePath) + Environment.NewLine);
                    progressBar2.Invoke(_progressBar2PerformStepCallBack);
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

        private void button5_Click(object sender, EventArgs e)
        {
            var path = new FolderBrowserDialog
            {
                Description = @"打开",
                ShowNewFolderButton = false,
                SelectedPath = textBox2.Text
            };
            path.ShowDialog();
            if (path.SelectedPath != string.Empty)
            {
                textBox2.Text = path.SelectedPath;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_DragDrop(object sender, DragEventArgs e)
        {
            textBox2.Text = ((Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
        }

        private void textBox2_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Link;
            }
            else
            {
                e.Effect = DragDropEffects.Link;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (File.Exists($@"{textBox2.Text}\OSPP.VBS"))
            {
                label3.Text = @"目录正确";
                label3.ForeColor = Color.Green;
            }
            else
            {
                label3.Text = @"该目录下不存在 OSPP.VBS";
                label3.ForeColor = Color.Red;
            }
        }

        #endregion

        #region delegate_define

        private delegate void LabelCallBack(string str);
        private readonly LabelCallBack _label3CallBack;

        private delegate void VoidMethodDelegate();
        private readonly VoidMethodDelegate _progressBar1PerformStepCallBack;
        private readonly VoidMethodDelegate _progressBar1ResetCallBack;
        private readonly VoidMethodDelegate _progressBar2PerformStepCallBack;
        private readonly VoidMethodDelegate _progressBar2ResetCallBack;

        private delegate void TextBoxTextCallBack(string str);
        private readonly TextBoxTextCallBack _textBox1Add1CallBack;
        private readonly TextBoxTextCallBack _textBox1ChangeCallBack;

        private delegate string GetTextCallBack();
        private readonly GetTextCallBack _getcombobox1;
        private readonly GetTextCallBack _getcombobox2;
        private readonly GetTextCallBack _gettextbox2;

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

        private void ResetProgressBar2()
        {
            progressBar2.Value = 0;
        }

        private void UpdateProgressBar2()
        {
            progressBar2.PerformStep();
        }

        private void DisableControl()
        {
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button6.Enabled = false;
            button7.Enabled = false;
            comboBox1.Enabled = false;
            comboBox2.Enabled = false;
            comboBox3.Enabled = false;
        }

        private void EnableControl()
        {
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            button6.Enabled = true;
            button7.Enabled = true;
            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
            comboBox3.Enabled = true;
        }

        private string GetTextbox2Text()
        {
            return textBox2.Text;
        }

        #endregion

        #region basefunction

        private static string GetOfficeOutput(string cmd)
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
            ret = string.Empty;
            for (var i = 3; i < ret2.Length; ++i)
            {
                ret += ret2[i]+ '\n';
            }
            return ret;
        }

        private static string GetSlmgrOutput(string cmd)
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
            return GetSlmgrOutput(cmd);
        }

        private static string GetWindowsLicense()
        {
            const string cmd = @"cscript %SystemRoot%\System32\slmgr.vbs /xpr";
            return GetSlmgrOutput(cmd);
        }

        private static string SetLicenseKey(string key)
        {
            var cmd = @"cscript %SystemRoot%\System32\slmgr.vbs /ipk " + key;
            return GetSlmgrOutput(cmd);
        }

        private static string SetWindowsKMSServer(string server)
        {
            var cmd = @"cscript %SystemRoot%\System32\slmgr.vbs /skms " + server;
            return GetSlmgrOutput(cmd);
        }

        private static string ActivateWindows()
        {
            const string cmd = @"cscript %SystemRoot%\System32\slmgr.vbs /ato";
            return GetSlmgrOutput(cmd);
        }

        private static string SetOfficeKMSServer(string server,string officePath)
        {
            var cmd = $@"cscript ""{officePath}\OSPP.VBS"" /sethst:{server}";
            return GetOfficeOutput(cmd);
        }

        private static string ActivateOffice(string officePath)
        {
            var cmd = $@"cscript ""{officePath}\OSPP.VBS"" /act";
            return GetOfficeOutput(cmd);
        }

        #endregion
        
    }
}