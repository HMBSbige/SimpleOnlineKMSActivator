using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace SimpleOnlineKMSActivator
{
    public partial class MainForm : Form
    {
        #region Keys

        private static readonly Dictionary<string, string> Kmsclientkeys = new Dictionary<string, string> {
{ @"Windows Server Datacenter", @"6Y6KB-N82V8-D8CQV-23MJW-BWTG6" },
{ @"Windows Server Standard", @"DPCNP-XQFKJ-BJF7R-FRC8D-GF6G4" },
{ @"Windows Server 2016 Datacenter", @"CB7KF-BWN84-R7R2Y-793K2-8XDDG" },
{ @"Windows Server 2016 Standard", @"WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY" },
{ @"Windows Server 2016 Essentials", @"JCKRF-N37P4-C2D82-9YXRT-4M63B" },
{ @"Windows 10 Professional Workstation", @"NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J" },
{ @"Windows 10 Professional Workstation N", @"9FNHH-K3HBT-3W4TD-6383H-6XYWF" },
{ @"Windows 10 Professional", @"W269N-WFGWX-YVC9B-4J6C9-T83GX" },
{ @"Windows 10 Professional N", @"MH37W-N47XK-V7XM9-C7227-GCQG9" },
{ @"Windows 10 Enterprise", @"NPPR9-FWDCX-D2C8J-H872K-2YT43" },
{ @"Windows 10 Enterprise N", @"DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4" },
{ @"Windows 10 Education", @"NW6C2-QMPVW-D7KKK-3GKT6-VCFB2" },
{ @"Windows 10 Education N", @"2WH4N-8QGBV-H22JP-CT43Q-MDWWJ" },
{ @"Windows 10 Enterprise 2015 LTSB", @"WNMTR-4C88C-JK8YV-HQ7T2-76DF9" },
{ @"Windows 10 Enterprise 2015 LTSB N", @"2F77B-TNFGY-69QQF-B8YKP-D69TJ" },
{ @"Windows 10 Enterprise 2016 LTSB", @"DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ" },
{ @"Windows 10 Enterprise 2016 LTSB N", @"QFFDN-GRT3P-VKWWX-X7T3R-8B639" },
{ @"Windows 8.1 Professional", @"GCRJD-8NW9H-F2CDX-CCM8D-9D6T9" },
{ @"Windows 8.1 Professional N", @"HMCNV-VVBFX-7HMBH-CTY9B-B4FXY" },
{ @"Windows 8.1 Enterprise", @"MHF9N-XY6XB-WVXMC-BTDCT-MKKG7" },
{ @"Windows 8.1 Enterprise N", @"TT4HM-HN7YT-62K67-RGRQJ-JFFXW" },
{ @"Windows Server 2012 R2 Server Standard", @"D2N9P-3P6X9-2R39C-7RTCD-MDVJX" },
{ @"Windows Server 2012 R2 Datacenter", @"W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9" },
{ @"Windows Server 2012 R2 Essentials", @"KNC87-3J2TX-XB4WP-VCPJV-M4FWM" },
{ @"Windows 8 Professional", @"NG4HW-VH26C-733KW-K6F98-J8CK4" },
{ @"Windows 8 Professional N", @"XCVCF-2NXM9-723PB-MHCB7-2RYQQ" },
{ @"Windows 8 Enterprise", @"32JNW-9KQ84-P47T8-D8GGY-CWCK7" },
{ @"Windows 8 Enterprise N", @"JMNMF-RHW7P-DMY6X-RF3DR-X2BQT" },
{ @"Windows Server 2012", @"BN3D2-R7TKB-3YPBD-8DRP2-27GG4" },
{ @"Windows Server 2012 N", @"8N2M2-HWPGY-7PGT9-HGDD8-GVGGY" },
{ @"Windows Server 2012 Single Language", @"2WN2H-YGCQR-KFX6K-CD6TF-84YXQ" },
{ @"Windows Server 2012 Country Specific", @"4K36P-JN4VD-GDC6V-KDT89-DYFKP" },
{ @"Windows Server 2012 Server Standard", @"XC9B7-NBPP2-83J2H-RHMBY-92BT4" },
{ @"Windows Server 2012 MultiPoint Standard", @"HM7DN-YVMH3-46JC3-XYTG7-CYQJJ" },
{ @"Windows Server 2012 MultiPoint Premium", @"XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G" },
{ @"Windows Server 2012 Datacenter", @"48HP8-DN98B-MYWDG-T2DCC-8W83P" },
{ @"Windows 7 Professional", @"FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4" },
{ @"Windows 7 Professional N", @"MRPKT-YTG23-K7D7T-X2JMM-QY7MG" },
{ @"Windows 7 Professional E", @"W82YF-2Q76Y-63HXB-FGJG9-GF7QX" },
{ @"Windows 7 Enterprise", @"33PXH-7Y6KF-2VJC9-XBBR8-HVTHH" },
{ @"Windows 7 Enterprise N", @"YDRBP-3D83W-TY26F-D46B2-XCKRJ" },
{ @"Windows 7 Enterprise E", @"C29WB-22CC8-VJ326-GHFJW-H9DH4" },
{ @"Windows Server 2008 R2 Web", @"6TPJF-RBVHG-WBW2R-86QPH-6RTM4" },
{ @"Windows Server 2008 R2 HPC edition", @"TT8MH-CG224-D3D7Q-498W2-9QCTX" },
{ @"Windows Server 2008 R2 Standard", @"YC6KT-GKW9T-YTKYR-T4X34-R7VHC" },
{ @"Windows Server 2008 R2 Enterprise", @"489J6-VHDMP-X63PK-3K798-CPX3Y" },
{ @"Windows Server 2008 R2 Datacenter", @"74YFP-3QFB3-KQT8W-PMXWJ-7M648" },
{ @"Windows Server 2008 R2 for Itanium-based Systems", @"GT63C-RJFQ3-4GMB6-BRFB9-CB83V" },
{ @"Windows Vista Business", @"YFKBB-PQJJV-G996G-VWGXY-2V3X8" },
{ @"Windows Vista Business N", @"HMBQG-8H2RH-C77VX-27R82-VMQBT" },
{ @"Windows Vista Enterprise", @"VKK3X-68KWM-X2YGT-QR4M6-4BWMV" },
{ @"Windows Vista Enterprise N", @"VTC42-BM838-43QHV-84HX6-XJXKV" },
{ @"Windows Web Server 2008", @"WYR28-R7TFJ-3X2YQ-YCY4H-M249D" },
{ @"Windows Server 2008 Standard", @"TM24T-X9RMF-VWXK6-X8JC9-BFGM2" },
{ @"Windows Server 2008 Standard without Hyper-V", @"W7VD6-7JFBR-RX26B-YKQ3Y-6FFFJ" },
{ @"Windows Server 2008 Enterprise", @"YQGMW-MPWTJ-34KDK-48M3W-X4Q6V" },
{ @"Windows Server 2008 Enterprise without Hyper-V", @"39BXF-X8Q23-P2WWT-38T2F-G3FPG" },
{ @"Windows Server 2008 HPC", @"RCTX3-KWVHP-BR6TB-RB6DM-6X7HP" },
{ @"Windows Server 2008 Datacenter", @"7M67G-PC374-GR742-YH8V4-TCBY3" },
{ @"Windows Server 2008 Datacenter without Hyper-V", @"22XQ2-VRXRG-P8D42-K34TD-G3QQC" },
{ @"Windows Server 2008 for Itanium-Based Systems", @"4DWFP-JF3DJ-B7DTH-78FJB-PDRHK" }
        };

        #endregion

        #region Form

        public MainForm()
        {
            InitializeComponent();
            //窗口初始化
            foreach (var key in Kmsclientkeys)
            {
                comboBox3.Items.Add(key.Key);
            }
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 7;
            var t = new Task(() =>
            {
                label_version.Invoke(_label3CallBack, GetWindowsVersion());
                textBox_OfficePath.Invoke(_textBox2ChangeCallBack, GetOfficePath());
            });
            t.Start();
            //绑定委托
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
            _textBox2ChangeCallBack = SetTextbox2Text;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            
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
                        $@"cscript %SystemRoot%\System32\slmgr.vbs /ipk {key}" + Environment.NewLine +
                        $@"cscript %SystemRoot%\System32\slmgr.vbs /skms {server}"+ Environment.NewLine +
                        @"cscript %SystemRoot%\System32\slmgr.vbs /ato" + Environment.NewLine);
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
                    textBox1.Invoke(_textBox1Add1CallBack, ActivateOffice(officePath));
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
                SelectedPath = textBox_OfficePath.Text
            };
            path.ShowDialog();
            if (path.SelectedPath != string.Empty)
            {
                textBox_OfficePath.Text = path.SelectedPath;
            }
        }

        private void button6_Click(object sender, EventArgs e)
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
                    textBox1.Invoke(_textBox1ChangeCallBack, string.Empty);
                    textBox1.Invoke(_textBox1Add1CallBack, GetOfficeLicense(officePath));
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

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                DisableControl();
                var t = new Task(() =>
                {
                    var officePath = Invoke(_gettextbox2) as string;
                    var server = Invoke(_getcombobox2) as string;
                    var cmd1 = $@"cscript ""{officePath}\OSPP.VBS"" /sethst:{server}";
                    var cmd2 = $@"cscript ""{officePath}\OSPP.VBS"" /act";
                    textBox1.Invoke(_textBox1ChangeCallBack, string.Empty);
                    textBox1.Invoke(_textBox1Add1CallBack,
                        cmd1 + Environment.NewLine +
                        cmd2 + Environment.NewLine);

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
        
        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                var t = new Task(() =>
                {
                    textBox_OfficePath.Invoke(_textBox2ChangeCallBack, GetOfficePath());
                });
                t.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"出错了", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox2_DragDrop(object sender, DragEventArgs e)
        {
            textBox_OfficePath.Text = ((Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
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
            if (File.Exists($@"{textBox_OfficePath.Text}\OSPP.VBS"))
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
        
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                WindowsKey.Text = Kmsclientkeys[comboBox3.Text];
            }
            catch
            {
                WindowsKey.Text = @"";
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
        private readonly TextBoxTextCallBack _textBox2ChangeCallBack;

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
            return WindowsKey.Text;
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
            button5.Enabled = false;
            button6.Enabled = false;
            button7.Enabled = false;
            button8.Enabled = false;
            WindowsKey.Enabled = false;
            comboBox2.Enabled = false;
            comboBox3.Enabled = false;
        }

        private void EnableControl()
        {
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            button5.Enabled = true;
            button6.Enabled = true;
            button7.Enabled = true;
            button8.Enabled = true;
            WindowsKey.Enabled = true;
            comboBox2.Enabled = true;
            comboBox3.Enabled = true;
        }

        private string GetTextbox2Text()
        {
            return textBox_OfficePath.Text;
        }

        private void SetTextbox2Text(string str)
        {
            textBox_OfficePath.Text = str;
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
                ret += ret2[i] + '\n';
            }
            return ret.Remove(ret.Length - 1) + Environment.NewLine;
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

        private static string GetOfficeLicense(string officePath)
        {
            var cmd = $@"cscript ""{officePath}\OSPP.VBS"" /dstatus";
            return GetOfficeOutput(cmd);
        }

        private static string GetOfficePath()
        {
            string[] officeversion = {@"7.0", @"8.0", @"9.0", @"10.0", @"11.0", @"12.0", @"14.0", @"15.0", @"16.0"};
            foreach (var ver in officeversion)
            {
                var officePath = Registry.LocalMachine.OpenSubKey($@"Software\Microsoft\Office\{ver}\Common\InstallRoot", false);
                if (officePath != null)
                {
                    return officePath.GetValue(@"Path", string.Empty) as string;
                }
            }
            return string.Empty;
        }

        #endregion
    }
}