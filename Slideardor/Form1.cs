using PP = Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Core;
using System.Timers;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace Slideardor
{
    public partial class Form1 : Form
    {
        private string configPath;
        private List<string> slidesToShow = new List<string>();
        private bool autoStart = false;
        private bool stayinloop = false;
        private bool killapp = false;
        private int appId = 0;

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        public Form1()
        {
            InitializeComponent();

            configPath = Path.GetDirectoryName(Application.ExecutablePath) + @"\slideador.config";
            SetBuffedConfig();

            appId = (Constants.CTRL) ^ (int)Keys.C ^ Handle.ToInt32();
            RegisterKey();

            if (autoStart)
                button1_Click(null, null);
        }

        private void SetBuffedConfig()
        {
            if (!File.Exists(configPath))
                return;

            try
            {
                string configText = File.ReadAllText(configPath).Replace("\r", "").Replace("\n", "");
                string[] configsLines = configText.Split(';');
                string[] configReg;

                foreach (var item in configsLines)
                {
                    configReg = item.Split('=');

                    switch (configReg[0].ToLower())
                    {
                        case "searchfolder":
                            textBox1.Text = configReg[1];
                            break;
                        case "interval":
                            numericUpDown1.Value = int.Parse(configReg[1]);
                            break;
                        case "autostart":
                            autoStart = bool.Parse(configReg[1]);
                            break;
                        case "stayinloop":
                            stayinloop = bool.Parse(configReg[1]);
                            break;
                    }
                }
            }
            catch (Exception e)
            {
                textBox2.AppendText(string.Format("Ocorreu um erro ao ler o arquivo de configuração\n\n {0}", e.Message));
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.SelectedPath = textBox1.Text;

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog1.SelectedPath;

                File.WriteAllText(configPath, textBox1.Text);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DetectSlidesPath();
            ShowSlides();
        }

        private void DetectSlidesPath()
        {
            textBox2.Clear();
            slidesToShow.Clear();

            foreach (var filePath in Directory.EnumerateFiles(textBox1.Text, "*.pptx"))
            {
                slidesToShow.Add(filePath);
                textBox2.AppendText(filePath + Environment.NewLine);
            }
        }

        PP.SlideShowView _showView;
        int slidesCount = 0;
        bool hasFinished = false;
        private void ShowSlides()
        {
            foreach (var slidePath in slidesToShow)
            {
                try
                {
                    PP.Application pptApplication = new PP.Application();
                    PP.Presentation pptPresentation = pptApplication.Presentations.Open(slidePath, 
                        Untitled: MsoTriState.msoTrue, WithWindow: MsoTriState.msoFalse);
                    PP.Slides slides = pptPresentation.Slides;
                    slidesCount = slides.Count;

                    pptPresentation.SlideShowSettings.ShowPresenterView = MsoTriState.msoFalse;
                    pptPresentation.SlideShowSettings.Run();

                    _showView = pptPresentation.SlideShowWindow.View;
                    hasFinished = false;

                    var slideTest = new System.Timers.Timer((int)numericUpDown1.Value * 1000);
                    slideTest.AutoReset = true;
                    slideTest.Elapsed += new ElapsedEventHandler(slidetest_Elapsed);
                    slideTest.Start();

                    while (!hasFinished) { Application.DoEvents(); }

                    Process[] pros = Process.GetProcesses();
                    for (int i = 0; i < pros.Count(); i++)
                    {
                        if (pros[i].ProcessName.ToLower().Contains("powerpnt"))
                        {
                            pros[i].Kill();
                        }
                    }
                }
                catch (Exception e)
                {
                    textBox2.AppendText(string.Format("Ocorreu um erro no arquivo {1}\n\n {0}", e.Message, slidePath));
                }
            }

            // Repete o loop
            if (!killapp && stayinloop)
                button1_Click(null, null);
        }

        void slidetest_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                if (_showView.CurrentShowPosition == slidesCount || killapp)
                {
                    _showView.Exit();
                    (sender as System.Timers.Timer).Stop();
                    hasFinished = true;
                    return;
                }

                _showView.Application.SlideShowWindows[1].Activate();
                _showView.Next();
            }
            catch
            {
                (sender as System.Timers.Timer).Stop();
                hasFinished = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DetectSlidesPath();
        }

        public bool RegisterKey()
        {
            return RegisterHotKey(Handle, appId, Constants.CTRL, (int)Keys.C);
        }

        public bool UnregiserKey()
        {
            return UnregisterHotKey(Handle, appId);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == Constants.WM_HOTKEY_MSG_ID)
                killapp = true;

            base.WndProc(ref m);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            UnregiserKey();
        }
    }

    public static class Constants
    {
        //modifiers
        public const int NOMOD = 0x0000;
        public const int ALT = 0x0001;
        public const int CTRL = 0x0002;
        public const int SHIFT = 0x0004;
        public const int WIN = 0x0008;

        //windows message id for hotkey
        public const int WM_HOTKEY_MSG_ID = 0x0312;
    }
}
