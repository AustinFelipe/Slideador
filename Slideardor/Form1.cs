using PP = Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Core;
using System.Timers;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Linq;

namespace Slideardor
{
    public partial class Form1 : Form
    {
        private string configPath;
        private List<string> slidesToShow = new List<string>();

        public Form1()
        {
            InitializeComponent();

            configPath = Path.GetDirectoryName(Application.ExecutablePath) + @"\lastsession.config";
            SetBuffedConfig();
        }

        private void SetBuffedConfig()
        {
            if (File.Exists(configPath))
                textBox1.Text = File.ReadAllText(configPath);
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
                PP.Application pptApplication = new PP.Application();
                PP.Presentation pptPresentation = pptApplication.Presentations.Open(slidePath, 
                    Untitled: MsoTriState.msoTrue);
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
        }

        void slidetest_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (_showView.CurrentShowPosition == slidesCount)
            {
                _showView.Exit();
                (sender as System.Timers.Timer).Stop();
                hasFinished = true;
                return;
            }

            _showView.Application.SlideShowWindows[1].Activate();
            _showView.Next();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DetectSlidesPath();
        }
    }
}
