using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;

namespace Alfred
{
    enum RecycleFlags : int
    {
    // No confirmation dialog when emptying the recycle bin
    SHERB_NOCONFIRMATION = 0x00000001,
    // No progress tracking window during the emptying of the recycle bin
    SHERB_NOPROGRESSUI = 0x00000001,
    // No sound whent the emptying of the recycle bin is complete
    SHERB_NOSOUND = 0x00000004
    }
   

    public partial class Form1 : Form
    {
        SpeechRecognitionEngine recEng = new SpeechRecognitionEngine();
        SpeechSynthesizer Alfred=new SpeechSynthesizer();
        SpeechRecognitionEngine startlistening = new SpeechRecognitionEngine();//FOR MICROPHONE (stop lst and ...)
        Random rnd =new Random();
        int RecTimeOut = 0;

        //emptying the recycle bin
        [DllImport("Shell32.dll")]
        static extern int SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath, RecycleFlags dwFlags);
        //Volume
        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        //volume
        const int WM_APPCOMMAND = 0x319;

        const int APPCOMMAND_VOLUME_MUTE = 0x80000;
        const int APPCOMMAND_VOLUME_DOWN = 0x90000;
        const int APPCOMMAND_VOLUME_UP = 0xA0000;


        //console
       

        public Form1()
        {
            InitializeComponent();

            recEng.SetInputToDefaultAudioDevice();//microphone on

            recEng.LoadGrammarAsync(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"grammar.txt")))));
          
            recEng.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(recEng_SpeechRecognized);
            recEng.SpeechDetected += new EventHandler<SpeechDetectedEventArgs>(recEng_SpeechDetected);


            recEng.RecognizeAsync(RecognizeMode.Multiple);

            startlistening.SetInputToDefaultAudioDevice();
            startlistening.LoadGrammarAsync(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"grammar.txt")))));
            recEng.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(start_SpeechRecognized);


        }

        private void recEng_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            int ranNum;
            string speech = e.Result.Text;
            if (speech == "Hello")
            {
                Alfred.SpeakAsync("My name is Alfred I am here to help you!");
            }
            if (speech == "Show commands")
            {
                Alfred.SpeakAsync("Yes mis");
                string[] commands = (File.ReadAllLines(@"grammar.txt"));
                listBox1.Items.Clear();
                listBox1.SelectionMode = SelectionMode.None;
                listBox1.Visible=true;
                foreach(string command in commands)
                {
                    listBox1.Items.Add(command);
                }
            }
            if (speech == "Hide commands")
            {
                Alfred.SpeakAsync("Yes mis");
                listBox1.Visible = false;
            }
            if (speech == "empty trash")
            {
                Alfred.SpeakAsync("Yes mis");
                SHEmptyRecycleBin(IntPtr.Zero, null, RecycleFlags.SHERB_NOSOUND | RecycleFlags.SHERB_NOCONFIRMATION);
                Thread.Sleep(1000);
            }
            if (speech == "volume up")
            {
                Alfred.SpeakAsync("Yes mis");
                //Console.Beep();
                SendMessage(this.Handle, WM_APPCOMMAND, IntPtr.Zero, (IntPtr)APPCOMMAND_VOLUME_UP);
                Thread.Sleep(1000);
            }
            if (speech == "volume down")
            {
                Alfred.SpeakAsync("Yes mis");
                SendMessage(this.Handle, WM_APPCOMMAND, IntPtr.Zero, (IntPtr)APPCOMMAND_VOLUME_DOWN);
            }
            if (speech == "mute")
            {
                Alfred.SpeakAsync("Yes mis");
                SendMessage(this.Handle, WM_APPCOMMAND, IntPtr.Zero, (IntPtr)APPCOMMAND_VOLUME_MUTE);
            }
            if (speech == "time")
            {
                Alfred.SpeakAsync(DateTime.Now.ToString("HH:mm"));
                Thread.Sleep(1000);
            }
            if (speech == "date")
            {
                Alfred.SpeakAsync(DateTime.Now.ToString("d"));
            }
            if (speech == "c m d")
            {
                Alfred.SpeakAsync("Yes mis");
                Process.Start("C:\\Windows\\System32\\cmd.exe", "/k @echo Hello & cd C:\\");
            }
            if (speech == "open settings")
            {
                Alfred.SpeakAsync("Yes mis");
                Process.Start("explorer.exe", @"shell:::{BB06C0E4-D293-4f75-8A90-CB05B6477EEE}");
            }
            if (speech == "open settings")
            {
                Alfred.SpeakAsync("Yes mis");
                Process.Start("explorer.exe", @"shell:::{BB06C0E4-D293-4f75-8A90-CB05B6477EEE}");
            }
            if (speech == "open opera")
            {
                Alfred.SpeakAsync("Yes mis");
                System.Diagnostics.Process.Start("http://google.com");
            }
            if (speech == "play music")
            {
                Alfred.SpeakAsync("Yes mis");
                System.Diagnostics.Process.Start("Spotify.exe");
            }
            if (speech == "Stop listening")
            {
                Alfred.SpeakAsyncCancelAll();
                ranNum = rnd.Next(1);
                if(ranNum == 1)
                {
                    Alfred.SpeakAsync("Yes");
                }
                if (ranNum == 2)
                {
                    Alfred.SpeakAsync("I am sorry ,I will be quiet");
                }
            }
        }
       
        private void recEng_SpeechDetected(object sender, SpeechDetectedEventArgs e)
        {
            RecTimeOut = 0;
        }

        private void start_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string speech = e.Result.Text;
            if(speech == "Wake up")
            {
                startlistening.RecognizeAsyncCancel();
                Alfred.SpeakAsync("Yes");
                recEng.RecognizeAsync(RecognizeMode.Multiple);
            }

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void tmr_Tick(object sender, EventArgs e)
        {
            if (RecTimeOut == 10)
            {
                recEng.RecognizeAsyncCancel();
            }
            else if (RecTimeOut == 11)
            {
                tmr.Stop();
                startlistening.RecognizeAsync(RecognizeMode.Multiple);
                RecTimeOut = 0;
            }
        }
    }
}
