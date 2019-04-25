using SimpleTCP;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Speech.Synthesis;
using System.Configuration;
using System.IO;
using WriteLogIn;

namespace TCPTTS
{
    public partial class Form1 : Form
    {
        private SpeechSynthesizer speech;

        /// <summary>
        /// 音量
        /// </summary>
        private int value = int.Parse(ConfigurationManager.AppSettings["volume"]);
        /// <summary>
        /// 語速
        /// </summary>
        private int rate = int.Parse(ConfigurationManager.AppSettings["rate"]);
        /// <summary>
        /// IP
        /// </summary>
        private string txtHost;
        /// <summary>
        /// Port
        /// </summary>
        private string txtPort;
        /// <summary>
        /// 音檔生成路徑
        /// </summary>
        private string DIRPATH = ConfigurationManager.AppSettings["dirpath"];

        delegate void StringArgReturningVoidDelegate(string text);

        public Form1()
        {
            InitializeComponent();
            //取得IP與Port
            txtHost = ConfigurationManager.AppSettings["Host"];
            txtPort = ConfigurationManager.AppSettings["Port"];
        }
        SimpleTcpServer server;

        private void Form1_Load(object sender, EventArgs e)
        {

            //應用程式開啟後，啟動TCP Server
            server = new SimpleTcpServer();
            server.Delimiter = 0x13;//enter
            server.StringEncoder = Encoding.UTF8;
            server.DataReceived += Server_DataReceived;

            try 
            {
                
                System.Net.IPAddress ip = System.Net.IPAddress.Parse(txtHost);
                server.Start(ip, Convert.ToInt32(txtPort));
                txtStatus.Text += "Server starting " + txtHost + ":" + txtPort + " \r\n---------------\r\n";
            }
            catch(Exception eve) //IP或Port無法使用
            {
                txtStatus.Text += "The IP or Port :" + txtHost + ":" + txtPort + " is invalid.\r\n---------------\r\n";
                LogRecord.WriteLog("Started IP error :" + txtHost + ":" + txtPort);
            }
            

        }
        private void checkSlan(string s)
        {

            if (s!="E" && s!="C")
            {
                
                throw new Exception("文頭必須為E或C");
            }
        }

        /// <summary>
        /// 收到client request後，生成音檔
        /// </summary>
        /// <param name="sevder"></param>
        /// <param name="e"></param>
        private void Server_DataReceived(object sevder, SimpleTCP.Message e)
        {
            //Split the Message
            //將Message分割成四個部分,Message型態: language^volume^speed^text
            string phrase = e.MessageString;
            string[] words = phrase.Split('^');

            speech = new SpeechSynthesizer();

            string slan = null;
            string text = "";
            string retext = "";

            //創建音檔時，是否有錯誤產生。
            bool auderror = false;

            try
            {
                //SpeakVolume
                speech.Volume = int.Parse(words[1]);
                //SpeakRate
                speech.Rate = int.Parse(words[2]);
                //Language
                slan = words[0];
                checkSlan(slan);
                //SpeakText
                text = words[3];
                
            }
            catch
            {
                text = "TCP client request error";
                retext = "Request instruction error"; //回傳client指令錯誤
                auderror = true;
            }         

            //若不存在資料夾目錄即新增
            if (!Directory.Exists(DIRPATH))
                Directory.CreateDirectory(DIRPATH);

            //string path = DIRPATH + ConfigurationManager.AppSettings["DefWavFile"]+ DateTime.Now.ToString("CMMddHHmmss")+".wav";
            string path = null;
            string FILENAME = null;


            //Define Language
            if (slan == "E")
            {
                try
                {
                    speech.SelectVoice("Microsoft Zira Desktop"); //中文
                }
                catch
                {
                    LogRecord.WriteLog("Can't find the Zira(english) speaker");
                }
                //指定路徑目的地，檔名為自動生成之時間，英文開頭為E
                FILENAME = DateTime.Now.ToString("EMMddHHmmss");
                path = DIRPATH + FILENAME + ".wav"; 

                retext = string.Format("{0}.wav", FILENAME); //回傳client訊息(wav檔名)

            }
            else if (slan == "C")
            {
                try
                {
                    speech.SelectVoice("Microsoft Hanhan Desktop"); //中文
                }
                catch
                {
                    LogRecord.WriteLog("Can't find the Hanhan(chinese) speaker");
                }
                //指定路徑目的地，檔名為自動生成之時間，中文開頭為C
                FILENAME = DateTime.Now.ToString("CMMddHHmmss");
                path = DIRPATH + FILENAME + ".wav"; 

                retext = string.Format("{0}.wav", FILENAME); //回傳client訊息(wav檔名)

            }

            //當音檔創建有錯誤時，不生成wav檔案
            if (!auderror)
            {
                speech.SetOutputToWaveFile(path);
                speech.Speak(text);
                speech.SetOutputToNull();
            }



            txtStatus.Invoke((MethodInvoker)delegate ()
            {
                //將音檔語音print出
                txtStatus.Text += text+ " \r\n---------------\r\n";
                //寫入log
                if(!auderror)
                    LogRecord.WriteLog("Recieve:"+ e.MessageString +"\r\n  :"+ "Create "+FILENAME+".wav:"+text);
                else
                    LogRecord.WriteLog("Recieve:" + e.MessageString + "\r\n  :" + text);

                e.ReplyLine(retext);
            });
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("是否確定要關閉程式", "關閉程式", MessageBoxButtons.YesNo) == DialogResult.No)
            {
                e.Cancel = true;
            }
        }
    }
}
