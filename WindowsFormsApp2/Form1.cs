using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Net;
using System.Text;
using System.Windows.Forms;
using WMPLib;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        public readonly string ServerAdress = ConfigurationManager.AppSettings["ServerAdress"];
        public readonly string DatabaseName = ConfigurationManager.AppSettings["DatabaseName"];
        public readonly string UsrName = ConfigurationManager.AppSettings["UsrName"];
        public readonly string Pw = ConfigurationManager.AppSettings["Pw"];
        SqlConnection baglanti;
        SqlDataAdapter sqlAdapter;
        SqlCommand komut;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            lblAlarm.Visible = false;
            timer1.Start();
        }

        WMPLib.WindowsMediaPlayer Player;

        private void PlayFile(String url)
        {
            Player = new WMPLib.WindowsMediaPlayer();
            Player.PlayStateChange +=
                new WMPLib._WMPOCXEvents_PlayStateChangeEventHandler(Player_PlayStateChange);
            Player.MediaError +=
                new WMPLib._WMPOCXEvents_MediaErrorEventHandler(Player_MediaError);
            Player.URL = url;
            Player.settings.autoStart = true;
            Player.settings.setMode("loop", true);
            Player.controls.play(); 
        }

        private void Player_PlayStateChange(int NewState)
        {
            if ((WMPLib.WMPPlayState)NewState == WMPLib.WMPPlayState.wmppsStopped)
            {
                this.Close();
            }
        }

        private void Player_MediaError(object pMediaObject)
        {
            MessageBox.Show("Cannot play media file.");
            this.Close();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            var url = "https://pos.tlscontact.com/ist_en/eshop/services/";

            var wc = new WebClient();
            var arr = wc.DownloadData(url);
            var txt = Encoding.UTF8.GetString(arr);

            var ret = txt.Contains("£250.00");

            if (ret)
            {
                lblAlarm.Visible = true;
                baglanti = new SqlConnection("Server=" + ServerAdress + ";Database=" + DatabaseName + ";User Id=" + UsrName + ";Password=" + Pw + ";");
                baglanti.Open();
                komut = new SqlCommand($"USE [StudentApp] GO DECLARE	@return_value int EXEC	@return_value = [dbo].[AutoMailTLS] SELECT	'Return Value' = @return_value GO", baglanti);
                komut.ExecuteScalar();
                baglanti.Close();
                PlayFile(@"Alarm01.wav");
            }
        }
    }
}
