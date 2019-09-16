using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.Net.Sockets;
using System.Threading;

namespace Server
{
    public partial class Form1 : Form
    {
        List<Thread> threads = new List<Thread>();
        public IPAddress IP;
        public int port = 80;
        TcpListener listener;
        string test;
        int countConnect = 0;
        string workingExcel = "";

        Thread thread;
        public Form1()
        {
            InitializeComponent();
            thread = new Thread(new ThreadStart(start));
            threads.Add(thread);
            thread.IsBackground = true;
            textBox1.Text = port.ToString();
            textBox2.Text = Convert.ToString(countConnect);
            Button1_Click(null, null);
        }

        void RefreshTest(ClientObject clientObject)
        {
            clientObject.test = test;
        }
        void start() {
            try
            {
                listener = new TcpListener(IPAddress.Any, port);
                listener.Start();
                
                while (true)
                {
                    TcpClient client = listener.AcceptTcpClient();
                    ClientObject clientObject = new ClientObject(client, test, workingExcel);
                    clientObject.NewTest += RefreshTest;
                    // создаем новый поток для обслуживания нового клиента
                    countConnect++;
                    Thread clientThread = new Thread(new ThreadStart(clientObject.Process));
                    threads.Add(clientThread);
                    clientThread.IsBackground = true;
                    clientThread.Start();
                }
            }
            catch (Exception ex)
            {
                label3.Text = ex.Message;
            }
            finally
            {
                if (listener != null)
                    listener.Stop();
                countConnect--;
            }
        }


        private void Button1_Click(object sender, EventArgs e)
        {
            thread.Abort();
            try
            {
                label1.Text = new WebClient().DownloadString("http://ident.me/");
                IP = IPAddress.Parse(label1.Text);
            } catch { label1.Text = "Невозможно определить ваш ip адрес"; }
            try
            {
                port = Convert.ToInt32(label2.Text);
            }
            catch
            {
            }
            thread = new Thread(new ThreadStart(start));
            thread.Start();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            openFileDialog1.Filter = "JSON файл (*.json)|*.json";
            openFileDialog1.FileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                label4.Text = openFileDialog1.FileName;
                test = System.IO.File.ReadAllText(openFileDialog1.FileName);
            }
        }

        private void Form1_Paint(object sender, PaintEventArgs e)
        {
            textBox2.Text = Convert.ToString(countConnect);
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            openFileDialog1.Filter = "Excel файл (*.xlsx)|*.xlsx";
            openFileDialog1.FileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                label5.Text = openFileDialog1.FileName;
                workingExcel = openFileDialog1.FileName;
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            foreach (Thread thread in threads)
                if (thread != null) 
                    if (thread.IsAlive)
                        thread.Abort();
        }
    }
}
