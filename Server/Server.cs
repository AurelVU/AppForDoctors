using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using BinConverterNamespace;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace Server
{
    public class ClientObject
    {
        public TcpClient client;
        public string test;
        public delegate void GetTest(ClientObject clientObject);
        public event GetTest NewTest;
        public string workingExcelFile;
        public ClientObject(TcpClient tcpClient, string test, string workingExcelFile)
        {
            client = tcpClient;
            this.test = test;
            this.workingExcelFile = workingExcelFile;
        }

        private string FromAes256(byte[] shifr, byte[] key)
        {
            byte[] bytesIv = new byte[16];
            byte[] mess = new byte[shifr.Length - 16];
            //Списываем соль
            for (int i = shifr.Length - 16, j = 0; i < shifr.Length; i++, j++)
                bytesIv[j] = shifr[i];
            //Списываем оставшуюся часть сообщения
            for (int i = 0; i < shifr.Length - 16; i++)
                mess[i] = shifr[i];
            //Объект класса Aes
            Aes aes = Aes.Create();
            //Задаем тот же ключ, что и для шифрования
            aes.Key = key;
            //Задаем соль
            aes.IV = bytesIv;
            //Строковая переменная для результата
            string text = "";
            byte[] data = mess;
            ICryptoTransform crypt = aes.CreateDecryptor(aes.Key, aes.IV);
            using (MemoryStream ms = new MemoryStream(data))
            {
                using (CryptoStream cs = new CryptoStream(ms, crypt, CryptoStreamMode.Read))
                {
                    using (StreamReader sr = new StreamReader(cs))
                    {
                        //Результат записываем в переменную text в вие исходной строки
                        text = sr.ReadToEnd();
                    }
                }
            }
            return text;
        }

        public delegate void MethodContainerConnection();
        public event MethodContainerConnection messageAboutConnection;

        public delegate void MethodContainerDisConnection();
        public event MethodContainerDisConnection messageAboutDisConnection;
        public void Process()
        {
            NetworkStream stream = null;
            try
            {

                stream = client.GetStream();
                byte[] data = new byte[64]; // буфер для получаемых данных
                while (true)
                {
                    // получаем сообщение
                    StringBuilder builder = new StringBuilder();
                    int bytes = 0;
                    do
                    {
                        bytes = stream.Read(data, 0, data.Length);
                        builder.Append(Encoding.Unicode.GetString(data, 0, bytes));
                    }
                    while (stream.DataAvailable);
                    if (builder.ToString() == "Hello")
                    {
                        GetData(stream);
                    }
                    else if (builder.ToString() == "Refresh")
                    {
                        Refresh(stream);
                    }
                }
            }
            catch (Exception ex)
            {
                if (messageAboutDisConnection != null)
                    messageAboutDisConnection();
            }
            finally
            {
                if (stream != null)
                    stream.Close();
                if (client != null)
                    client.Close();
                if (messageAboutDisConnection != null)
                    messageAboutDisConnection();
            }
        }

        private void Refresh(NetworkStream stream)
        {
            if (NewTest != null)
                NewTest(this);
            byte[] data = Encoding.Unicode.GetBytes(this.test);
            stream.Write(data, 0, data.Length);
        }

        private void GetData(NetworkStream stream)
        {
            RSACryptoServiceProvider rsa = new RSACryptoServiceProvider();
            Aes aes = null;
            RSAParameters publicKey;
            RSAParameters privateKey;

            rsa.PersistKeyInCsp = false;
            publicKey = rsa.ExportParameters(false);
            privateKey = rsa.ExportParameters(true);


            byte[] data = BinConverter.ObjectToByteArray(publicKey);
            stream.Write(data, 0, data.Length);

            List<byte> fulldata = new List<byte>();
            int bytes = 0;
            do
            {
                bytes = stream.Read(data, 0, data.Length);
                //builder.Append(Encoding.Unicode.GetString(data, 0, bytes));
                for (int i = 0; i < bytes; i++)
                    fulldata.Add(data[i]);
            }
            while (stream.DataAvailable);

            List<byte> fullkey = new List<byte>();
            bytes = 0;
            do
            {
                bytes = stream.Read(data, 0, data.Length);
                //builder.Append(Encoding.Unicode.GetString(data, 0, bytes));
                for (int i = 0; i < bytes; i++)
                    fullkey.Add(data[i]);
            }
            while (stream.DataAvailable);

            byte[] key;
            rsa.PersistKeyInCsp = false;
            rsa.ImportParameters(privateKey);
            key = rsa.Decrypt(fullkey.ToArray(), true);

            string result = FromAes256(fulldata.ToArray(), key);
            Dictionary<string, string> ansver = JsonConvert.DeserializeObject<Dictionary<string, string>>(result);



            //рабоата с Excel
            Excel.Range Rng;
            Excel.Workbook xlWB;
            Excel.Worksheet xlSht;
            int iLastRow, iLastCol;

            Excel.Application xlApp = new Excel.Application(); //создаём приложение Excel
            xlWB = xlApp.Workbooks.Open(workingExcelFile); //открываем наш файл           
            xlSht = xlWB.Worksheets["Лист1"]; //или так xlSht = xlWB.ActiveSheet //активный лист

            iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row; //последняя заполненная строка в столбце А

            if (iLastRow == 1)
            {
                int j = 1;
                foreach (var col in ansver)
                {
                    xlSht.Cells[1, j] = col.Key;
                    xlSht.Cells[2, j] = col.Value;
                    j++;
                }
            }
            else
            {
                int j = 1;
                foreach (var col in ansver)
                {
                    xlSht.Cells[iLastRow + 1, j] = col.Value;
                    j++;
                }
            }
            
            //закрытие Excel
            xlWB.Close(true); //сохраняем и закрываем файл
            xlApp.Quit();
        }
    }
}
