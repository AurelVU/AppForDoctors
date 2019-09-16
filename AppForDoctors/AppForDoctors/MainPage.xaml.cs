using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Xamarin.Forms;
using Debriefing;
using System.IO;
using Newtonsoft.Json;
using System.Reflection;
using System.Security.Cryptography;
using System.Net.Sockets;
using BinConverterNamespace;

namespace AppForDoctors
{
    
    public partial class MainPage : ContentPage
    {
        private RSACryptoServiceProvider rsa = new RSACryptoServiceProvider();
        private Aes aes = null;
        private RSAParameters publicKey;
        private RSAParameters privateKey; // Впоследствии удалить 
        private Dictionary<string, string> ansver = new Dictionary<string, string>();
        private byte[] cryptKey;

        private void Des(string json)
        {
            List<Question> questions = new List<Question>();
            json = json.Trim(new char[] { '\n' });
            int indb = json.IndexOf("[");

            int inde = json.LastIndexOf("]");
            string tq = json.Substring(indb + 1, inde - indb);

            List<string> vs = new List<string>();
            int cnt = 0;
            string str = "";
            for (int i = 0; i < tq.Length; i++)
            {
                if (tq[i] == '{') cnt++;
                if (tq[i] == '}') cnt--;
                if (cnt != 0) str += tq[i];
                if (cnt == 0) { str += tq[i]; vs.Add(str); str = ""; while (i != tq.Length - 1 && tq[i + 1] != '{') i++; }
            }
            foreach (string s in vs)
            {
                if (s.IndexOf("Options") != -1)
                {
                    var q = JsonConvert.DeserializeObject<Debriefing.OptionQuestion>(s);

                    questions.Add(q);
                }
                else
                {
                    var q = JsonConvert.DeserializeObject<Debriefing.TextQuestion>(s);

                    questions.Add(q);
                }
            }
            debriefing.questions = questions;
        }

        private Debriefing.Debriefing debriefing = new Debriefing.Debriefing();
        private int i = 0;

        public MainPage()
        {
            InitializeComponent();
            rsa.PersistKeyInCsp = false;
            InitStartScreen();
        }
        private void InitStartScreen()
        {
            i = 0;
            ansver = new Dictionary<string, string>();
            StackLayout stackLayout = new StackLayout();
            Xamarin.Forms.Button button = new Xamarin.Forms.Button();
            button.Text = "Начать тестирование";
            button.Clicked += ButtonClickStart;
            Xamarin.Forms.Button button2 = new Xamarin.Forms.Button();
            button2.Text = "Обновить тест";
            button2.Clicked += ButtonClickRefreshTest;
            stackLayout.Children.Add(button);
            stackLayout.Children.Add(button2);
            Content = stackLayout;
        }

        private void ButtonClickRefreshTest(object sender, EventArgs e)
        {
            StackLayout stackLayout = new StackLayout();
            Xamarin.Forms.Label label = new Xamarin.Forms.Label();
            label.Text = "Обновление теста";
            Xamarin.Forms.Label label2 = new Xamarin.Forms.Label();
            label2.Text = "IP адрес";
            Entry entry = new Entry();
            Xamarin.Forms.Label label3 = new Xamarin.Forms.Label();
            label3.Text = "Порт";
            Entry entry2 = new Entry();
            Xamarin.Forms.Button button = new Xamarin.Forms.Button();
            button.Text = "Обновить";
            button.Clicked += ButtonRefresh;
            stackLayout.Children.Add(label);
            stackLayout.Children.Add(label2);
            stackLayout.Children.Add(entry);
            stackLayout.Children.Add(label3);
            stackLayout.Children.Add(entry2);
            stackLayout.Children.Add(button);
            Content = stackLayout;
        }

        private void ButtonRefresh(object sender, EventArgs e)
        {
            string ip = ((Content as StackLayout).Children[2] as Entry).Text;
            int port = Convert.ToInt32(((Content as StackLayout).Children[4] as Entry).Text);
            
            TcpClient client = null;
            try
            {
                client = new TcpClient(ip, port);
                NetworkStream stream = client.GetStream();

                string message = "Refresh";
                // преобразуем сообщение в массив байтов
                byte[] data = Encoding.Unicode.GetBytes(message);
                // отправка сообщения
                stream.Write(data, 0, data.Length);

                // получаем ответ
                List<byte> fulldata = new List<byte>();
                data = new byte[64]; // буфер для получаемых данных
                
                int bytes = 0;
                do
                {
                    bytes = stream.Read(data, 0, data.Length);
                    for (int i = 0; i < bytes; i++)
                        fulldata.Add(data[i]);
                }
                while (stream.DataAvailable);

                string test = Encoding.Unicode.GetString(fulldata.ToArray());

                File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "TEST.json"), test);

                DisplayAlert("Успех", "Тест успешно обновлен", "ОK");
            }
            catch (Exception ex)
            {
                DisplayAlert("ОШИБКА", ex.Message, "ОK");
                InitStartScreen();
            }
            finally
            {
                client.Close();
                InitStartScreen();
            }

        }

        
        private void ButtonClickStart(object sender, EventArgs e)
        {
            try
            {
                var assembly = IntrospectionExtensions.GetTypeInfo(typeof(AppForDoctors.MainPage)).Assembly;
                Stream stream = new FileStream(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "TEST.json"), FileMode.Open); //assembly.GetManifestResourceStream("AppForDoctors.125.json");
                string text = "";
                using (var reader = new System.IO.StreamReader(stream))
                {
                    text = reader.ReadToEnd();
                }
                Des(text);
                NextQuestion();
            }
            catch
            {
                DisplayAlert("ОШИБКА", "Обновите тест", "ОK");
            }
        }

        private void InitTextQuestion(int num, string question)
        {
            StackLayout stackLayout = new StackLayout();
            Xamarin.Forms.Label label = new Xamarin.Forms.Label();
            label.Text = "Вопрос " + Convert.ToString(num);
            Xamarin.Forms.Label label2 = new Xamarin.Forms.Label();
            label2.Text = question;
            Entry entry = new Entry();
            Xamarin.Forms.Button button = new Xamarin.Forms.Button();
            button.Text = "Ввод";
            button.Clicked += ButtonClickTextQuestion;
            stackLayout.Children.Add(label);
            stackLayout.Children.Add(label2);
            stackLayout.Children.Add(entry);
            stackLayout.Children.Add(button);
            Content = stackLayout;
        }

        private void InitChoiseQuation(int num, string question, List<string> options)
        {
            StackLayout stackLayout = new StackLayout();
            Xamarin.Forms.Label label = new Xamarin.Forms.Label();
            label.Text = "Вопрос " + Convert.ToString(num);
            Xamarin.Forms.Label label2 = new Xamarin.Forms.Label();
            label2.Text = question;
            Picker picker;
            picker = new Picker
            {
                Title = "Варианты"
            };
            foreach (var s in options)
                picker.Items.Add(s);

            picker.SelectedIndexChanged += picker_SelectedIndexChanged;
            stackLayout.Children.Add(label);
            stackLayout.Children.Add(label2);
            stackLayout.Children.Add(picker);
            Content = stackLayout;
        }

        private void NextQuestion()
        {
            if (i < debriefing.questions.Count)
            {
                if (debriefing.questions[i] == null) { i++; NextQuestion(); return; } else
                if (debriefing.questions[i] is Debriefing.TextQuestion)
                    InitTextQuestion(i + 1, debriefing.questions[i].MainQuestion);
                else if (debriefing.questions[i] is Debriefing.OptionQuestion)
                {
                    InitChoiseQuation(i + 1, debriefing.questions[i].MainQuestion, (debriefing.questions[i] as Debriefing.OptionQuestion).Options);
                }
                i++;
            }
            else
            {
                Finish();
            }
        }

        private byte[] ToAes256(string src)
        {
            //Объявляем объект класса AES
            aes = Aes.Create();
            //Генерируем соль
            aes.GenerateIV();
            //Присваиваем ключ. aeskey - переменная (массив байт), сгенерированная методом GenerateKey() класса AES
            aes.GenerateKey();
            byte[] encrypted;
            ICryptoTransform crypt = aes.CreateEncryptor(aes.Key, aes.IV);
            using (MemoryStream ms = new MemoryStream())
            {
                using (CryptoStream cs = new CryptoStream(ms, crypt, CryptoStreamMode.Write))
                {
                    using (StreamWriter sw = new StreamWriter(cs))
                    {
                        sw.Write(src);
                    }
                }
                //Записываем в переменную encrypted зашиврованный поток байтов
                encrypted = ms.ToArray();
            }
            //Возвращаем поток байт + крепим соль
            return encrypted.Concat(aes.IV).ToArray();
        }
        private byte[] Save()
        {
            rsa = new RSACryptoServiceProvider(2048);
            string text = JsonConvert.SerializeObject(ansver);

            rsa.PersistKeyInCsp = false;
            rsa.ImportParameters(publicKey);
            byte[] res = ToAes256(text);
            cryptKey = rsa.Encrypt(aes.Key, true);
            return res;
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

        byte[] Decrypt(byte[] input)
        {
            byte[] decryped;
            rsa.PersistKeyInCsp = false;
            rsa.ImportParameters(privateKey);
            decryped = rsa.Decrypt(input, true);
            return decryped;
        }
        private void Finish()
        {
            
            StackLayout stackLayout = new StackLayout();
            Xamarin.Forms.Label label = new Xamarin.Forms.Label();
            label.Text = "Отправка результатов";
            Xamarin.Forms.Label label2 = new Xamarin.Forms.Label();
            label2.Text = "IP адрес";
            Entry entry = new Entry();
            Xamarin.Forms.Label label3 = new Xamarin.Forms.Label();
            label3.Text = "Порт";
            Entry entry2 = new Entry();
            Xamarin.Forms.Button button = new Xamarin.Forms.Button();
            button.Text = "Ввод";
            button.Clicked += ButtonFinish;
            stackLayout.Children.Add(label);
            stackLayout.Children.Add(label2);
            stackLayout.Children.Add(entry);
            stackLayout.Children.Add(label3);
            stackLayout.Children.Add(entry2);
            stackLayout.Children.Add(button);
            Content = stackLayout;
        }

        private void ButtonFinish(object sender, EventArgs e)
        {
            string ip = ((Content as StackLayout).Children[2] as Entry).Text;
            int port = Convert.ToInt32(((Content as StackLayout).Children[4] as Entry).Text);
                
            TcpClient client = null;
            try
            {
                client = new TcpClient(ip, port);
                NetworkStream stream = client.GetStream();

                string message = "Hello";
                // преобразуем сообщение в массив байтов
                byte[] data = Encoding.Unicode.GetBytes(message);
                // отправка сообщения
                stream.Write(data, 0, data.Length);

                // получаем ответ
                List<byte> fulldata = new List<byte>();
                data = new byte[64]; // буфер для получаемых данных
                
                int bytes = 0;
                do
                {
                    bytes = stream.Read(data, 0, data.Length);
                    for (int i = 0; i < bytes; i++)
                        fulldata.Add(data[i]);
                }
                while (stream.DataAvailable);

                publicKey = (RSAParameters)BinConverter.ByteArrayToObject(fulldata.ToArray());
                byte[] res = Save();
                stream.Write(res, 0, res.Length);
                stream.Write(cryptKey, 0, cryptKey.Length);
                DisplayAlert("Успех", "Данные успешно отправлены", "ОK");
            }
            catch (Exception ex)
            {
                DisplayAlert("ОШИБКА", ex.Message, "ОK");
                InitStartScreen();
            }
            finally
            {
                client.Close();
                InitStartScreen();
            }       
        }

        private void ButtonClickTextQuestion(object obj, EventArgs e)
        {
            var res = ((Content as StackLayout).Children[2] as Entry).Text;
            ansver.Add(debriefing.questions[i - 1].NameValue, res);
            NextQuestion();
        }

        private void picker_SelectedIndexChanged(object obj, EventArgs e)
        {
            var res = ((Content as StackLayout).Children[2] as Picker).Items[((Content as StackLayout).Children[2] as Picker).SelectedIndex];
            try
            {
                ansver.Add(debriefing.questions[i - 1].NameValue, res);
            }
            catch (ArgumentException) { }
            NextQuestion();
        }

    }
}
