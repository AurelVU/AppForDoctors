using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Debriefing;
using System.IO;
using Newtonsoft.Json;


namespace GenerateFile
{
    public partial class Form1 : Form
    {


        List<Debriefing.Question> Questions = new List<Question>();

        public Form1()
        {
            InitializeComponent();
            comboBox1.Items.Add("Текст");
            comboBox1.Items.Add("Выбор");
            comboBox1.Items.Add("Комбинированный");
            comboBox1.Tag = panel1;
            comboBox2.Items.Add(1);
            comboBox2.SelectedIndex = 0;
            Text = "Генерация опроса";
        }

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
            Questions = questions;
        }


        void CreateOptions(Panel panel)
        {
            Label label = new Label() { Text = "Название параметра: " };
            label.Width += 30;
            Label label3 = new Label { Text = "Вопрос: ", Top = label.Size.Height + 15 };
            TextBox textBox1 = new TextBox() { Left = label.Size.Width + 15 };
            TextBox textBox2 = new TextBox() { Left = label3.Size.Width + 15, Top = label3.Top };
            panel.Top = label3.Top + label3.Size.Height;
            panel.Controls.Add(label);
            panel.Controls.Add(label3);
            panel.Controls.Add(textBox1);
            panel.Controls.Add(textBox2);
            Panel panel2 = new Panel();
            Button buttonplus = new Button() { Text = "Добавить вариант", Tag = panel2 };
            buttonplus.Click += CliclAdd;
            buttonplus.Top = label3.Top + label3.Size.Height + 15; 
            Button buttondiff = new Button() { Text = "Убрать вариант", Tag = panel2 };
            buttondiff.Click += CliclDiff;
            buttondiff.Top = label3.Top + label3.Size.Height + 15;
            buttondiff.Left = buttondiff.Size.Width + 15;
            panel2.Top = buttondiff.Top + buttondiff.Size.Height + 15;
            panel.Controls.Add(buttonplus);
            panel.Controls.Add(buttondiff);
            panel.Controls.Add(panel2);
        }


        void CreateText(Panel panel)
        {
            panel.Controls.Clear();
            Label label = new Label() { Text = "Название параметра: " };
            label.Width += 30;
            Label label3 = new Label { Text = "Вопрос: ", Top = label.Size.Height + 15 };
            TextBox textBox1 = new TextBox() { Left = label.Size.Width + 15};
            TextBox textBox2 = new TextBox() { Left = label3.Size.Width + 15, Top = label3.Top };
            panel.Controls.Add(label);
            panel.Controls.Add(label3);
            panel.Controls.Add(textBox1);
            panel.Controls.Add(textBox2);
        }

        void CreateMulti(Panel panel)
        {
            Panel panel2 = new Panel();
            Panel panel3 = new Panel();
            ComboBox comboBox = new ComboBox();
            panel2.Top = comboBox.Height + 10;
            comboBox.Items.Add("Текст");
            comboBox.Items.Add("Выбор");
            comboBox.Items.Add("Комбинированный");
            comboBox.Tag = panel2;
            comboBox.SelectedIndexChanged += comboBox1_SelectedIndexChanged;

            Label label = new Label() { Text = "Привязка к параметру №", Top = panel2.Top + panel2.Height + 10 };
            TextBox textBox = new TextBox() { Top = panel2.Top + panel2.Height + 10, Left = label.Width + 10 };

            ComboBox comboBox2 = new ComboBox() { Top = label.Top + label.Height + 10 };
            comboBox2.Items.Add("Текст");
            comboBox2.Items.Add("Выбор");
            comboBox2.Items.Add("Комбинированный");
            comboBox2.Tag = panel3;
            comboBox2.SelectedIndexChanged += comboBox1_SelectedIndexChanged;

            panel.Controls.Add(panel2);
            panel.Controls.Add(label);
            panel.Controls.Add(textBox);
            panel.Controls.Add(panel3);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Panel panel2 = (sender as ComboBox).Tag as Panel;
            panel2.Controls.Clear();
            if (comboBox1.SelectedIndex == 0)
            {
                CreateText(panel2);
            }
            else if (comboBox1.SelectedIndex == 1)
            {
                CreateOptions(panel2);
            }
            else if (comboBox1.SelectedIndex == 2)
            {
                CreateMulti(panel2);
            }
        }

        private void CliclAdd(object obj, EventArgs e)
        {
            Panel panel = (obj as Button).Tag as Panel;
            TextBox textBox = new TextBox();
            if (panel.Controls.Count != 0)
                textBox.Top = panel.Controls[panel.Controls.Count - 1].Top + panel.Controls[panel.Controls.Count - 1].Size.Height + 5;
            panel.Controls.Add(textBox);
            panel.Height = 600;
        }

        private void CliclDiff(object obj, EventArgs e)
        {
            Panel panel = (obj as Button).Tag as Panel;
            if (panel.Controls.Count != 0)
                panel.Controls.RemoveAt(panel.Controls.Count - 1);
        }

        Question GetTextQuestion(Panel panel)
        {
            return new TextQuestion(panel.Controls[2].Text, panel.Controls[3].Text);
        }

        Question GetOptionsQuestion(Panel panel)
        {
            List<string> argv = new List<string>();
            var panel2 = panel.Controls[panel.Controls.Count - 1] as Panel;
            foreach (var s in panel2.Controls)
                argv.Add((s as TextBox).Text);
            Question question = new OptionQuestion(panel1.Controls[2].Text, panel1.Controls[3].Text, argv);
            return question;
        }

        /*Question GetMultyQuestion(Panel panel)
        {
            Question question = new Debriefing.CompositeQuestion(panel.Controls[2].Text, panel.Controls[3].Text, )
        }*/

        



        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                Questions.Add(GetTextQuestion(panel1));
            }
            else if (comboBox1.SelectedIndex == 1)
            {
                Questions.Add(GetOptionsQuestion(panel1));
            }
            else if (comboBox1.SelectedIndex == 2) { }
            comboBox2.Items.Add(Questions.Count + 1);
            comboBox2.SelectedIndex = comboBox2.Items.Count - 1;
            panel1.Controls.Clear();
            comboBox1_SelectedIndexChanged(comboBox1, new EventArgs());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Debriefing.Debriefing debriefing = new Debriefing.Debriefing();
            debriefing.questions = Questions;
            string res = JsonConvert.SerializeObject(debriefing);
            saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            saveFileDialog1.Filter = "JSON файл (*.json)|*.json";
            saveFileDialog1.FileName = "";
            var s = saveFileDialog1.ShowDialog();
            using (StreamWriter sw = new StreamWriter(saveFileDialog1.FileName, false, System.Text.Encoding.Unicode))
            {
                sw.WriteLine(res);
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            openFileDialog1.Filter = "JSON файл (*.json)|*.json";
            openFileDialog1.FileName = "";
            var s = openFileDialog1.ShowDialog();
            StreamReader sw = new StreamReader(openFileDialog1.FileName);
            Des(sw.ReadToEnd());
        }

        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
