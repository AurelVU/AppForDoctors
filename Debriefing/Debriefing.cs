using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Debriefing
{
    public class Question
    {
        public string NameValue { get; set; }
        public string MainQuestion { get; set; }    
    }

    public class TextQuestion : Question
    {
        public TextQuestion(string nameval, string question)
        {
            NameValue = nameval;
            this.MainQuestion = question;
        }
    }

    public class OptionQuestion : Question
    {
        public OptionQuestion(string nameval, string question, List<string> argv)
        {
            NameValue = nameval;
            this.Options = argv;
            this.MainQuestion = question;
        }
        public List<string> Options { get; set; }
    }

    public class CompositeQuestion : Question
    {
        public CompositeQuestion(string nameval, string q, Question question, int? qlock, Question question2)
        {
            NameValue = nameval;
            MainQuestion = q;
            mainquestion = question;
            this.qlock = qlock;
            semiquestion = question2;
        }
        public Question mainquestion { get; set; }
        public int? qlock { get; set; }
        public Question semiquestion { get; set; }
    }

    public class Debriefing
    {
        public enum Type
        {
            Text,
            Option,
            Composite
        } 
        public List<Question> questions = new List<Question>();
        void AddQuestions(Type type, string nameval, params object[] argv)
        {
            if (type == Type.Text)
            {
                questions.Add(new TextQuestion(nameval, argv[0] as string));
            }
            else if (type == Type.Option)
            {
                int n = argv.Length - 1;
                List<string> Options = new List<string>();
                for (int i = 1; i <= n; i++)
                {
                    Options.Add(argv[i] as string);
                }
                questions.Add(new OptionQuestion(nameval, argv[0] as string, Options));
            }
            else if (type == Type.Composite)
            {
                questions.Add(new CompositeQuestion(nameval, argv[0] as string, argv[1] as Question, argv[2] as int?, argv[3] as Question));
            }
        }
    }
}
