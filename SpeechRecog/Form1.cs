using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using SKYPE4COMLib;


namespace SpeechRecog
{
    public partial class Form1 : Form
    {
        SpeechRecognitionEngine engine = new SpeechRecognitionEngine();
        SpeechSynthesizer synthesizer = new SpeechSynthesizer();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            engine.RecognizeAsync(RecognizeMode.Multiple);
            button2.Enabled = true;
            button1.Enabled = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Choices commands = new Choices();
            commands.Add(new String[] { "say hello", "say my name" });
            GrammarBuilder builder = new GrammarBuilder(commands);
            Grammar grammar = new Grammar(builder);
            engine.LoadGrammarAsync(grammar);
            engine.SetInputToDefaultAudioDevice();
            engine.SpeechRecognized += engine_SpeechRecognized;    


        }

        void engine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            switch(e.Result.Text){
                case("say hello"):
                    richTextBox1.Text += "\nsay hello processed";
                    synthesizer.SpeakAsync("Hello! How you doing");
                    break;

                case("say my name"):
                    synthesizer.SpeakAsync("\n Hi Ayush");
                    richTextBox1.Text += "\nsay my name processed";
                    break;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            engine.RecognizeAsyncStop();
            button2.Enabled = false;
            button1.Enabled = true;
        }


    }
}
