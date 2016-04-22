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
using System.Diagnostics;
using System.Threading;
using System.Collections;


namespace SpeechRecog
{


    public partial class Form1 : Form
    {
        Process myMusic;
        Process myMap;
        SpeechRecognitionEngine engine = new SpeechRecognitionEngine();
        SpeechSynthesizer synthesizer = new SpeechSynthesizer();
        Skype skype = new Skype();
        Call call = null;
        Choices commands = new Choices();
        List<User> contacts = new List<User>();

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
            commands.Add(new String[] { "say hello", "say my name", "navigate to", "tempe", "turn navigation off",
                "turn blue tooth on", "blue tooth off", "play some music", "stop music",
                "turn navigation on", "call someone", "end call"});
            GrammarBuilder builder = new GrammarBuilder(commands);
            Grammar grammar = new Grammar(builder);
            engine.LoadGrammarAsync(grammar);
            engine.SetInputToDefaultAudioDevice();
            engine.SpeechRecognized += engine_SpeechRecognized;    


        }

        void engine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            richTextBox1.Text += e.Result.Text;
            switch (e.Result.Text){
                case ("turn blue tooth on"):
                    label2.Text = "Bluetooth ---> On";
                    break;
                case ("blue tooth off"):
                    label2.Text = "Bluetooth ---> Off";
                    break;
                case ("play some music"):
                    label3.Text = "Music ---> On";
                    myMusic = Process.Start("wmplayer.exe", "C:/Users/Bharat/Downloads/Music/firestone.mp3");
                    break;
                case ("stop music"):
                    label3.Text = "Music ---> Off";
                    if (myMusic != null)
                    {
                        myMusic.CloseMainWindow();
                        myMusic.Close();
                    }
                    break;
                case ("turn navigation off"):
                    label3.Text = "Music ---> Off";
                    if (myMap != null){
                        myMap.CloseMainWindow();
                        myMap.Close();
                    }
                    break;
                case ("navigate to"):
                    richTextBox1.Text += "\nNavigation On";
                    myMap = Process.Start("https://www.google.com/maps/dir/Tempe/San+Diego");
                    break;
                case ("say hello"):
                    richTextBox1.Text += "\nsay hello processed";
                    synthesizer.SpeakAsync("Hello! How you doing");
                    break;
                case("say my name"):
                    synthesizer.SpeakAsync("\n Hi Ayush");
                    richTextBox1.Text += "\nsay my name processed";
                    break;
                case ("call someone"):
                    string name = "echo123";
                    User caller = contacts.Where(user => user.FullName.ToLower().IndexOf(name.ToLower()) > -1).
                                                Select(user => new User
                                                {
                                                    Handle = user.Handle
                                                }).FirstOrDefault();
                    if (caller != null)
                    {
                        skype.PlaceCall(caller.Handle);
                        synthesizer.SpeakAsync("\n Calling " + name);
                        richTextBox1.Text += "\nCall processed";
                    }
                    else
                    {
                        synthesizer.SpeakAsync("Contact Not Found");
                    }
                    break;
                case ("end call"):
                    if (call != null)
                    {
                        call.Finish();
                        call = null;
                    }
                    break;
                default:
                    richTextBox1.Text += "\nDefault in";
                    break;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            engine.RecognizeAsyncStop();
            button2.Enabled = false;
            button1.Enabled = true;
        }

        public void loadContacts()
        {
            contacts.Clear();
            foreach (User user in skype.Friends)
            {
                contacts.Add(user);
                //contactsView.Items.Add(user.FullName);
            }
        }
    }
}