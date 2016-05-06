using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using SKYPE4COMLib;
using System.Diagnostics;


namespace SpeechRecog
{
    public partial class Form1 : Form
    {
        Boolean callInProgress = false;
        Boolean navInProgress = false;
        Process myMusic;
        Process myMap;
        SpeechRecognitionEngine engine = new SpeechRecognitionEngine();
        SpeechSynthesizer synthesizer = new SpeechSynthesizer();
        Skype skype = new Skype();
        Call call = null;
        List<User> contacts = new List<User>();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //TimeSpan t = new TimeSpan(8);
            //engine.Recognize(t);
            engine.RecognizeAsync(RecognizeMode.Multiple);
            button2.Enabled = true;
            button1.Enabled = false;
            label7.Text = "Speech Recognition--> On";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            loadContacts();
            loadCities();
            Choices commands = new Choices();
            commands.Add(new String[] {"stop navigation",
                "turn bluetooth on", "turn bluetooth off", "play some music", "stop music",
                 "end call", "cooling off", "cooling on"});
            GrammarBuilder builder = new GrammarBuilder(commands);
            Grammar grammar = new Grammar(builder);
            engine.LoadGrammarAsync(grammar);
            engine.SetInputToDefaultAudioDevice();
            engine.SpeechRecognized += engine_SpeechRecognized;
        }

        void engine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            switch (e.Result.Text){
                case ("cooling on"):
                    richTextBox1.Text += e.Result.Text + "\n";
                    label6.Text = "Cooling ---> On";
                    pictureBox5.ImageLocation = "C:/Bharat/Masters Material/SER 516/SpeechRecognition/co.png";
                    break;
                case ("cooling off"):
                    richTextBox1.Text += e.Result.Text + "\n";
                    label6.Text = "Cooling ---> Off";
                    pictureBox5.ImageLocation = "C:/Bharat/Masters Material/SER 516/SpeechRecognition/cooff.png";
                    break;
                case ("turn bluetooth on"):
                    richTextBox1.Text += "Turn Bluetooth On\n";
                    label2.Text = "Bluetooth ---> On";
                    pictureBox1.ImageLocation = "C:/Bharat/Masters Material/SER 516/SpeechRecognition/bt.png";
                    break;
                case ("turn bluetooth off"):
                    richTextBox1.Text += "turn Bluetooth off\n";
                    label2.Text = "Bluetooth ---> Off";
                    pictureBox1.ImageLocation = "C:/Bharat/Masters Material/SER 516/SpeechRecognition/btoff.png";
                    break;
                case ("play some music"):
                    richTextBox1.Text += e.Result.Text + "\n";
                    label3.Text = "Music ---> On";
                    if (myMusic == null)
                    {
                        myMusic = Process.Start("wmplayer.exe", "C:/Users/Bharat/Downloads/Music/firestone.mp3");
                    }
                    pictureBox2.ImageLocation = "C:/Bharat/Masters Material/SER 516/SpeechRecognition/ms.png";
                    break;
                case ("stop music"):
                    richTextBox1.Text += e.Result.Text + "\n";
                    label3.Text = "Music ---> Off";
                    if (myMusic != null)
                    {
                        myMusic.CloseMainWindow();
                        myMusic.Close();
                        myMusic = null;
                    }
                    pictureBox2.ImageLocation = "C:/Bharat/Masters Material/SER 516/SpeechRecognition/msoff.png";
                    break;
                case ("stop navigation"):
                    richTextBox1.Text += e.Result.Text + "\n";
                    if (myMap != null){
                        myMap.CloseMainWindow();
                        myMap.Close();
                        myMap = null;
                    }
                    navInProgress = false;
                    label4.Text = "Navigation ---> Off";
                    pictureBox3.ImageLocation = "C:/Bharat/Masters Material/SER 516/SpeechRecognition/nvoff.png";
                    break;
                case ("end call"):
                    if (call != null)
                    {
                        call.Finish();
                        call = null;
                        richTextBox1.Text += "End Call\n";
                    }
                    callInProgress = false;
                    label5.Text = "Call ---> Off";
                    pictureBox4.ImageLocation = "C:/Bharat/Masters Material/SER 516/SpeechRecognition/cloff.png";
                    break;
                default:
                    if (!callInProgress && e.Result.Text.Contains("call")){
                        richTextBox1.Text += e.Result.Text + "\n";
                        string name = e.Result.Semantics["person"].Value.ToString();
                        User caller = contacts.Where(user => user.FullName.ToLower().IndexOf(name.ToLower()) > -1).
                                                    Select(user => new User
                                                    {
                                                        Handle = user.Handle
                                                    }).FirstOrDefault();
                        if (caller != null)
                        {
                            call = skype.PlaceCall(caller.Handle);
                            synthesizer.SpeakAsync("Calling " + name);
                            richTextBox1.Text += "Call processed\n";
                            callInProgress = true;
                            label5.Text = "Call ---> On";
                            pictureBox4.ImageLocation = "C:/Bharat/Masters Material/SER 516/SpeechRecognition/cl.png";
                        }
                        else
                        {
                            synthesizer.SpeakAsync("Contact Not Found");
                        }
                        // richTextBox1.Text += "\nDefault in";
                    }
                    else if (!navInProgress && e.Result.Text.Contains("navigate to"))
                    {
                        richTextBox1.Text += e.Result.Text + "\n";
                        string destination = e.Result.Semantics["city"].Value.ToString();
                        myMap = Process.Start("https://www.google.com/maps/dir/Mesa/" + destination);
                        navInProgress = true;
                        label4.Text = "Navigation ---> On";
                        pictureBox3.ImageLocation = "C:/Bharat/Masters Material/SER 516/SpeechRecognition/nv.png";

                    }

                    break; 
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            engine.RecognizeAsyncStop();
            button2.Enabled = false;
            button1.Enabled = true;
            label7.Text = "Speech Recognition--> Off";
        }

        public void loadContacts()
        {
            contacts.Clear();
            List<string> contactsName = new List<string>();
            foreach (User user in skype.Friends)
            {
                if (user.FullName != null && !user.FullName.Equals(""))
                {
                    contacts.Add(user);
                    contactsName.Add(user.FullName.Split(' ')[0]);
                }
                //contactsView.Items.Add(user.FullName);
            }
            Choices contactChoices = new Choices();
            contactChoices.Add(contactsName.ToArray());

            GrammarBuilder builder = new GrammarBuilder();
            builder.Append("call");
            builder.Append(new SemanticResultKey("person", contactChoices));
            Grammar grammar = new Grammar(builder);
            engine.LoadGrammarAsync(grammar);
        }

        public void loadCities()
        {
            Choices citiesChoices = new Choices(new string[] { "Tempe", "Phoenix", "Chicago", "Austin", "Miami", "Dallas" });
            GrammarBuilder builder = new GrammarBuilder();
            builder.Append("navigate to");
            builder.Append(new SemanticResultKey("city", citiesChoices));
            Grammar grammar = new Grammar(builder);
            engine.LoadGrammarAsync(grammar);
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}