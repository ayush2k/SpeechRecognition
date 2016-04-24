﻿using System;
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
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            loadContacts();
            loadCities();
            Choices commands = new Choices();
            commands.Add(new String[] {"turn navigation off",
                "turn tooth on", "tooth off", "play some music", "stop music",
                 "end call"});
            GrammarBuilder builder = new GrammarBuilder(commands);
            Grammar grammar = new Grammar(builder);
            engine.LoadGrammarAsync(grammar);
            engine.SetInputToDefaultAudioDevice();
            engine.SpeechRecognized += engine_SpeechRecognized;    


        }

        void engine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            switch (e.Result.Text){
                case ("turn tooth on"):
                    richTextBox1.Text += e.Result.Text + "\n";
                    label2.Text = "Bluetooth ---> On";
                    pictureBox1.ImageLocation = "C:/Bharat/Masters Material/SER 516/SpeechRecognition/bt.png";
                    break;
                case ("tooth off"):
                    richTextBox1.Text += e.Result.Text + "\n";
                    label2.Text = "Bluetooth ---> Off";
                    pictureBox1.ImageLocation = "C:/Bharat/Masters Material/SER 516/SpeechRecognition/btoff.png";
                    break;
                case ("play some music"):
                    richTextBox1.Text += e.Result.Text + "\n";
                    label3.Text = "Music ---> On";
                    myMusic = Process.Start("wmplayer.exe", "C:/Users/Bharat/Downloads/Music/firestone.mp3");
                    pictureBox2.ImageLocation = "C:/Bharat/Masters Material/SER 516/SpeechRecognition/ms.png";
                    break;
                case ("stop music"):
                    richTextBox1.Text += e.Result.Text + "\n";
                    label3.Text = "Music ---> Off";
                    if (myMusic != null)
                    {
                        myMusic.CloseMainWindow();
                        myMusic.Close();
                    }
                    pictureBox2.ImageLocation = "C:/Bharat/Masters Material/SER 516/SpeechRecognition/msoff.png";
                    break;
                case ("turn navigation off"):
                    richTextBox1.Text += e.Result.Text + "\n";
                    if (myMap != null){
                        myMap.CloseMainWindow();
                        myMap.Close();
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
                    }
                    callInProgress = false;
                    label4.Text = "Call ---> Off";
                    pictureBox3.ImageLocation = "C:/Bharat/Masters Material/SER 516/SpeechRecognition/cloff.png";
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
                            synthesizer.SpeakAsync("\n Calling " + name);
                            richTextBox1.Text += "\nCall processed";
                            callInProgress = true;
                            label5.Text = "Call ---> On";
                            pictureBox4.ImageLocation = "C:/Bharat/Masters Material/SER 516/SpeechRecognition/cl.png";
                        }
                        else
                        {
                            synthesizer.SpeakAsync("Contact Not Found");
                        }
                        richTextBox1.Text += "\nDefault in";
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