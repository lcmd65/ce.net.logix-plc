using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Logix;
using System.Timers;

namespace TWINCAT_ADS_Client
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }
        private string plc_Address;
        private Controller myPLC = new Controller();

        //Recipe - Cleaning 1
        private Tag Cleaning1_Seq1_cleaning_medium = new Tag("recipe_CleanMachine_edit.Step1.W1_medium");
        private Tag Cleaning1_Seq1_cleaning_top = new Tag("recipe_CleanMachine_edit.Step1.W2_pressure_top");
        private Tag Cleaning1_Seq1_cleaning_bottom = new Tag("recipe_CleanMachine_edit.Step1.W3_buttom");
        private Tag Cleaning1_Seq1_Supply = new Tag("recipe_CleanCirculation_edit.Step1.W4_valves_top");
        private Tag Cleaning1_Seq1_Backflow = new Tag("recipe_CleanCirculation_edit.Step1.W5_valves_buttom");
        private Tag Cleaning1_Seq1_duration = new Tag("recipe_CleanMachine_edit.Step1.W6_time_step");
        private Tag Cleaning1_Seq2_cleaning_medium = new Tag("recipe_CleanMachine_edit.Step2.W1_medium");
        private Tag Cleaning1_Seq2_cleaning_top = new Tag("recipe_CleanMachine_edit.Step2.W2_pressure_top");
        private Tag Cleaning1_Seq2_cleaning_bottom = new Tag("recipe_CleanMachine_edit.Step2.W3_buttom");
        private Tag Cleaning1_Seq2_Supply = new Tag("recipe_CleanCirculation_edit.Step2.W4_valves_top");
        private Tag Cleaning1_Seq2_Backflow = new Tag("recipe_CleanCirculation_edit.Step2.W5_valves_buttom");
        private Tag Cleaning1_Seq2_duration = new Tag("recipe_CleanMachine_edit.Step2.W6_time_step");
        private Tag Cleaning1_Seq3_cleaning_medium = new Tag("recipe_CleanMachine_edit.Step3.W1_medium");
        private Tag Cleaning1_Seq3_cleaning_top = new Tag("recipe_CleanMachine_edit.Step3.W2_pressure_top");
        private Tag Cleaning1_Seq3_cleaning_bottom = new Tag("recipe_CleanMachine_edit.Step3.W3_buttom");
        private Tag Cleaning1_Seq3_Supply = new Tag("recipe_CleanCirculation_edit.Step3.W4_valves_top");
        private Tag Cleaning1_Seq3_Backflow = new Tag("recipe_CleanCirculation_edit.Step3.W5_valves_buttom");
        private Tag Cleaning1_Seq3_duration = new Tag("recipe_CleanMachine_edit.Step3.W6_time_step");
        private Tag Cleaning1_Seq4_cleaning_medium = new Tag("recipe_CleanMachine_edit.Step4.W1_medium");
        private Tag Cleaning1_Seq4_cleaning_top = new Tag("recipe_CleanMachine_edit.Step4.W2_pressure_top");
        private Tag Cleaning1_Seq4_cleaning_bottom = new Tag("recipe_CleanMachine_edit.Step4.W3_buttom");
        private Tag Cleaning1_Seq4_Supply = new Tag("recipe_CleanCirculation_edit.Step4.W4_valves_top");
        private Tag Cleaning1_Seq4_Backflow = new Tag("recipe_CleanCirculation_edit.Step4.W5_valves_buttom");
        private Tag Cleaning1_Seq4_duration = new Tag("recipe_CleanMachine_edit.Step4.W6_time_step");
        private Tag Cleaning1_Seq5_cleaning_medium = new Tag("recipe_CleanMachine_edit.Step5.W1_medium");
        private Tag Cleaning1_Seq5_cleaning_top = new Tag("recipe_CleanMachine_edit.Step5.W2_pressure_top");
        private Tag Cleaning1_Seq5_cleaning_bottom = new Tag("recipe_CleanMachine_edit.Step5.W3_buttom");
        private Tag Cleaning1_Seq5_Supply = new Tag("recipe_CleanCirculation_edit.Step5.W4_valves_top");
        private Tag Cleaning1_Seq5_Backflow = new Tag("recipe_CleanCirculation_edit.Step5.W5_valves_buttom");
        private Tag Cleaning1_Seq5_duration = new Tag("recipe_CleanMachine_edit.Step5.W6_time_step");
        private Tag Cleaning1_Seq6_cleaning_medium = new Tag("recipe_CleanMachine_edit.Step6.W1_medium");
        private Tag Cleaning1_Seq6_cleaning_top = new Tag("recipe_CleanMachine_edit.Step6.W2_pressure_top");
        private Tag Cleaning1_Seq6_cleaning_bottom = new Tag("recipe_CleanMachine_edit.Step6.W3_buttom");
        private Tag Cleaning1_Seq6_Supply = new Tag("recipe_CleanCirculation_edit.Step6.W4_valves_top");
        private Tag Cleaning1_Seq6_Backflow = new Tag("recipe_CleanCirculation_edit.Step6.W5_valves_buttom");
        private Tag Cleaning1_Seq6_duration = new Tag("recipe_CleanMachine_edit.Step6.W6_time_step");
        private Tag Cleaning1_Seq7_cleaning_medium = new Tag("recipe_CleanMachine_edit.Step7.W1_medium");
        private Tag Cleaning1_Seq7_cleaning_top = new Tag("recipe_CleanMachine_edit.Step7.W2_pressure_top");
        private Tag Cleaning1_Seq7_cleaning_bottom = new Tag("recipe_CleanMachine_edit.Step7.W3_buttom");
        private Tag Cleaning1_Seq7_Supply = new Tag("recipe_CleanCirculation_edit.Step7.W4_valves_top");
        private Tag Cleaning1_Seq7_Backflow = new Tag("recipe_CleanCirculation_edit.Step7.W5_valves_buttom");
        private Tag Cleaning1_Seq7_duration = new Tag("recipe_CleanMachine_edit.Step7.W6_time_step");
        private Tag Cleaning1_Seq8_cleaning_medium = new Tag("recipe_CleanMachine_edit.Step8.W1_medium");
        private Tag Cleaning1_Seq8_cleaning_top = new Tag("recipe_CleanMachine_edit.Step8.W2_pressure_top");
        private Tag Cleaning1_Seq8_cleaning_bottom = new Tag("recipe_CleanMachine_edit.Step8.W3_buttom");
        private Tag Cleaning1_Seq8_Supply = new Tag("recipe_CleanCirculation_edit.Step8.W4_valves_top");
        private Tag Cleaning1_Seq8_Backflow = new Tag("recipe_CleanCirculation_edit.Step8.W5_valves_buttom");
        private Tag Cleaning1_Seq8_duration = new Tag("recipe_CleanMachine_edit.Step8.W6_time_step");

        //Recipe - Cleaning 2
        private Tag Cleaning2_Seq1_cleaning_medium = new Tag("recipe_CleanCirculation_machine.Step1.W1_medium");
        private Tag Cleaning2_Seq1_cleaning_top = new Tag("recipe_CleanCirculation_machine.Step1.W2_pressure_top");
        private Tag Cleaning2_Seq1_cleaning_bottom = new Tag("recipe_CleanCirculation_machine.Step1.W3_buttom");
        private Tag Cleaning2_Seq1_Supply = new Tag("recipe_CleanCirculation_machine.Step1.W4_valves_top");
        private Tag Cleaning2_Seq1_Backflow = new Tag("recipe_CleanCirculation_machine.Step1.W5_valves_buttom");
        private Tag Cleaning2_Seq1_duration = new Tag("recipe_CleanCirculation_machine.Step1.W6_time_step");
        private Tag Cleaning2_Seq2_cleaning_medium = new Tag("recipe_CleanCirculation_machine.Step2.W1_medium");
        private Tag Cleaning2_Seq2_cleaning_top = new Tag("recipe_CleanCirculation_machine.Step2.W2_pressure_top");
        private Tag Cleaning2_Seq2_cleaning_bottom = new Tag("recipe_CleanCirculation_machine.Step2.W3_buttom");
        private Tag Cleaning2_Seq2_Supply = new Tag("recipe_CleanCirculation_machine.Step2.W4_valves_top");
        private Tag Cleaning2_Seq2_Backflow = new Tag("recipe_CleanCirculation_machine.Step2.W5_valves_buttom");
        private Tag Cleaning2_Seq2_duration = new Tag("recipe_CleanCirculation_machine.Step2.W6_time_step");
        private Tag Cleaning2_Seq3_cleaning_medium = new Tag("recipe_CleanCirculation_machine.Step3.W1_medium");
        private Tag Cleaning2_Seq3_cleaning_top = new Tag("recipe_CleanCirculation_machine.Step3.W2_pressure_top");
        private Tag Cleaning2_Seq3_cleaning_bottom = new Tag("recipe_CleanCirculation_machine.Step3.W3_buttom");
        private Tag Cleaning2_Seq3_Supply = new Tag("recipe_CleanCirculation_machine.Step3.W4_valves_top");
        private Tag Cleaning2_Seq3_Backflow = new Tag("recipe_CleanCirculation_machine.Step3.W5_valves_buttom");
        private Tag Cleaning2_Seq3_duration = new Tag("recipe_CleanCirculation_machine.Step3.W6_time_step");
        private Tag Cleaning2_Seq4_cleaning_medium = new Tag("recipe_CleanCirculation_machine.Step4.W1_medium");
        private Tag Cleaning2_Seq4_cleaning_top = new Tag("recipe_CleanCirculation_machine.Step4.W2_pressure_top");
        private Tag Cleaning2_Seq4_cleaning_bottom = new Tag("recipe_CleanCirculation_machine.Step4.W3_buttom");
        private Tag Cleaning2_Seq4_Supply = new Tag("recipe_CleanCirculation_machine.Step4.W4_valves_top");
        private Tag Cleaning2_Seq4_Backflow = new Tag("recipe_CleanCirculation_machine.Step4.W5_valves_buttom");
        private Tag Cleaning2_Seq4_duration = new Tag("recipe_CleanCirculation_machine.Step4.W6_time_step");
        private Tag Cleaning2_Seq5_cleaning_medium = new Tag("recipe_CleanCirculation_machine.Step5.W1_medium");
        private Tag Cleaning2_Seq5_cleaning_top = new Tag("recipe_CleanCirculation_machine.Step5.W2_pressure_top");
        private Tag Cleaning2_Seq5_cleaning_bottom = new Tag("recipe_CleanCirculation_machine.Step5.W3_buttom");
        private Tag Cleaning2_Seq5_Supply = new Tag("recipe_CleanCirculation_machine.Step5.W4_valves_top");
        private Tag Cleaning2_Seq5_Backflow = new Tag("recipe_CleanCirculation_machine.Step5.W5_valves_buttom");
        private Tag Cleaning2_Seq5_duration = new Tag("recipe_CleanCirculation_machine.Step5.W6_time_step");
        private Tag Cleaning2_Seq6_cleaning_medium = new Tag("recipe_CleanCirculation_machine.Step6.W1_medium");
        private Tag Cleaning2_Seq6_cleaning_top = new Tag("recipe_CleanCirculation_machine.Step6.W2_pressure_top");
        private Tag Cleaning2_Seq6_cleaning_bottom = new Tag("recipe_CleanCirculation_machine.Step6.W3_buttom");
        private Tag Cleaning2_Seq6_Supply = new Tag("recipe_CleanCirculation_machine.Step6.W4_valves_top");
        private Tag Cleaning2_Seq6_Backflow = new Tag("recipe_CleanCirculation_machine.Step6.W5_valves_buttom");
        private Tag Cleaning2_Seq6_duration = new Tag("recipe_CleanCirculation_machine.Step6.W6_time_step");
        private Tag Cleaning2_Seq7_cleaning_medium = new Tag("recipe_CleanCirculation_machine.Step7.W1_medium");
        private Tag Cleaning2_Seq7_cleaning_top = new Tag("recipe_CleanCirculation_machine.Step7.W2_pressure_top");
        private Tag Cleaning2_Seq7_cleaning_bottom = new Tag("recipe_CleanCirculation_machine.Step7.W3_buttom");
        private Tag Cleaning2_Seq7_Supply = new Tag("recipe_CleanCirculation_machine.Step7.W4_valves_top");
        private Tag Cleaning2_Seq7_Backflow = new Tag("recipe_CleanCirculation_machine.Step7.W5_valves_buttom");
        private Tag Cleaning2_Seq7_duration = new Tag("recipe_CleanCirculation_machine.Step7.W6_time_step");
        private Tag Cleaning2_Seq8_cleaning_medium = new Tag("recipe_CleanCirculation_machine.Step8.W1_medium");
        private Tag Cleaning2_Seq8_cleaning_top = new Tag("recipe_CleanCirculation_machine.Step8.W2_pressure_top");
        private Tag Cleaning2_Seq8_cleaning_bottom = new Tag("recipe_CleanCirculation_machine.Step8.W3_buttom");
        private Tag Cleaning2_Seq8_Supply = new Tag("recipe_CleanCirculation_machine.Step8.W4_valves_top");
        private Tag Cleaning2_Seq8_Backflow = new Tag("recipe_CleanCirculation_machine.Step8.W5_valves_buttom");
        private Tag Cleaning2_Seq8_duration = new Tag("recipe_CleanCirculation_machine.Step8.W6_time_step");

        //
        private System.Timers.Timer tagSampleTimer;
        private const double tagSampleTime = 1000;

        private void triggerTagUpdate(object sender, ElapsedEventArgs e)
        {
            try
            {
                // Cleaning 1 machine
                if (myPLC.ReadTag(Cleaning1_Seq1_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    textBox2.Text = Cleaning1_Seq1_cleaning_medium.Value.ToString();
                    TimeStamp.Text = Cleaning1_Seq1_cleaning_medium.TimeStamp.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq1_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    textBox3.Text = Cleaning1_Seq1_cleaning_top.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq1_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    textBox4.Text = Cleaning1_Seq1_cleaning_bottom.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq1_Supply) == ResultCode.E_SUCCESS)
                {
                    textBox5.Text = Cleaning1_Seq1_Supply.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq1_Backflow) == ResultCode.E_SUCCESS)
                {
                    textBox6.Text = Cleaning1_Seq1_Backflow.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq1_duration) == ResultCode.E_SUCCESS)
                {
                    textBox7.Text = Cleaning1_Seq1_duration.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq2_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    textBox8.Text = Cleaning1_Seq2_cleaning_medium.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq2_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    textBox9.Text = Cleaning1_Seq2_cleaning_top.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq2_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    textBox10.Text = Cleaning1_Seq2_cleaning_bottom.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq2_Supply) == ResultCode.E_SUCCESS)
                {
                    textBox11.Text = Cleaning1_Seq2_Supply.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq2_Backflow) == ResultCode.E_SUCCESS)
                {
                    textBox12.Text = Cleaning1_Seq2_Backflow.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq2_duration) == ResultCode.E_SUCCESS)
                {
                    textBox13.Text = Cleaning1_Seq2_duration.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq3_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    textBox14.Text = Cleaning1_Seq3_cleaning_medium.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq3_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    textBox15.Text = Cleaning1_Seq3_cleaning_top.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq3_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    textBox16.Text = Cleaning1_Seq3_cleaning_bottom.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq3_Supply) == ResultCode.E_SUCCESS)
                {
                    textBox17.Text = Cleaning1_Seq3_Supply.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq3_Backflow) == ResultCode.E_SUCCESS)
                {
                    textBox18.Text = Cleaning1_Seq3_Backflow.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq3_duration) == ResultCode.E_SUCCESS)
                {
                    textBox19.Text = Cleaning1_Seq3_duration.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq4_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    textBox20.Text = Cleaning1_Seq4_cleaning_medium.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq4_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    textBox21.Text = Cleaning1_Seq4_cleaning_top.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq4_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    textBox22.Text = Cleaning1_Seq4_cleaning_bottom.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq4_Supply) == ResultCode.E_SUCCESS)
                {
                    textBox23.Text = Cleaning1_Seq4_Supply.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq4_Backflow) == ResultCode.E_SUCCESS)
                {
                    textBox24.Text = Cleaning1_Seq4_Backflow.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq4_duration) == ResultCode.E_SUCCESS)
                {
                    textBox25.Text = Cleaning1_Seq4_duration.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq5_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    textBox49.Text = Cleaning1_Seq5_cleaning_medium.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq5_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    textBox48.Text = Cleaning1_Seq5_cleaning_top.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq5_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    textBox47.Text = Cleaning1_Seq5_cleaning_bottom.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq5_Supply) == ResultCode.E_SUCCESS)
                {
                    textBox46.Text = Cleaning1_Seq5_Supply.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq5_Backflow) == ResultCode.E_SUCCESS)
                {
                    textBox45.Text = Cleaning1_Seq5_Backflow.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq5_duration) == ResultCode.E_SUCCESS)
                {
                    textBox44.Text = Cleaning1_Seq5_duration.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq6_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    textBox43.Text = Cleaning1_Seq6_cleaning_medium.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq6_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    textBox42.Text = Cleaning1_Seq6_cleaning_top.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq6_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    textBox41.Text = Cleaning1_Seq6_cleaning_bottom.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq6_Supply) == ResultCode.E_SUCCESS)
                {
                    textBox40.Text = Cleaning1_Seq6_Supply.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq6_Backflow) == ResultCode.E_SUCCESS)
                {
                    textBox39.Text = Cleaning1_Seq6_Backflow.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq6_duration) == ResultCode.E_SUCCESS)
                {
                    textBox38.Text = Cleaning1_Seq6_duration.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq7_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    textBox37.Text = Cleaning1_Seq7_cleaning_medium.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq7_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    textBox36.Text = Cleaning1_Seq7_cleaning_top.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq7_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    textBox35.Text = Cleaning1_Seq7_cleaning_bottom.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq7_Supply) == ResultCode.E_SUCCESS)
                {
                    textBox34.Text = Cleaning1_Seq7_Supply.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq7_Backflow) == ResultCode.E_SUCCESS)
                {
                    textBox33.Text = Cleaning1_Seq7_Backflow.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq7_duration) == ResultCode.E_SUCCESS)
                {
                    textBox32.Text = Cleaning1_Seq7_duration.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq8_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    textBox31.Text = Cleaning1_Seq8_cleaning_medium.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq8_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    textBox30.Text = Cleaning1_Seq8_cleaning_top.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq8_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    textBox29.Text = Cleaning1_Seq8_cleaning_bottom.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq8_Supply) == ResultCode.E_SUCCESS)
                {
                    textBox28.Text = Cleaning1_Seq8_Supply.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq8_Backflow) == ResultCode.E_SUCCESS)
                {
                    textBox27.Text = Cleaning1_Seq8_Backflow.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning1_Seq8_duration) == ResultCode.E_SUCCESS)
                {
                    textBox26.Text = Cleaning1_Seq8_duration.Value.ToString();
                }

                // Cleaning 2 machine
                if (myPLC.ReadTag(Cleaning2_Seq1_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    textBox97.Text = Cleaning2_Seq1_cleaning_medium.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq1_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    textBox96.Text = Cleaning2_Seq1_cleaning_top.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq1_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    textBox95.Text = Cleaning2_Seq1_cleaning_bottom.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq1_Supply) == ResultCode.E_SUCCESS)
                {
                    textBox94.Text = Cleaning2_Seq1_Supply.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq1_Backflow) == ResultCode.E_SUCCESS)
                {
                    textBox93.Text = Cleaning2_Seq1_Backflow.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq1_duration) == ResultCode.E_SUCCESS)
                {
                    textBox92.Text = Cleaning2_Seq1_duration.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq2_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    textBox91.Text = Cleaning2_Seq2_cleaning_medium.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq2_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    textBox90.Text = Cleaning2_Seq2_cleaning_top.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq2_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    textBox89.Text = Cleaning2_Seq2_cleaning_bottom.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq2_Supply) == ResultCode.E_SUCCESS)
                {
                    textBox88.Text = Cleaning2_Seq2_Supply.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq2_Backflow) == ResultCode.E_SUCCESS)
                {
                    textBox87.Text = Cleaning2_Seq2_Backflow.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq2_duration) == ResultCode.E_SUCCESS)
                {
                    textBox86.Text = Cleaning2_Seq2_duration.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq3_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    textBox85.Text = Cleaning2_Seq3_cleaning_medium.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq3_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    textBox84.Text = Cleaning2_Seq3_cleaning_top.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq3_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    textBox83.Text = Cleaning2_Seq3_cleaning_bottom.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq3_Supply) == ResultCode.E_SUCCESS)
                {
                    textBox82.Text = Cleaning2_Seq3_Supply.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq3_Backflow) == ResultCode.E_SUCCESS)
                {
                    textBox81.Text = Cleaning2_Seq3_Backflow.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq3_duration) == ResultCode.E_SUCCESS)
                {
                    textBox80.Text = Cleaning2_Seq3_duration.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq4_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    textBox79.Text = Cleaning2_Seq4_cleaning_medium.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq4_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    textBox78.Text = Cleaning2_Seq4_cleaning_top.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq4_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    textBox77.Text = Cleaning2_Seq4_cleaning_bottom.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq4_Supply) == ResultCode.E_SUCCESS)
                {
                    textBox76.Text = Cleaning2_Seq4_Supply.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq4_Backflow) == ResultCode.E_SUCCESS)
                {
                    textBox75.Text = Cleaning2_Seq4_Backflow.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq4_duration) == ResultCode.E_SUCCESS)
                {
                    textBox74.Text = Cleaning2_Seq4_duration.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq5_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    textBox73.Text = Cleaning2_Seq5_cleaning_medium.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq5_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    textBox72.Text = Cleaning2_Seq5_cleaning_top.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq5_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    textBox71.Text = Cleaning2_Seq5_cleaning_bottom.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq5_Supply) == ResultCode.E_SUCCESS)
                {
                    textBox70.Text = Cleaning2_Seq5_Supply.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq5_Backflow) == ResultCode.E_SUCCESS)
                {
                    textBox69.Text = Cleaning2_Seq5_Backflow.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq5_duration) == ResultCode.E_SUCCESS)
                {
                    textBox68.Text = Cleaning2_Seq5_duration.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq6_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    textBox67.Text = Cleaning2_Seq6_cleaning_medium.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq6_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    textBox66.Text = Cleaning2_Seq6_cleaning_top.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq6_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    textBox65.Text = Cleaning2_Seq6_cleaning_bottom.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq6_Supply) == ResultCode.E_SUCCESS)
                {
                    textBox64.Text = Cleaning2_Seq6_Supply.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq6_Backflow) == ResultCode.E_SUCCESS)
                {
                    textBox63.Text = Cleaning2_Seq6_Backflow.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq6_duration) == ResultCode.E_SUCCESS)
                {
                    textBox62.Text = Cleaning2_Seq6_duration.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq7_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    textBox61.Text = Cleaning2_Seq7_cleaning_medium.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq7_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    textBox60.Text = Cleaning2_Seq7_cleaning_top.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq7_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    textBox59.Text = Cleaning2_Seq7_cleaning_bottom.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq7_Supply) == ResultCode.E_SUCCESS)
                {
                    textBox58.Text = Cleaning2_Seq7_Supply.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq7_Backflow) == ResultCode.E_SUCCESS)
                {
                    textBox57.Text = Cleaning2_Seq7_Backflow.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq7_duration) == ResultCode.E_SUCCESS)
                {
                    textBox56.Text = Cleaning2_Seq7_duration.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq8_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    textBox55.Text = Cleaning2_Seq8_cleaning_medium.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq8_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    textBox54.Text = Cleaning2_Seq8_cleaning_top.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq8_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    textBox53.Text = Cleaning2_Seq8_cleaning_bottom.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq8_Supply) == ResultCode.E_SUCCESS)
                {
                    textBox52.Text = Cleaning2_Seq8_Supply.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq8_Backflow) == ResultCode.E_SUCCESS)
                {
                    textBox51.Text = Cleaning2_Seq8_Backflow.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning2_Seq8_duration) == ResultCode.E_SUCCESS)
                {
                    textBox50.Text = Cleaning2_Seq8_duration.Value.ToString();
                }
            }
            catch (Exception) // return when runnning false
            {
                MessageBox.Show("False");
            }
        }
        // click button Connect PLC
        private void button1_Click(object sender, EventArgs e)
        {
                try
                {
                    plc_Address = textBox1.Text.ToString();
                    myPLC = new Controller(Controller.CPU.LOGIX, plc_Address);
                    myPLC.Connect();

                    if (myPLC.IsConnected == true)
                    {
                        label2.Text = "PLC Connected";
                        label2.ForeColor = Color.Green;
                        tagSampleTimer = new System.Timers.Timer(tagSampleTime);
                        tagSampleTimer.Enabled = true;
                        tagSampleTimer.AutoReset = true;
                        tagSampleTimer.Elapsed += new System.Timers.ElapsedEventHandler(triggerTagUpdate);
                    }
                    if (myPLC.IsConnected == false)
                    {
                        label2.Text = "PLC Disconnect";
                        label2.ForeColor = Color.Red;
                    }
                }
                catch (Exception) // Return if false
                {
                    MessageBox.Show("Connect error");
                    label2.Text = "PLC Disconnect";
                    label2.ForeColor = Color.Red;
                }
        }

        // Another click
        private void label1_Click(object sender, EventArgs e) { }
        private void Form2_Load(object sender, EventArgs e) { }
        private void TimeStemp_Click(object sender, EventArgs e) { }
        private void label2_Click(object sender, EventArgs e) { }
        private void button4_Click(object sender, EventArgs e) 
        {
            if (myPLC.IsConnected == true)
            {
                label2.Text = "PLC Disconnect";
                label2.ForeColor = Color.Red;
            }
            if (myPLC.IsConnected == false)
            {
                label2.Text = "PLC Disconnect";
                label2.ForeColor = Color.Red;
            }
            myPLC.Disconnect();
            tagSampleTimer.Stop();
        }
        private void button2_Click(object sender, EventArgs e) { }
        private void groupBox1_Enter(object sender, EventArgs e) { }
        private void textBox2_TextChanged(object sender, EventArgs e) { }
        private void groupBox2_Enter(object sender, EventArgs e) { }
        private void groupBox4_Enter(object sender, EventArgs e) { }

    }
}
