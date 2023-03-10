using Logix;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;

namespace TWINCAT_ADS_Client
{
    public partial class Form4 : Form
    {
        private string plc_Address;
        private Controller myPLC = new Controller();

        public Form4()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        //Jet Dryer - Transport
        private Tag Jet_Dry_Oven_Transport_set_value = new Tag("drive._030TRA_100T2.set_speed");
        private Tag Jet_Dry_Oven_Transport_offset = new Tag("drive._030TRA_100T2.set_offset_speed");
        private Tag Jet_Dry_Oven_Transport_actual = new Tag("drive._030TRA_100T2.actual_speed_r");
        private Tag Heating_Zone1_Set_Value = new Tag("DB102.set_value.temp_heat1");
        private Tag Heating_Zone2_Set_Value = new Tag("DB102.set_value.temp_heat2");
        private Tag Cooling_temperature = new Tag("Pen3_Act_Temperature");
        private Tag Jet_Dry_Oven_Transport_on_off = new Tag("M.after_run_on");
        private Tag Heating_Zone1_on_off = new Tag("DB5.Heat_heating1.Run");
        private Tag Heating_Zone2_on_off = new Tag("DB5.Heat_heating2.Run");

        //Jet Dryer - Ventilation
        private Tag Jet1_Exhaust = new Tag("drive._030JET1_100T2.set_speed");
        private Tag Jet1_Circulation1 = new Tag("drive._030JET1_120T2.set_speed");
        private Tag Jet1_Circulation2 = new Tag("drive._030JET1_130T2.set_speed");
        private Tag Jet2_Exhaust = new Tag("drive._030JET2_100T2.set_speed");
        private Tag Jet2_Circulation1 = new Tag("drive._030JET2_120T2.set_speed");
        private Tag Jet2_Circulation2 = new Tag("drive._030JET2_130T2.set_speed");
        private Tag JetK_Exhaust = new Tag("drive._030JETK_100T2.set_speed");
        private Tag JetK_Circulation1 = new Tag("drive._030JETK_120T2.set_speed");
        private Tag JetK_Circulation2 = new Tag("drive._030JETK_130T2.set_speed");
        private Tag JetK_Exhaust2 = new Tag("drive._030JETK_110T2.set_speed");

        // Infeed stranport 
        private Tag Infeed_Transport_set_value = new Tag("drive._040TRA_100T2.set_speed");
        private Tag Infeed_Transport_offset = new Tag("drive._040TRA_100T2.set_offset_speed");
        private Tag Infeed_Transport_actual = new Tag("drive._040TRA_100T2.actual_speed_r");
        private Tag Infeed_Transport_on_off = new Tag("DB5.Roller_infeed.Run");

        // Tag Sample Timer
        private System.Timers.Timer tagSampleTimer;
        private const double tagSampleTime = 1000;

        private void triggerTagUpdate(object sender, ElapsedEventArgs e)
        {
            try
            {
                if (myPLC.ReadTag(Jet_Dry_Oven_Transport_set_value) == ResultCode.E_SUCCESS)
                {
                    textBox2.Text = Jet_Dry_Oven_Transport_set_value.Value.ToString();
                    TimeStamp.Text = Jet_Dry_Oven_Transport_set_value.TimeStamp.ToString();
                }
                if (myPLC.ReadTag(Jet_Dry_Oven_Transport_offset) == ResultCode.E_SUCCESS)
                {
                    textBox3.Text = Jet_Dry_Oven_Transport_offset.Value.ToString();
                }
                if (myPLC.ReadTag(Jet_Dry_Oven_Transport_actual) == ResultCode.E_SUCCESS)
                {
                    textBox4.Text = Jet_Dry_Oven_Transport_actual.Value.ToString();
                }
                if (myPLC.ReadTag(Heating_Zone1_Set_Value) == ResultCode.E_SUCCESS)
                {
                    textBox5.Text = Heating_Zone1_Set_Value.Value.ToString();
                }
                if (myPLC.ReadTag(Heating_Zone2_Set_Value) == ResultCode.E_SUCCESS)
                {
                    textBox6.Text = Heating_Zone2_Set_Value.Value.ToString();
                }
                if (myPLC.ReadTag(Cooling_temperature) == ResultCode.E_SUCCESS)
                {
                    textBox7.Text = Cooling_temperature.Value.ToString();
                }
                if (myPLC.ReadTag(Jet_Dry_Oven_Transport_on_off) == ResultCode.E_SUCCESS)
                {
                    textBox8.Text = Jet_Dry_Oven_Transport_on_off.Value.ToString();
                }
                if (myPLC.ReadTag(Heating_Zone1_on_off) == ResultCode.E_SUCCESS)
                {
                    textBox9.Text = Heating_Zone1_on_off.Value.ToString();
                }
                if (myPLC.ReadTag(Heating_Zone2_on_off) == ResultCode.E_SUCCESS)
                {
                    textBox10.Text = Heating_Zone2_on_off.Value.ToString();
                }

                // Jet Dry Ventilation
                if (myPLC.ReadTag(Jet1_Exhaust) == ResultCode.E_SUCCESS)
                {
                    textBox11.Text = Jet1_Exhaust.Value.ToString();
                }
                if (myPLC.ReadTag(Jet1_Circulation1) == ResultCode.E_SUCCESS)
                {
                    textBox12.Text = Jet1_Circulation1.Value.ToString();
                }
                if (myPLC.ReadTag(Jet1_Circulation2) == ResultCode.E_SUCCESS)
                {
                    textBox13.Text = Jet1_Circulation2.Value.ToString();
                }
                if (myPLC.ReadTag(Jet2_Exhaust) == ResultCode.E_SUCCESS)
                {
                    textBox14.Text = Jet2_Exhaust.Value.ToString();
                }
                if (myPLC.ReadTag(Jet2_Circulation1) == ResultCode.E_SUCCESS)
                {
                    textBox15.Text = Jet2_Circulation1.Value.ToString();
                }
                if (myPLC.ReadTag(Jet2_Circulation2) == ResultCode.E_SUCCESS)
                {
                    textBox16.Text = Jet2_Circulation2.Value.ToString();
                }
                if (myPLC.ReadTag(JetK_Exhaust) == ResultCode.E_SUCCESS)
                {
                    textBox17.Text = JetK_Exhaust.Value.ToString();
                }
                if (myPLC.ReadTag(JetK_Circulation1) == ResultCode.E_SUCCESS)
                {
                    textBox18.Text = JetK_Circulation1.Value.ToString();
                }
                if (myPLC.ReadTag(JetK_Circulation2) == ResultCode.E_SUCCESS)
                {
                    textBox19.Text = JetK_Circulation2.Value.ToString();
                }
                if (myPLC.ReadTag(JetK_Exhaust2) == ResultCode.E_SUCCESS)
                {
                    textBox20.Text = JetK_Exhaust2.Value.ToString();
                }

                // Infeed transport
                if (myPLC.ReadTag(Infeed_Transport_set_value) == ResultCode.E_SUCCESS)
                {
                    textBox21.Text = Infeed_Transport_set_value.Value.ToString();
                }
                if (myPLC.ReadTag(Infeed_Transport_offset) == ResultCode.E_SUCCESS)
                {
                    textBox22.Text = Infeed_Transport_offset.Value.ToString();
                }
                if (myPLC.ReadTag(Infeed_Transport_actual) == ResultCode.E_SUCCESS)
                {
                    textBox23.Text = Infeed_Transport_actual.Value.ToString();
                }
                if (myPLC.ReadTag(Infeed_Transport_on_off) == ResultCode.E_SUCCESS)
                {
                    textBox24.Text = Infeed_Transport_on_off.Value.ToString();
                }

            }
            catch (Exception)
            {
                MessageBox.Show("Error");
            }
        }

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
            catch (Exception)
            {
                MessageBox.Show("Connect error");
                label2.Text = "PLC Disconnect";
                label2.ForeColor = Color.Red;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                tagSampleTimer = new System.Timers.Timer(tagSampleTime);
                tagSampleTimer.Enabled = true;
                tagSampleTimer.AutoReset = true;
                tagSampleTimer.Elapsed += new System.Timers.ElapsedEventHandler(triggerTagUpdate);
            }
            catch (Exception)
            {
                MessageBox.Show("Connect error");
            }
        }

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
    }
}
