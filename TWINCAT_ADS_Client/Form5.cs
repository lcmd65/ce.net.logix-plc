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
    public partial class Form5 : Form
    {
        private string plc_Address;
        private Controller myPLC = new Controller();


        //Configuration - Alarm
        private Tag Heating_Zone1_Alarm_Lower_Limit = new Tag("check_HZ1.ValueLL");
        private Tag Heating_Zone1_Alarm_Upper_Limit = new Tag("check_HZ1.ValueHH");
        private Tag Heating_Zone1_Warning_Lower_Limit = new Tag("check_HZ1.ValueLH");
        private Tag Heating_Zone1_Warning_Upper_Limit = new Tag("check_HZ1.ValueHL");
        private Tag Heating_Zone2_Alarm_Lower_Limit = new Tag("check_HZ2.ValueLL");
        private Tag Heating_Zone2_Alarm_Upper_Limit = new Tag("check_HZ2.ValueHH");
        private Tag Heating_Zone2_Warning_Lower_Limit = new Tag("check_HZ2.ValueLH");
        private Tag Heating_Zone2_Warning_Upper_Limit = new Tag("check_HZ2.ValueHL");
        private Tag Cooling_Zone_Alarm_Lower_Limit = new Tag("check_Cooling.ValueLL");
        private Tag Cooling_Zone_Alarm_Upper_Limit = new Tag("check_Cooling.ValueHH");
        private Tag NPR_gel_Level_Alarm_Lower_Limit = new Tag("check_Resist_level.ValueLL");
        private Tag NPR_gel_Level_Alarm_Upper_Limit = new Tag("check_Resist_level.ValueHH");
        private Tag NPR_gel_Level_Warning_Lower_Limit = new Tag("check_Resist_level.ValueLH");
        private Tag NPR_gel_Level_Warning_Upper_Limit = new Tag("check_Resist_level.ValueHL");
        private Tag NPR_gel_Flow_Alarm_Lower_Limit = new Tag("check_Resist_flow.ValueLL");
        private Tag NPR_gel_Flow_Alarm_Upper_Limit = new Tag("check_Resist_flow.ValueHH");
        private Tag NPR_gel_Flow_Warning_Lower_Limit = new Tag("check_Resist_flow.ValueLH");
        private Tag NPR_gel_Flow_Warning_Upper_Limit = new Tag("check_Resist_flow.ValueHL");
        private Tag Density_Alarm_Lower_Limit = new Tag("check_Density.ValueLL");
        private Tag Density_Alarm_Upper_Limit = new Tag("check_Density.ValueHH");
        private Tag Density_Warning_Lower_Limit = new Tag("check_Density.ValueLH");
        private Tag Density_Warning_Upper_Limit = new Tag("check_Density.ValueHL");
        private Tag NPR_gel_temperature_Alarm_Lower_Limit = new Tag("check_Temperature.ValueLL");
        private Tag NPR_gel_temperature_Alarm_Upper_Limit = new Tag("check_Temperature.ValueHH");
        private Tag NPR_gel_temperature_Warning_Lower_Limit = new Tag("check_Temperature.ValueLH");
        private Tag NPR_gel_temperature_Warning_Upper_Limit = new Tag("check_Temperature.ValueHL");

        //Parameters Times
        private Tag Heat_up_time = new Tag("heatup");
        private Tag Delay_shut_off_auto_cleaning = new Tag("DelaytimeAutocleaning");


        // Tag Sample Timer
        private System.Timers.Timer tagSampleTimer;
        private const double tagSampleTime = 1000;

        public Form5()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        // Connect to PLC
        private void button1_Click_1(object sender, EventArgs e)
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

        // Trigger and fill data realtime from PLC signals
        private void triggerTagUpdate(object sender, ElapsedEventArgs e)
        {
            try
            {
                //Configuration - Alarm
                if (myPLC.ReadTag(Heating_Zone1_Alarm_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox38.Text = Heating_Zone1_Alarm_Lower_Limit.ToString();
                }
                if (myPLC.ReadTag(Heating_Zone1_Alarm_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox39.Text = Heating_Zone1_Alarm_Upper_Limit.ToString();
                }
                if (myPLC.ReadTag(Heating_Zone1_Warning_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox40.Text = Heating_Zone1_Warning_Lower_Limit.ToString();
                }
                if (myPLC.ReadTag(Heating_Zone1_Warning_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox41.Text = Heating_Zone1_Warning_Upper_Limit.ToString();
                }
                if (myPLC.ReadTag(Heating_Zone2_Alarm_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox42.Text = Heating_Zone2_Alarm_Lower_Limit.ToString();
                }
                if (myPLC.ReadTag(Heating_Zone2_Alarm_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox43.Text = Heating_Zone2_Alarm_Upper_Limit.ToString();
                }
                if (myPLC.ReadTag(Heating_Zone2_Warning_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox44.Text = Heating_Zone2_Warning_Lower_Limit.ToString();
                }
                if (myPLC.ReadTag(Heating_Zone2_Warning_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox45.Text = Heating_Zone2_Warning_Upper_Limit.ToString();
                }
                if (myPLC.ReadTag(Cooling_Zone_Alarm_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox46.Text = Cooling_Zone_Alarm_Lower_Limit.ToString();
                }
                if (myPLC.ReadTag(Cooling_Zone_Alarm_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox47.Text = Cooling_Zone_Alarm_Upper_Limit.ToString();
                }
                if (myPLC.ReadTag(NPR_gel_Level_Alarm_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox48.Text = NPR_gel_Level_Alarm_Lower_Limit.ToString();
                }
                if (myPLC.ReadTag(NPR_gel_Level_Alarm_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox49.Text = NPR_gel_Level_Alarm_Upper_Limit.ToString();
                }
                if (myPLC.ReadTag(NPR_gel_Level_Warning_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox50.Text = NPR_gel_Level_Warning_Lower_Limit.ToString();
                }
                if (myPLC.ReadTag(NPR_gel_Level_Warning_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox51.Text = NPR_gel_Level_Warning_Upper_Limit.ToString();
                }
                if (myPLC.ReadTag(NPR_gel_Flow_Alarm_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox52.Text = NPR_gel_Flow_Alarm_Lower_Limit.ToString();
                }
                if (myPLC.ReadTag(NPR_gel_Flow_Alarm_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox53.Text = NPR_gel_Flow_Alarm_Upper_Limit.ToString();
                }
                if (myPLC.ReadTag(NPR_gel_Flow_Warning_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox54.Text = NPR_gel_Flow_Warning_Lower_Limit.ToString();
                }
                if (myPLC.ReadTag(NPR_gel_Flow_Warning_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox55.Text = NPR_gel_Flow_Warning_Upper_Limit.ToString();
                }
                if (myPLC.ReadTag(Density_Alarm_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox56.Text = Density_Alarm_Lower_Limit.ToString();
                }
                if (myPLC.ReadTag(Density_Alarm_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox57.Text = Density_Alarm_Upper_Limit.ToString();
                }
                if (myPLC.ReadTag(Density_Warning_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox58.Text = Density_Warning_Lower_Limit.ToString();
                }
                if (myPLC.ReadTag(Density_Warning_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox59.Text = Density_Warning_Upper_Limit.ToString();
                }
                if (myPLC.ReadTag(NPR_gel_temperature_Alarm_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox60.Text = NPR_gel_temperature_Alarm_Lower_Limit.ToString();
                }
                if (myPLC.ReadTag(NPR_gel_temperature_Alarm_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox61.Text = NPR_gel_temperature_Alarm_Upper_Limit.ToString();
                }
                if (myPLC.ReadTag(NPR_gel_temperature_Warning_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox62.Text = NPR_gel_temperature_Warning_Lower_Limit.ToString();
                }
                if (myPLC.ReadTag(NPR_gel_temperature_Warning_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    textBox63.Text = NPR_gel_temperature_Warning_Upper_Limit.ToString();
                }
            }

            catch (Exception)
            {
                MessageBox.Show("Error");
            }
        }
        private void button2_Click_1(object sender, EventArgs e)
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
        private void button4_Click_1(object sender, EventArgs e)
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
