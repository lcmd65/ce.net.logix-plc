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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace TWINCAT_ADS_Client
{
    public partial class Form6 : Form
    {
        private string plc_Address;
        private Controller myPLC = new Controller();

        //Parameters Times
        private Tag Heat_up_time = new Tag("heatup");
        private Tag Delay_shut_off_auto_cleaning = new Tag("DelaytimeAutocleaning");

        //Manual
        private Tag Pump_out_material_tank_manual = new Tag("DB5.Pump_out_material.Run");
        private Tag Material_pump_manual = new Tag("DB5.Material_pump.Run");
        private Tag Waste_water_pump_manual = new Tag("DB5.Waste_pump.Run");
        private Tag Infeed_transport_manual = new Tag("DB5.Roller_infeed.Run");
        //private Tag Transport_UV_manual = new Tag("");
        private Tag Enable_lift_lower_manual = new Tag("DB5.Rel_lift_lower.Run");
        private Tag Roller_coater_transport_manual = new Tag("DB5.Coater_drive.Run");
        private Tag Heating_Zone_1_Manual = new Tag("DB5.Heat_heating1.Run");
        private Tag Exhaust_air_heating_zone_1_Manual = new Tag("DB5.Heat_exhaust_fan1.Run");
        private Tag Recirculating_air_1_heating_zone_1_Manual = new Tag("DB5.Heat_circ_fan1.Run");
        private Tag Recirculating_air_2_heating_zone_1_Manual = new Tag("DB5.Heat_circ_fan2.Run");
        private Tag Dryer_transport = new Tag("DB5.Dryer_drive.Run");
        private Tag Heating_Zone_2_Manual = new Tag("DB5.Heat_heating2.Run");
        private Tag Exhaust_air_heating_zone_2_Manual = new Tag("DB5.Heat_exhaust_fan2.Run");
        private Tag Recirculating_air_1_heating_zone_2_Manual = new Tag("DB5.Heat_circ_fan3.Run");
        private Tag Recirculating_air_2_heating_zone_2_Manual = new Tag("DB5.Heat_circ_fan4.Run");
        private Tag Exhaust_air_1_cooling_Manual = new Tag("DB5.Cooling_Exhaust_1.Run");
        private Tag Exhaust_air_2_cooling_Manual = new Tag("DB5.Cooling_Exhaust_2.Run");
        private Tag Recirculating_air_1_cooling_Manual = new Tag("DB5.Cooling_Circulation_1.Run");
        private Tag Recirculating_air_2_cooling_Manual = new Tag("DB5.Cooling_Circulation_2.Run");
        private Tag Scraper_oscillation_roller_top_Manual = new Tag("DB5.Swifel_oszillation.Run");
        private Tag Coating_Manual = new Tag("DB5.Coating_on.Run");

        // Tag Sample Timer
        private System.Timers.Timer tagSampleTimer;
        private const double tagSampleTime = 1000;

        public Form6()
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
                // Roller Coater - Machine
                if (myPLC.ReadTag(Pump_out_material_tank_manual) == ResultCode.E_SUCCESS)
                {
                    textBox2.Text = Pump_out_material_tank_manual.Value.ToString();
                    TimeStamp.Text = Pump_out_material_tank_manual.TimeStamp.ToString();
                }
                if (myPLC.ReadTag(Material_pump_manual) == ResultCode.E_SUCCESS)
                {
                    textBox3.Text = Material_pump_manual.Value.ToString();
                }
                if (myPLC.ReadTag(Waste_water_pump_manual) == ResultCode.E_SUCCESS)
                {
                    textBox4.Text = Waste_water_pump_manual.Value.ToString();
                }
                if (myPLC.ReadTag(Infeed_transport_manual) == ResultCode.E_SUCCESS)
                {
                    textBox5.Text = Infeed_transport_manual.Value.ToString();
                }
                //if (myPLC.ReadTag(Transport_UV_manual) == ResultCode.E_SUCCESS)
                //{
                    //textBox6.Text = Transport_UV_manual.Value.ToString();
                //}
                if (myPLC.ReadTag(Enable_lift_lower_manual) == ResultCode.E_SUCCESS)
                {
                    textBox7.Text = Enable_lift_lower_manual.Value.ToString();
                }
                if (myPLC.ReadTag(Roller_coater_transport_manual) == ResultCode.E_SUCCESS)
                {
                    textBox8.Text = Roller_coater_transport_manual.Value.ToString();
                }
                if (myPLC.ReadTag(Heating_Zone_1_Manual) == ResultCode.E_SUCCESS)
                {
                    textBox9.Text = Heating_Zone_1_Manual.Value.ToString();
                }
                if (myPLC.ReadTag(Exhaust_air_heating_zone_1_Manual) == ResultCode.E_SUCCESS)
                {
                    textBox10.Text = Exhaust_air_heating_zone_1_Manual.Value.ToString();
                }
                if (myPLC.ReadTag(Recirculating_air_1_heating_zone_1_Manual) == ResultCode.E_SUCCESS)
                {
                    textBox11.Text = Recirculating_air_1_heating_zone_1_Manual.Value.ToString();
                }
                if (myPLC.ReadTag(Recirculating_air_2_heating_zone_1_Manual) == ResultCode.E_SUCCESS)
                {
                    textBox12.Text = Recirculating_air_2_heating_zone_1_Manual.Value.ToString();
                }
                if (myPLC.ReadTag(Dryer_transport) == ResultCode.E_SUCCESS)
                {
                    textBox13.Text = Dryer_transport.Value.ToString();
                }
                if (myPLC.ReadTag(Heating_Zone_2_Manual) == ResultCode.E_SUCCESS)
                {
                    textBox14.Text = Heating_Zone_2_Manual.Value.ToString();
                }
                if (myPLC.ReadTag(Exhaust_air_heating_zone_2_Manual) == ResultCode.E_SUCCESS)
                {
                    textBox15.Text = Exhaust_air_heating_zone_2_Manual.Value.ToString();
                }
                if (myPLC.ReadTag(Recirculating_air_1_heating_zone_2_Manual) == ResultCode.E_SUCCESS)
                {
                    textBox16.Text = Recirculating_air_1_heating_zone_2_Manual.Value.ToString();
                }
                if (myPLC.ReadTag(Recirculating_air_2_heating_zone_2_Manual) == ResultCode.E_SUCCESS)
                {
                    textBox17.Text = Recirculating_air_2_heating_zone_2_Manual.Value.ToString();
                }
                if (myPLC.ReadTag(Exhaust_air_1_cooling_Manual) == ResultCode.E_SUCCESS)
                {
                    textBox18.Text = Exhaust_air_1_cooling_Manual.Value.ToString();
                }
                if (myPLC.ReadTag(Exhaust_air_2_cooling_Manual) == ResultCode.E_SUCCESS)
                {
                    textBox19.Text = Exhaust_air_2_cooling_Manual.Value.ToString();
                }
                if (myPLC.ReadTag(Recirculating_air_1_cooling_Manual) == ResultCode.E_SUCCESS)
                {
                    textBox20.Text = Recirculating_air_1_cooling_Manual.Value.ToString();
                }
                if (myPLC.ReadTag(Recirculating_air_2_cooling_Manual) == ResultCode.E_SUCCESS)
                {
                    textBox21.Text = Recirculating_air_2_cooling_Manual.Value.ToString();
                }
                if (myPLC.ReadTag(Scraper_oscillation_roller_top_Manual) == ResultCode.E_SUCCESS)
                {
                    textBox22.Text = Scraper_oscillation_roller_top_Manual.Value.ToString();
                }
                if (myPLC.ReadTag(Coating_Manual) == ResultCode.E_SUCCESS)
                {
                    textBox23.Text = Coating_Manual.Value.ToString();
                }

                if (myPLC.ReadTag(Heat_up_time) == ResultCode.E_SUCCESS)
                {
                    textBox24.Text = Heat_up_time.Value.ToString();
                }
                if (myPLC.ReadTag(Delay_shut_off_auto_cleaning) == ResultCode.E_SUCCESS)
                {
                    textBox25.Text = Delay_shut_off_auto_cleaning.Value.ToString();
                }
            }

            catch (Exception)
            {
                MessageBox.Show("Error");
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

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
