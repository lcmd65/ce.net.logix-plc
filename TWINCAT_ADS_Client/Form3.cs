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
    public partial class Form3 : Form
    {
        private string plc_Address;
        private Controller myPLC = new Controller();

        //UVExpose_Machine UV
        private Tag UV_Transport_set_value = new Tag("UVDB.v.SetSpeed");
        private Tag UV_Transport_offset = new Tag("UVDB.v.OffsetSpeed");
        private Tag UV_Transport_actual = new Tag("UVDB.v.IWSpeed");
        private Tag Number_of_boards_untill_next_measurement = new Tag("UVDB.v.UV_Meas_Piece_Set");
        private Tag Number_of_boards_untill_next_measurement_actual = new Tag("UVDB.v.UV_Measrement_PieceCoun");
        private Tag UV_measurement_start = new Tag("UVDB.v.Set_autom_UV1_Measureme");
        private Tag UV_transport_on_off = new Tag("UVDB.m.M_TranspOn");
        private Tag automatic_measurement_on_off = new Tag("UVDB.v.Presel_autom_Meas_hf");
        private Tag Preselection_UV1 = new Tag("UVDB.v.presel_fl.1");
        private Tag Preselection_UV2 = new Tag("UVDB.v.presel_fl.2");

        //UVExpose_Machine Ventilation
        private Tag UV1_Exhaust_set = new Tag("drive._040UVM1_100T2.set_speed");
        private Tag UV1_Exhaust_actual = new Tag("drive._040UVM1_100T2.actual_speed_r");
        private Tag UV2_Exhaust_set = new Tag("drive._040UVM2_100T2.set_speed");
        private Tag UV2_Exhaust_actual = new Tag("drive._040UVM2_100T2.actual_speed_r");

        //UVExpose_Manual general
        private Tag UV_lamp1_on_off = new Tag("UVDB.m.M_On_UVS1");
        private Tag UV_lamp2_on_off = new Tag("UVDB.m.M_On_UVS2");
        private Tag Exhaust_air_UV_lamp1_on_off = new Tag("UVDB.v.ExAirOn_UVS1");
        private Tag Exhaust_air_UV_lamp2_on_off = new Tag("UVDB.v.ExAirOn_UVS2");

        //UVExpose_Prameter_UV_calibration
        private Tag teach_max_UV_lamp1_on_off = new Tag("UVDB.v.TeachMax_UVS1");
        private Tag teach_min_UV_lamp1_on_off = new Tag("UVDB.v.TeachMin_UVS1");
        private Tag teach_max_UV_lamp1_calibration_value = new Tag("UVDB.v.Cal_MaxValue_UVS1");
        private Tag teach_min_UV_lamp1_calibration_value = new Tag("UVDB.v.Cal_MinValue_UVS1");
        private Tag UV_lamp1_sensor_max_values = new Tag("UVDB.v.ANA_MaxValue_UVS1");
        private Tag UV_lamp1_sensor_min_values = new Tag("UVDB.v.ANA_MinValue_UVS1");
        private Tag actual_lamp1_value = new Tag("UVDB.v.ActValue_Lamp_UVS1");
        private Tag teach_max_UV_lamp2_on_off = new Tag("UVDB.v.TeachMax_UVS2");
        private Tag teach_min_UV_lamp2_on_off = new Tag("UVDB.v.TeachMin_UVS2");
        private Tag teach_max_UV_lamp2_calibration_value = new Tag("UVDB.v.Cal_MaxValue_UVS2");
        private Tag teach_min_UV_lamp2_calibration_value = new Tag("UVDB.v.Cal_MinValue_UVS2");
        private Tag UV_lamp2_sensor_max_values = new Tag("UVDB.v.ANA_MaxValue_UVS2");
        private Tag UV_lamp2_sensor_min_values = new Tag("UVDB.v.ANA_MinValue_UVS2");
        private Tag actual_lamp2_value = new Tag("UVDB.v.ActValue_Lamp_UVS2");

        //UVExpose_Parameter_Configuration
        private Tag UV_radiation_alarm_lower_limit = new Tag("UVDB.v.Error_UV_min");
        private Tag UV_radiation_alarm_upper_limit = new Tag("UVDB.v.Error_UV_max");
        private Tag UV_radiation_warning_lower_limit = new Tag("UVDB.v.Warning_UV_min");
        private Tag UV_radiation_warning_upper_limit = new Tag("UVDB.v.Warning_UV_max");
        private Tag economic_life_time_UV_lamp1 = new Tag("UVDB_SetOpTimeUV1_Temp");
        private Tag economic_life_time_UV_lamp2 = new Tag("UVDB_SetOpTimeUV2_Temp");
        private Tag radiation_switch_UV1_UV2 = new Tag("UVDB.v.SwitchOver_UV");
        private Tag automatic_switch_UV1_UV2_on_off = new Tag("UVDB.v.MesAutoOff");

        // UVExpose_Parameter_timer
        private Tag Delay_time_switch_off_exhaust_air = new Tag("ValFollowUpExAirPanel");
        private Tag Delay_time_start_exhaust_air = new Tag("UV_DelOnExAir");

        // UVExpose_Data UV_Recipe
        private Tag UV1ActValue = new Tag("UV1ActValue");
        private Tag UV2ActValue = new Tag("UV2ActValue");
        private Tag UV_lamp1_operating_grade = new Tag("UVDB.v.Percent_Lamp_UVS1");
        private Tag UVsetValue1 = new Tag("UVsetValue1");
        private Tag UVsetValue2 = new Tag("UVsetValue2");
        private Tag UV_lamp2_operating_grade = new Tag("UVDB.v.Percent_Lamp_UVS2");

        private System.Timers.Timer tagSampleTimer;
        private const double tagSampleTime = 1000;

        public Form3()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        // Trigger and fill data realtime from PLC signals
        private void triggerTagUpdate(object sender, ElapsedEventArgs e)
        {
            try
            {
                // Roller Coater - Machine
                if (myPLC.ReadTag(UV_Transport_set_value) == ResultCode.E_SUCCESS)
                {
                    textBox2.Text = UV_Transport_set_value.Value.ToString();
                    TimeStamp.Text = UV_Transport_set_value.TimeStamp.ToString();
                }
                if (myPLC.ReadTag(UV_Transport_offset) == ResultCode.E_SUCCESS)
                {
                    textBox3.Text = UV_Transport_offset.Value.ToString();
                }
                if (myPLC.ReadTag(UV_Transport_actual) == ResultCode.E_SUCCESS)
                {
                    textBox4.Text = UV_Transport_actual.Value.ToString();
                }
                if (myPLC.ReadTag(Number_of_boards_untill_next_measurement) == ResultCode.E_SUCCESS)
                {
                    textBox5.Text = Number_of_boards_untill_next_measurement.Value.ToString();
                }
                if (myPLC.ReadTag(Number_of_boards_untill_next_measurement_actual) == ResultCode.E_SUCCESS)
                {
                    textBox6.Text = Number_of_boards_untill_next_measurement_actual.Value.ToString();
                }
                if (myPLC.ReadTag(UV_measurement_start) == ResultCode.E_SUCCESS)
                {
                    textBox7.Text = UV_measurement_start.Value.ToString();
                }
                if (myPLC.ReadTag(UV_transport_on_off) == ResultCode.E_SUCCESS)
                {
                    textBox8.Text = UV_transport_on_off.Value.ToString();
                }
                if (myPLC.ReadTag(automatic_measurement_on_off) == ResultCode.E_SUCCESS)
                {
                    textBox9.Text = automatic_measurement_on_off.Value.ToString();
                }
                if (myPLC.ReadTag(Preselection_UV1) == ResultCode.E_SUCCESS)
                {
                    textBox10.Text = Preselection_UV1.Value.ToString();
                }
                if (myPLC.ReadTag(Preselection_UV2) == ResultCode.E_SUCCESS)
                {
                    textBox11.Text = Preselection_UV2.Value.ToString();
                }

                // UV ventilation
                if (myPLC.ReadTag(UV1_Exhaust_set) == ResultCode.E_SUCCESS)
                {
                    textBox12.Text = UV1_Exhaust_set.Value.ToString();
                }
                if (myPLC.ReadTag(UV1_Exhaust_actual) == ResultCode.E_SUCCESS)
                {
                    textBox13.Text = UV1_Exhaust_actual.Value.ToString();
                }
                if (myPLC.ReadTag(UV2_Exhaust_set) == ResultCode.E_SUCCESS)
                {
                    textBox14.Text = UV2_Exhaust_set.Value.ToString();
                }
                if (myPLC.ReadTag(UV2_Exhaust_actual) == ResultCode.E_SUCCESS)
                {
                    textBox15.Text = UV2_Exhaust_actual.Value.ToString();
                }

                // UV manual General
                if (myPLC.ReadTag(UV_lamp1_on_off) == ResultCode.E_SUCCESS)
                {
                    textBox16.Text = UV_lamp1_on_off.Value.ToString();
                }
                if (myPLC.ReadTag(UV_lamp2_on_off) == ResultCode.E_SUCCESS)
                {
                    textBox17.Text = UV_lamp2_on_off.Value.ToString();
                }
                if (myPLC.ReadTag(Exhaust_air_UV_lamp1_on_off) == ResultCode.E_SUCCESS)
                {
                    textBox18.Text = Exhaust_air_UV_lamp1_on_off.Value.ToString();
                }
                if (myPLC.ReadTag(Exhaust_air_UV_lamp2_on_off) == ResultCode.E_SUCCESS)
                {
                    textBox19.Text = Exhaust_air_UV_lamp2_on_off.Value.ToString();
                }

                // UV 
                if (myPLC.ReadTag(teach_max_UV_lamp1_on_off) == ResultCode.E_SUCCESS)
                {
                    textBox20.Text = teach_max_UV_lamp1_on_off.Value.ToString();
                }
                if (myPLC.ReadTag(teach_min_UV_lamp1_on_off) == ResultCode.E_SUCCESS)
                {
                    textBox21.Text = teach_min_UV_lamp1_on_off.Value.ToString();
                }
                if (myPLC.ReadTag(teach_max_UV_lamp1_calibration_value) == ResultCode.E_SUCCESS)
                {
                    textBox22.Text = teach_max_UV_lamp1_calibration_value.Value.ToString();
                }
                if (myPLC.ReadTag(teach_min_UV_lamp1_calibration_value) == ResultCode.E_SUCCESS)
                {
                    textBox23.Text = teach_min_UV_lamp1_calibration_value.Value.ToString();
                }
                if (myPLC.ReadTag(UV_lamp1_sensor_max_values) == ResultCode.E_SUCCESS)
                {
                    textBox24.Text = UV_lamp1_sensor_max_values.Value.ToString();
                }
                if (myPLC.ReadTag(UV_lamp1_sensor_min_values) == ResultCode.E_SUCCESS)
                {
                    textBox25.Text = UV_lamp1_sensor_min_values.Value.ToString();
                }
                if (myPLC.ReadTag(actual_lamp1_value) == ResultCode.E_SUCCESS)
                {
                    textBox26.Text = actual_lamp1_value.Value.ToString();
                }
                if (myPLC.ReadTag(teach_max_UV_lamp2_on_off) == ResultCode.E_SUCCESS)
                {
                    textBox27.Text = teach_max_UV_lamp2_on_off.Value.ToString();
                }
                if (myPLC.ReadTag(teach_min_UV_lamp2_on_off) == ResultCode.E_SUCCESS)
                {
                    textBox28.Text = teach_min_UV_lamp2_on_off.Value.ToString();
                }
                if (myPLC.ReadTag(teach_max_UV_lamp2_calibration_value) == ResultCode.E_SUCCESS)
                {
                    textBox29.Text = teach_max_UV_lamp2_calibration_value.Value.ToString();
                }
                if (myPLC.ReadTag(teach_min_UV_lamp2_calibration_value) == ResultCode.E_SUCCESS)
                {
                    textBox30.Text = teach_min_UV_lamp2_calibration_value.Value.ToString();
                }
                if (myPLC.ReadTag(UV_lamp2_sensor_max_values) == ResultCode.E_SUCCESS)
                {
                    textBox31.Text = UV_lamp2_sensor_max_values.Value.ToString();
                }
                if (myPLC.ReadTag(UV_lamp2_sensor_min_values) == ResultCode.E_SUCCESS)
                {
                    textBox32.Text = UV_lamp2_sensor_min_values.Value.ToString();
                }
                if (myPLC.ReadTag(actual_lamp2_value) == ResultCode.E_SUCCESS)
                {
                    textBox33.Text = actual_lamp2_value.Value.ToString();
                }

                //
                if (myPLC.ReadTag(UV_radiation_alarm_lower_limit) == ResultCode.E_SUCCESS)
                {
                    textBox34.Text = UV_radiation_alarm_lower_limit.Value.ToString();
                }
                if (myPLC.ReadTag(UV_radiation_alarm_upper_limit) == ResultCode.E_SUCCESS)
                {
                    textBox35.Text = UV_radiation_alarm_upper_limit.Value.ToString();
                }
                if (myPLC.ReadTag(UV_radiation_warning_lower_limit) == ResultCode.E_SUCCESS)
                {
                    textBox36.Text = UV_radiation_warning_lower_limit.Value.ToString();
                }
                if (myPLC.ReadTag(UV_radiation_warning_upper_limit) == ResultCode.E_SUCCESS)
                {
                    textBox37.Text = UV_radiation_warning_upper_limit.Value.ToString();
                }
                if (myPLC.ReadTag(economic_life_time_UV_lamp1) == ResultCode.E_SUCCESS)
                {
                    textBox38.Text = economic_life_time_UV_lamp1.Value.ToString();
                }
                if (myPLC.ReadTag(economic_life_time_UV_lamp2) == ResultCode.E_SUCCESS)
                {
                    textBox39.Text = economic_life_time_UV_lamp2.Value.ToString();
                }
                if (myPLC.ReadTag(radiation_switch_UV1_UV2) == ResultCode.E_SUCCESS)
                {
                    textBox40.Text = radiation_switch_UV1_UV2.Value.ToString();
                }
                if (myPLC.ReadTag(automatic_switch_UV1_UV2_on_off) == ResultCode.E_SUCCESS)
                {
                    textBox41.Text = automatic_switch_UV1_UV2_on_off.Value.ToString();
                }

                //
                if (myPLC.ReadTag(Delay_time_switch_off_exhaust_air) == ResultCode.E_SUCCESS)
                {
                    textBox42.Text = Delay_time_switch_off_exhaust_air.Value.ToString();
                }
                if (myPLC.ReadTag(Delay_time_start_exhaust_air) == ResultCode.E_SUCCESS)
                {
                    textBox43.Text = Delay_time_start_exhaust_air.Value.ToString();
                }

                //UV1ActValue
                if (myPLC.ReadTag(UV1ActValue) == ResultCode.E_SUCCESS)
                {
                    textBox44.Text = UV1ActValue.Value.ToString();
                }
                if (myPLC.ReadTag(UV2ActValue) == ResultCode.E_SUCCESS)
                {
                    textBox45.Text = UV2ActValue.Value.ToString();
                }
                if (myPLC.ReadTag(UV_lamp1_operating_grade) == ResultCode.E_SUCCESS)
                {
                    textBox46.Text = UV_lamp1_operating_grade.Value.ToString();
                }
                if (myPLC.ReadTag(UVsetValue1) == ResultCode.E_SUCCESS)
                {
                    textBox47.Text = UVsetValue1.Value.ToString();
                }
                if (myPLC.ReadTag(UVsetValue2) == ResultCode.E_SUCCESS)
                {
                    textBox48.Text = UVsetValue2.Value.ToString();
                }
                if (myPLC.ReadTag(UV_lamp2_operating_grade) == ResultCode.E_SUCCESS)
                {
                    textBox49.Text = UV_lamp2_operating_grade.Value.ToString();
                }
            }
            catch (Exception) // Return when running false
            {
                MessageBox.Show("Error");
            }
        }
 
        private void Form3_Load(object sender, EventArgs e) { }
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
        private void button2_Click_1(object sender, EventArgs e) { }

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

        private void label31_Click(object sender, EventArgs e)
        {

        }
    }
}
