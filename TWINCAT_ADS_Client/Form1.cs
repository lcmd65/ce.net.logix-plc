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
    public partial class Form1 : Form
    {
        private string plc_Address;
        private Controller myPLC = new Controller();

        // Roller Coater Machine
        private Tag Transport_set_value = new Tag("drive._010GF_100T2.set_speed");
        private Tag Transport_offset = new Tag("drive._010GF_100T2.set_offset_speed");
        private Tag Transport_actual = new Tag("drive._010GF_100T2.actual_speed_r");
        private Tag Transport_on_off = new Tag("drive._010GF_100T2.start");
        private Tag Transport_Reverse = new Tag("DB5.Coater_TR_reverse.Run");
        private Tag Applicator_roller_set_value = new Tag("drive._020RCLx_360T2.set_speed");
        private Tag Applicator_roller_offset = new Tag("drive._020RCLx_360T2.set_offset_speed");
        private Tag Applicator_roller_actual = new Tag("drive._020RCLx_360T2.actual_speed_r");
        private Tag Applicator_roller_on_off = new Tag("drive._020RCLx_360T2.start");
        private Tag Applicator_roller_Reverse = new Tag("DB5.Coater_AW_reverse.Run");
        private Tag Doctor_roller_set_value = new Tag("drive._020RCLx_370T2.set_speed");
        private Tag Doctor_roller_offset = new Tag("drive._020RCLx_370T2.set_offset_speed");
        private Tag Doctor_roller_actual = new Tag("drive._020RCLx_370T2.actual_speed_r");
        private Tag Doctor_roller_on_off = new Tag("drive._020RCLx_370T2.start");
        private Tag Doctor_roller_Reverse = new Tag("DB5.Coater_DW_reverse.Run");
        private Tag Endplate = new Tag("DB5.Plates_press.Run");
        private Tag Cleaning_tray_in_out = new Tag("DB5.Cleaning_tray.Run");
        private Tag Prepare_for_production = new Tag("DB5.MB_PrepareProd.Run");
        private Tag Roll_Cleaning = new Tag("DB5.Doctor_roller_cleaning.Run");
        private Tag RCLM_TR_Reverse = new Tag("DB5.Coater_reverse.Run");

        //Roller_Coater_ Passing Height
        private Tag Applying_roller_set_value = new Tag("L2_HA_Set");
        private Tag Applying_roller_offset = new Tag("L2_Ha_Offset");
        private Tag Applying_roller_actual = new Tag("L2_HA_Actual");
        private Tag Applying_roller_Overlift = new Tag("L2_HA_Overlift");
        private Tag Applying_roller_Diameter = new Tag("L2_HA_Diameter");
        private Tag Leading_edge = new Tag("ID1_HA_Data6");
        private Tag Trailing_edge = new Tag("ID1_HA_Data7");
        private Tag PH_Speed = new Tag("L2_HA_speed");
        private Tag Ramp_Acc = new Tag("L2_HA_Acc");
        private Tag Ramp_Dec = new Tag("L2_HA_Dec");
        private Tag Calibration = new Tag("L2_HA_Cal");
        private Tag Pos_up = new Tag("L2_HA_Pos_up");
        private Tag Pos_down = new Tag("L2_HA_Pos_down");

        //Roller_Coater_Adj_Doctor
        private Tag Roller_gap_set_value = new Tag("L2_DA_Set");
        private Tag Roller_gap_offset = new Tag("L2_DA_Offset");
        private Tag Roller_gap_actual = new Tag("L2_DA_Actual");
        private Tag Side_to_side_adj = new Tag("Side_To_Side_Adjustment");
        private Tag Adj_doctor_Speed = new Tag("L2_DA_speed");
        private Tag Adj_doctor_Ramp_Acc = new Tag("L2_DA_Acc");
        private Tag Adj_doctor_Ramp_Dec = new Tag("L2_DA_Dec");
        private Tag Adj_doctor_Calibration = new Tag("L2_DA_Cal");

        // Roller Coater -Cirulation
        private Tag Resist_level = new Tag("actual_level_arc"); // float percent
        private Tag Resist_temperature = new Tag("DB102.resist_temperature_actual"); // float 
        private Tag Manual_resist_tank_fill = new Tag("DB5.Manual_tank_fill_ARC.Set"); //binarry
        private Tag Automatic_resist_tank_fill = new Tag("DB5.Automatic_tank_fill_ARC.Set"); //binarry
        private Tag Pump_out_material_tank = new Tag("DB5.Pump_out_material.Set");  //binarry
        private Tag Clean_transport_roller = new Tag("DB5.Cleaning_TR_RCLM.Run");  //binarry
        private Tag Scraper_oscillation_roller_top = new Tag("DB5.Swifel_oszillation.Set");  //binarry
        private Tag Set_overfill_level = new Tag("DB5.Overfill_Protection.Set");  //binarry
        private Tag Set_resist_level_0 = new Tag("DB5.ARC_level_0_percent.Set");  //binarry
        private Tag Set_resist_level_100 = new Tag("DB5.ARC_level_100_percent.Set");  //binarry

        // Tag Sample Timer
        private System.Timers.Timer tagSampleTimer;
        private const double tagSampleTime = 1000;

        public Form1() 
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        // Connect to PLC
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

        // Trigger and fill data realtime from PLC signals
        private void triggerTagUpdate(object sender, ElapsedEventArgs e)
        {
            try
            {
                // Roller Coater - Machine
                if (myPLC.ReadTag(Transport_set_value) == ResultCode.E_SUCCESS)
                {
                    textBox2.Text = Transport_set_value.Value.ToString();
                    TimeStamp.Text = Transport_set_value.TimeStamp.ToString();
                }
                if (myPLC.ReadTag(Transport_offset) == ResultCode.E_SUCCESS)
                {
                    textBox3.Text = Transport_offset.Value.ToString();
                }
                if (myPLC.ReadTag(Transport_actual) == ResultCode.E_SUCCESS)
                {
                    textBox4.Text = Transport_actual.Value.ToString();
                }
                if(myPLC.ReadTag(Transport_on_off) == ResultCode.E_SUCCESS)
                {
                    textBox5.Text = Transport_on_off.Value.ToString();
                }
                if (myPLC.ReadTag(Transport_Reverse) == ResultCode.E_SUCCESS)
                {
                    textBox6.Text = Transport_Reverse.Value.ToString();
                }
                if (myPLC.ReadTag(Applicator_roller_set_value) == ResultCode.E_SUCCESS)
                {
                    textBox7.Text = Applicator_roller_set_value.Value.ToString();
                }
                if (myPLC.ReadTag(Applicator_roller_offset) == ResultCode.E_SUCCESS)
                {
                    textBox8.Text = Applicator_roller_offset.Value.ToString();
                }
                if (myPLC.ReadTag(Applicator_roller_actual) == ResultCode.E_SUCCESS)
                {
                    textBox9.Text = Applicator_roller_actual.Value.ToString();
                }
                if (myPLC.ReadTag(Applicator_roller_on_off) == ResultCode.E_SUCCESS)
                {
                    textBox10.Text = Applicator_roller_on_off.Value.ToString();
                }
                if (myPLC.ReadTag(Applicator_roller_Reverse) == ResultCode.E_SUCCESS)
                {
                    textBox11.Text = Applicator_roller_Reverse.Value.ToString();
                }
                if (myPLC.ReadTag(Doctor_roller_set_value) == ResultCode.E_SUCCESS)
                {
                    textBox12.Text = Doctor_roller_set_value.Value.ToString();
                }
                if (myPLC.ReadTag(Doctor_roller_offset) == ResultCode.E_SUCCESS)
                {
                    textBox13.Text = Doctor_roller_offset.Value.ToString();
                }
                if (myPLC.ReadTag(Doctor_roller_actual) == ResultCode.E_SUCCESS)
                {
                    textBox14.Text = Doctor_roller_actual.Value.ToString();
                }
                if (myPLC.ReadTag(Doctor_roller_on_off) == ResultCode.E_SUCCESS)
                {
                    textBox15.Text = Doctor_roller_on_off.Value.ToString();
                }
                if (myPLC.ReadTag(Doctor_roller_Reverse) == ResultCode.E_SUCCESS)
                {
                    textBox16.Text = Doctor_roller_Reverse.Value.ToString();
                }
                if (myPLC.ReadTag(Endplate) == ResultCode.E_SUCCESS)
                {
                    textBox17.Text = Endplate.Value.ToString();
                }
                if (myPLC.ReadTag(Cleaning_tray_in_out) == ResultCode.E_SUCCESS)
                {
                    textBox18.Text = Cleaning_tray_in_out.Value.ToString();
                }
                if (myPLC.ReadTag(Prepare_for_production) == ResultCode.E_SUCCESS)
                {
                    textBox19.Text = Prepare_for_production.Value.ToString();
                }
                if (myPLC.ReadTag(Roll_Cleaning) == ResultCode.E_SUCCESS)
                {
                    textBox20.Text = Roll_Cleaning.Value.ToString();
                }
                if (myPLC.ReadTag(RCLM_TR_Reverse) == ResultCode.E_SUCCESS)
                {
                    textBox21.Text = RCLM_TR_Reverse.Value.ToString();
                }


                //Roller_Coater_ Passing Height
                if (myPLC.ReadTag(Applying_roller_set_value) == ResultCode.E_SUCCESS)
                {
                    textBox22.Text = Applying_roller_set_value.Value.ToString();
                }
                if (myPLC.ReadTag(Applying_roller_offset) == ResultCode.E_SUCCESS)
                {
                    textBox23.Text = Applying_roller_offset.ToString();
                }
                if (myPLC.ReadTag(Applying_roller_actual) == ResultCode.E_SUCCESS)
                {
                    textBox24.Text = Applying_roller_actual.ToString();
                }
                if (myPLC.ReadTag(Applying_roller_Overlift) == ResultCode.E_SUCCESS)
                {
                    textBox25.Text = Applying_roller_Overlift.ToString();
                }
                if (myPLC.ReadTag(Applying_roller_Diameter) == ResultCode.E_SUCCESS)
                {
                    textBox26.Text = Applying_roller_Diameter.ToString();
                }
                if (myPLC.ReadTag(Leading_edge) == ResultCode.E_SUCCESS)
                {
                    textBox27.Text = Leading_edge.ToString();
                }
                if (myPLC.ReadTag(Trailing_edge) == ResultCode.E_SUCCESS)
                {
                    textBox28.Text = Trailing_edge.ToString();
                }
                if (myPLC.ReadTag(PH_Speed) == ResultCode.E_SUCCESS)
                {
                    textBox29.Text = PH_Speed.ToString();
                }
                if (myPLC.ReadTag(Ramp_Acc) == ResultCode.E_SUCCESS)
                {
                    textBox30.Text = Ramp_Acc.ToString();
                }
                if (myPLC.ReadTag(Ramp_Dec) == ResultCode.E_SUCCESS)
                {
                    textBox31.Text = Ramp_Dec.ToString();
                }
                if (myPLC.ReadTag(Calibration) == ResultCode.E_SUCCESS)
                {
                    textBox32.Text = Calibration.ToString();
                }
                if (myPLC.ReadTag(Pos_up) == ResultCode.E_SUCCESS)
                {
                    textBox33.Text = Pos_up.ToString();
                }
                if (myPLC.ReadTag(Pos_down) == ResultCode.E_SUCCESS)
                {
                    textBox34.Text = Pos_down.ToString();
                }

                // Adj - Doctor
                if (myPLC.ReadTag(Roller_gap_set_value) == ResultCode.E_SUCCESS)
                {
                    textBox35.Text = Roller_gap_set_value.ToString();
                }
                if (myPLC.ReadTag(Roller_gap_offset) == ResultCode.E_SUCCESS)
                {
                    textBox36.Text = Roller_gap_offset.ToString();
                }
                if (myPLC.ReadTag(Roller_gap_actual) == ResultCode.E_SUCCESS)
                {
                    textBox37.Text = Roller_gap_actual.ToString();
                }
                if (myPLC.ReadTag(Side_to_side_adj) == ResultCode.E_SUCCESS)
                {
                    textBox38.Text = Side_to_side_adj.ToString();
                }
                if (myPLC.ReadTag(Adj_doctor_Speed) == ResultCode.E_SUCCESS)
                {
                    textBox39.Text = Adj_doctor_Speed.ToString();
                }
                if (myPLC.ReadTag(Adj_doctor_Ramp_Acc) == ResultCode.E_SUCCESS)
                {
                    textBox40.Text = Adj_doctor_Ramp_Acc.ToString();
                }
                if (myPLC.ReadTag(Adj_doctor_Ramp_Dec) == ResultCode.E_SUCCESS)
                {
                    textBox41.Text = Adj_doctor_Ramp_Dec.ToString();
                }
                if (myPLC.ReadTag(Adj_doctor_Calibration) == ResultCode.E_SUCCESS)
                {
                    textBox42.Text = Adj_doctor_Calibration.ToString();
                }

                // Roller Coater -Cirulation
                if (myPLC.ReadTag(Resist_level) == ResultCode.E_SUCCESS)
                {
                    textBox43.Text = Resist_level.ToString();
                }
                if (myPLC.ReadTag(Resist_temperature) == ResultCode.E_SUCCESS)
                {
                    textBox44.Text = Resist_temperature.ToString();
                }
                if (myPLC.ReadTag(Manual_resist_tank_fill) == ResultCode.E_SUCCESS)
                {
                    textBox45.Text = Manual_resist_tank_fill.ToString();
                }
                if (myPLC.ReadTag(Automatic_resist_tank_fill) == ResultCode.E_SUCCESS)
                {
                    textBox46.Text = Automatic_resist_tank_fill.ToString();
                }
                if (myPLC.ReadTag(Pump_out_material_tank) == ResultCode.E_SUCCESS)
                {
                    textBox47.Text = Pump_out_material_tank.ToString();
                }
                if (myPLC.ReadTag(Clean_transport_roller) == ResultCode.E_SUCCESS)
                {
                    textBox48.Text = Clean_transport_roller.ToString();
                }
                if (myPLC.ReadTag(Scraper_oscillation_roller_top) == ResultCode.E_SUCCESS)
                {
                    textBox49.Text = Scraper_oscillation_roller_top.ToString();
                }
                if (myPLC.ReadTag(Set_overfill_level) == ResultCode.E_SUCCESS)
                {
                    textBox50.Text = Set_overfill_level.ToString();
                }
                if (myPLC.ReadTag(Set_resist_level_0) == ResultCode.E_SUCCESS)
                {
                    textBox51.Text = Set_resist_level_0.ToString();
                }
                if (myPLC.ReadTag(Set_resist_level_100) == ResultCode.E_SUCCESS)
                {
                    textBox52.Text = Set_resist_level_100.ToString();
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
        private void textBox1_TextChanged(object sender, EventArgs e) { }
        private void Form1_Load(object sender, EventArgs e) { }
        private void textBox72_TextChanged(object sender, EventArgs e) { }
        private void textBox2_TextChanged_1(object sender, EventArgs e) { }
        private void label2_Click(object sender, EventArgs e) { }
        private void label40_Click(object sender, EventArgs e) { }
        private void label41_Click(object sender, EventArgs e) { }
        private void label1_Click(object sender, EventArgs e) { }
        private void TimeStamp_Click(object sender, EventArgs e) { }
    }
}
