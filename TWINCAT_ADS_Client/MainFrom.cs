using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Logix;
using IronXL;
using System.Timers;
using ExcelApp = Microsoft.Office.Interop.Excel;
using System.Net;
using System.Net.Mail;
using Common.Mail;


namespace TWINCAT_ADS_Client
{

    public partial class MainFrom : Form
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

        // System tag sample timer
        private System.Timers.Timer tagSampleTimer;
        private const double tagSampleTime = 1000;

        // Prototype main form
        public MainFrom()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = true;
        }
        private void Form3_Load(object sender, EventArgs e) { }
        // Method Data table
        private DataTable ReadExcel(string fileName)
        {
            WorkBook workbook = WorkBook.Load(fileName);
            //// Work with a single WorkSheet.
            ////you can pass static sheet name like Sheet1 to get that sheet
            ////WorkSheet sheet = workbook.GetWorkSheet("Sheet1");
            //You can also use workbook.DefaultWorkSheet to get default in case you want to get first sheet only
            WorkSheet sheet = workbook.DefaultWorkSheet;
            //Convert the worksheet to System.Data.DataTable
            //Boolean parameter sets the first row as column names of your table.
            return sheet.ToDataTable(true);
        }

        // Machine tab load
        private void button1_Click(object sender, EventArgs e)
        {
            Form1 from_temp1 = new Form1();
            from_temp1.ShowDialog();
        }
        // Jet Dry tab load
        private void button8_Click(object sender, EventArgs e)
        {
            Form4 from_temp4 = new Form4();
            from_temp4.ShowDialog();
        }
        // Configuration tab load
        private void button9_Click(object sender, EventArgs e)
        {
            Form5 from_temp5 = new Form5();
            from_temp5.ShowDialog();
        }
        // Manual tab load
        private void button10_Click(object sender, EventArgs e)
        {
            Form6 from_temp6 = new Form6();
            from_temp6.ShowDialog();
        }
        // Cleaning tab load
        private void button2_Click(object sender, EventArgs e)
        {
            Form2 from_temp2 = new Form2();
            from_temp2.ShowDialog();
        }

        // UV tab load
        private void button3_Click(object sender, EventArgs e)
        {
            Form3 from_temp3 = new Form3();
            from_temp3.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Create COM Objects.
            ExcelApp.Application excelApp = new ExcelApp.Application();
            if (excelApp == null) // Init false
            {
                Console.WriteLine("Excel is not installed!!");
                return;
            }
            else
            {
                //Make a File Dialog choose
                // Form file choose get PAth of this file _ file address
                OpenFileDialog openFile = new OpenFileDialog();
                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    string fileExt = Path.GetExtension(openFile.FileName); //get the file extension
                    string fileName = openFile.FileName;
                    string Path_file = Path.GetFullPath(fileName); // get file path - address
                    if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                    {
                        ExcelApp.Workbook excelBook = excelApp.Workbooks.Open(Path_file);
                        ExcelApp.Worksheet excelSheet = (ExcelApp.Worksheet)excelBook.Sheets[1];
                        ExcelApp.Range excelRange = excelSheet.Range["A1", "O400"];

                        // make variable for processing base on Range
                        int rows = excelRange.Rows.Count;
                        int cols = excelRange.Columns.Count;

                        // Return to virsulize on Grid
                        dataGridView1.RowCount = rows;
                        dataGridView1.ColumnCount = cols;

                        // POR value: E6 - (i,5) - E296
                        // Actual value: G6 - (i,7) - G296
                        // Type of checking: F6 - (i,6) - F296
                        // 287 line parameter ~ 296 excel row
                        // Function compare and fill parameter
                        triggerTag(excelRange);
                        // For loop compare POR & actual

                        for (int i = 1; i <= rows; i++)
                        {
                            // Read new line
                            for (int j = 1; j <= cols; j++)
                            {
                                // Write to cell
                                var temp = (ExcelApp.Range)excelRange.Cells[i, j];
                                if (excelRange.Cells[i, j] != null && temp.Value2 != null)
                                {
                                    dataGridView1.Rows[i - 1].Cells[j - 1].Value = temp.Value2.ToString();
                                }
                            }
                        }

                        string compare_value_type = "R";
                        for (int i = 6; i <= 366; i++) // For loop to Highlight in excel CE file.
                        {
                            // Making temp variable for getting Value2 object
                            // Type of data to Range
                            var temp_next = (ExcelApp.Range)excelRange.Cells[i, 6];
                            if (temp_next.Value2 != null)
                            {
                                if (temp_next.Value2.ToString() == compare_value_type)
                                {
                                    // Init value POR and toolactual to compare
                                    object POR_value = dataGridView1.Rows[i - 1].Cells[4].Value;
                                    object tool_value = dataGridView1.Rows[i - 1].Cells[6].Value;
                                    if (POR_value != null && tool_value != null)
                                    {
                                        if (POR_value.Equals(tool_value) == true)// Adjustment actual value
                                        {
                                            // Current Cell Red color make 
                                            ExcelApp.Range xColor = (ExcelApp.Range)excelRange.Cells[i, 7];
                                            xColor.Interior.ColorIndex = 34; // Red color
                                        }
                                        else
                                        {
                                            // Current Cell None color  make
                                            ExcelApp.Range xColor = (ExcelApp.Range)excelRange.Cells[i, 7];
                                            xColor.Interior.ColorIndex = 3; // None color
                                        }
                                    }
                                }
                                else
                                {
                                    // Current Cell None color  make
                                    ExcelApp.Range xColor = (ExcelApp.Range)excelRange.Cells[i, 7];
                                    xColor.Interior.ColorIndex = 34; // None color
                                }
                            }
                        }
                        MessageBox.Show("Done");
                        // After reading, relaase the excel project
                        // Turn off Alert saving file
                        // Saving and  Close file
                        excelApp.DisplayAlerts = false;
                        excelBook.Close(true);
                        excelApp.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                        // Send Email
                        Email email = new Email();
                        email.SendEmail("dat.lemindast@gmail.com", "CE AUTO PROGRAM", "Auto Email From NPR Roco CE PROGRAM", Path_file);
                        //email.SendEmail("NguyenThanh.Hai@fisrtsolar.com", "CE AUTO PROGRAM", "Auto Email From NPR Roco CE PROGRAM", Path_file);
                        //email.SendEmail("Thuong.Do@fisrtsolar.com", "CE AUTO PROGRAM", "Auto Email From NPR Roco CE PROGRAM", Path_file);
                        MessageBox.Show("Send Mail");
                    }
                    else // Choose false or not .xls or .xlsx file Excel template
                    {
                        MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
                    }
                }
            }
        }
        // Click to Trigger PLC signals
        private void triggerTag(ExcelApp.Range excelRange)
        {
            try
            {
                AutoClosingMessageBox.Show("Wait a minute ...", "Note", 10000);
                // Connect and fil PLC signals to Excel
                if (myPLC.ReadTag(Transport_set_value) == ResultCode.E_SUCCESS)
                {
                    ((ExcelApp.Range)excelRange.Cells[6, 7]).Value2 = Transport_set_value.Value;
                    TimeStamp.Text = Transport_set_value.TimeStamp.ToString();
                }
                if (myPLC.ReadTag(Transport_offset) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[7, 7].Value2 = Transport_offset.Value;
                }
                if (myPLC.ReadTag(Transport_actual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[8, 7].Value2 = Transport_actual.Value;
                }
                if (myPLC.ReadTag(Transport_on_off) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[9, 7].Value2 = Transport_on_off.Value;
                }
                if (myPLC.ReadTag(Transport_Reverse) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[10, 7].Value2 = Transport_Reverse.Value;
                }
                if (myPLC.ReadTag(Applicator_roller_set_value) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[11, 7].Value2 = Applicator_roller_set_value.Value;
                }
                if (myPLC.ReadTag(Applicator_roller_offset) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[12, 7].Value2 = Applicator_roller_offset.Value;
                }
                if (myPLC.ReadTag(Applicator_roller_actual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[13, 7].Value2 = Applicator_roller_actual.Value;
                }
                if (myPLC.ReadTag(Applicator_roller_on_off) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[14, 7].Value2 = Applicator_roller_on_off.Value;
                }
                if (myPLC.ReadTag(Applicator_roller_Reverse) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[15, 7].Value2 = Applicator_roller_Reverse.Value;
                }
                if (myPLC.ReadTag(Doctor_roller_set_value) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[16, 7].Value2 = Doctor_roller_set_value.Value;
                }
                if (myPLC.ReadTag(Doctor_roller_offset) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[17, 7].Value2 = Doctor_roller_offset.Value;
                }
                if (myPLC.ReadTag(Doctor_roller_actual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[18, 7].Value2 = Doctor_roller_actual.Value;
                }
                if (myPLC.ReadTag(Doctor_roller_on_off) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[19, 7].Value2 = Doctor_roller_on_off.Value;
                }
                if (myPLC.ReadTag(Doctor_roller_Reverse) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[20, 7].Value2 = Doctor_roller_Reverse.Value;
                }
                if (myPLC.ReadTag(Endplate) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[21, 7].Value2 = Endplate.Value;
                }
                if (myPLC.ReadTag(Cleaning_tray_in_out) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[22, 7].Value2 = Cleaning_tray_in_out.Value;
                }
                if (myPLC.ReadTag(Prepare_for_production) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[23, 7].Value2 = Prepare_for_production.Value;
                }
                if (myPLC.ReadTag(Roll_Cleaning) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[24, 7].Value2 = Roll_Cleaning.Value;
                }
                if (myPLC.ReadTag(RCLM_TR_Reverse) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[25, 7].Value2 = RCLM_TR_Reverse.Value;
                }

                // Passing Height
                if (myPLC.ReadTag(Applying_roller_set_value) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[26, 7].Value2 = Applying_roller_set_value.Value;
                }
                if (myPLC.ReadTag(Applying_roller_offset) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[27, 7].Value2 = Applying_roller_offset.Value;
                }
                if (myPLC.ReadTag(Applying_roller_actual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[28, 7].Value2 = Applying_roller_actual.Value;
                }
                if (myPLC.ReadTag(Applying_roller_Overlift) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[29, 7].Value2 = Applying_roller_Overlift.Value;
                }
                if (myPLC.ReadTag(Applying_roller_Diameter) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[30, 7].Value2 = Applying_roller_Diameter.Value;
                }
                if (myPLC.ReadTag(Leading_edge) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[31, 7].Value2 = Leading_edge.Value;
                }
                if (myPLC.ReadTag(Trailing_edge) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[32, 7].Value2 = Trailing_edge.Value;
                }
                if (myPLC.ReadTag(PH_Speed) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[33, 7].Value2 = PH_Speed.Value;
                }
                if (myPLC.ReadTag(Ramp_Acc) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[34, 7].Value2 = Ramp_Acc.Value;
                }
                if (myPLC.ReadTag(Ramp_Dec) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[35, 7].Value2 = Ramp_Dec.Value;
                }
                if (myPLC.ReadTag(Calibration) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[36, 7].Value2 = Calibration.Value;
                }
                if (myPLC.ReadTag(Pos_up) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[37, 7].Value2 = Pos_up.Value;
                }
                if (myPLC.ReadTag(Pos_down) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[38, 7].Value2 = Pos_down.Value;
                }

                // Adj Doctor 
                if (myPLC.ReadTag(Roller_gap_set_value) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[41, 7].Value2 = Roller_gap_set_value.Value;
                }
                if (myPLC.ReadTag(Roller_gap_offset) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[42, 7].Value2 = Roller_gap_offset.Value;
                }
                if (myPLC.ReadTag(Roller_gap_actual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[43, 7].Value2 = Roller_gap_actual.Value;
                }
                if (myPLC.ReadTag(Side_to_side_adj) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[44, 7].Value2 = Side_to_side_adj.Value;
                }
                if (myPLC.ReadTag(Adj_doctor_Speed) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[45, 7].Value2 = Adj_doctor_Speed.Value;
                }
                if (myPLC.ReadTag(Adj_doctor_Ramp_Acc) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[46, 7].Value2 = Adj_doctor_Ramp_Acc.Value;
                }
                if (myPLC.ReadTag(Adj_doctor_Ramp_Dec) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[47, 7].Value2 = Adj_doctor_Ramp_Dec.Value;
                }

                // Roller Coater - Circulation
                if (myPLC.ReadTag(Adj_doctor_Ramp_Dec) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[53, 7].Value2 = Adj_doctor_Ramp_Dec.Value;
                }
                if (myPLC.ReadTag(Resist_temperature) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[54, 7].Value2 = Resist_temperature.Value;
                }
                if (myPLC.ReadTag(Manual_resist_tank_fill) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[55, 7].Value2 = Manual_resist_tank_fill.Value;
                }
                if (myPLC.ReadTag(Automatic_resist_tank_fill) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[56, 7].Value2 = Automatic_resist_tank_fill.Value;
                }
                if (myPLC.ReadTag(Pump_out_material_tank) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[57, 7].Value2 = Pump_out_material_tank.Value;
                }
                if (myPLC.ReadTag(Clean_transport_roller) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[58, 7].Value2 = Clean_transport_roller.Value;
                }
                if (myPLC.ReadTag(Scraper_oscillation_roller_top) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[59, 7].Value2 = Scraper_oscillation_roller_top.Value;
                }
                if (myPLC.ReadTag(Set_overfill_level) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[60, 7].Value2 = Set_overfill_level.Value;
                }
                if (myPLC.ReadTag(Set_resist_level_0) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[61, 7].Value2 = Set_resist_level_0.Value;
                }
                if (myPLC.ReadTag(Set_resist_level_100) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[62, 7].Value2 = Set_resist_level_100.Value;
                }

                // Jet dry
                if (myPLC.ReadTag(Jet_Dry_Oven_Transport_set_value) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[63, 7].Value2 = Jet_Dry_Oven_Transport_set_value.Value;
                }
                if (myPLC.ReadTag(Jet_Dry_Oven_Transport_offset) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[64, 7].Value2 = Jet_Dry_Oven_Transport_offset.Value;
                }
                if (myPLC.ReadTag(Jet_Dry_Oven_Transport_actual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[65, 7].Value2 = Jet_Dry_Oven_Transport_actual.Value;
                }
                if (myPLC.ReadTag(Heating_Zone1_Set_Value) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[66, 7].Value2 = Heating_Zone1_Set_Value.Value;
                }
                if (myPLC.ReadTag(Heating_Zone2_Set_Value) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[67, 7].Value2 = Heating_Zone2_Set_Value.Value;
                }
                if (myPLC.ReadTag(Cooling_temperature) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[68, 7].Value2 = Cooling_temperature.Value;
                }
                if (myPLC.ReadTag(Jet_Dry_Oven_Transport_on_off) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[69, 7].Value2 = Jet_Dry_Oven_Transport_on_off.Value;
                }
                if (myPLC.ReadTag(Heating_Zone1_on_off) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[70, 7].Value2 = Heating_Zone1_on_off.Value;
                }
                if (myPLC.ReadTag(Heating_Zone2_on_off) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[71, 7].Value2 = Heating_Zone2_on_off.Value;
                }

                // Jet dry ventilation
                if (myPLC.ReadTag(Jet1_Exhaust) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[72, 7].Value2 = Jet1_Exhaust.Value;
                }
                if (myPLC.ReadTag(Jet1_Circulation1) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[73, 7].Value2 = Jet1_Circulation1.Value;
                }
                if (myPLC.ReadTag(Jet1_Circulation2) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[74, 7].Value2 = Jet1_Circulation2.Value;
                }
                if (myPLC.ReadTag(Jet2_Exhaust) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[75, 7].Value2 = Jet2_Exhaust.Value;
                }
                if (myPLC.ReadTag(Jet2_Circulation1) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[76, 7].Value2 = Jet2_Circulation1.Value;
                }
                if (myPLC.ReadTag(Jet2_Circulation2) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[77, 7].Value2 = Jet2_Circulation2.Value;
                }
                if (myPLC.ReadTag(JetK_Exhaust) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[78, 7].Value2 = JetK_Exhaust.Value;
                }
                if (myPLC.ReadTag(JetK_Circulation1) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[79, 7].Value2 = JetK_Circulation1.Value;
                }
                if (myPLC.ReadTag(JetK_Circulation2) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[80, 7].Value2 = JetK_Circulation2.Value;
                }
                if (myPLC.ReadTag(JetK_Exhaust2) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[81, 7].Value2 = JetK_Exhaust2.Value;
                }

                // Infeed Transport 
                if (myPLC.ReadTag(Infeed_Transport_set_value) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[82, 7].Value2 = Infeed_Transport_set_value.Value;
                }
                if (myPLC.ReadTag(Infeed_Transport_offset) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[83, 7].Value2 = Infeed_Transport_offset.Value;
                }
                if (myPLC.ReadTag(Infeed_Transport_actual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[84, 7].Value2 = Infeed_Transport_actual.Value;
                }
                if (myPLC.ReadTag(Infeed_Transport_on_off) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[85, 7].Value2 = Infeed_Transport_on_off.Value;
                }

                //Configuration - Alarm
                if (myPLC.ReadTag(Heating_Zone1_Alarm_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[92, 7].Value2 = Heating_Zone1_Alarm_Lower_Limit.Value;
                }
                if (myPLC.ReadTag(Heating_Zone1_Alarm_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[93, 7].Value2 = Heating_Zone1_Alarm_Upper_Limit.Value;
                }
                if (myPLC.ReadTag(Heating_Zone1_Warning_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[94, 7].Value2 = Heating_Zone1_Warning_Lower_Limit.Value;
                }
                if (myPLC.ReadTag(Heating_Zone1_Warning_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[95, 7].Value2 = Heating_Zone1_Warning_Upper_Limit.Value;
                }
                if (myPLC.ReadTag(Heating_Zone2_Alarm_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[96, 7].Value2 = Heating_Zone2_Alarm_Lower_Limit.Value;
                }
                if (myPLC.ReadTag(Heating_Zone2_Alarm_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[97, 7].Value2 = Heating_Zone2_Alarm_Upper_Limit.Value;
                }
                if (myPLC.ReadTag(Heating_Zone2_Warning_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[98, 7].Value2 = Heating_Zone2_Warning_Lower_Limit.Value;
                }
                if (myPLC.ReadTag(Heating_Zone2_Warning_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[99, 7].Value2 = Heating_Zone2_Warning_Upper_Limit.Value;
                }
                if (myPLC.ReadTag(Cooling_Zone_Alarm_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[100, 7].Value2 = Cooling_Zone_Alarm_Lower_Limit.Value;
                }
                if (myPLC.ReadTag(Cooling_Zone_Alarm_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[101, 7].Value2 = Cooling_Zone_Alarm_Upper_Limit.Value;
                }
                if (myPLC.ReadTag(NPR_gel_Level_Alarm_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[102, 7].Value2 = NPR_gel_Level_Alarm_Lower_Limit.Value;
                }
                if (myPLC.ReadTag(NPR_gel_Level_Alarm_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[103, 7].Value2 = NPR_gel_Level_Alarm_Upper_Limit.Value;
                }
                if (myPLC.ReadTag(NPR_gel_Level_Warning_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[104, 7].Value2 = NPR_gel_Level_Warning_Lower_Limit.Value;
                }
                if (myPLC.ReadTag(NPR_gel_Level_Warning_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[105, 7].Value2 = NPR_gel_Level_Warning_Upper_Limit.Value;
                }
                if (myPLC.ReadTag(NPR_gel_Flow_Alarm_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[106, 7].Value2 = NPR_gel_Flow_Alarm_Lower_Limit.Value;
                }
                if (myPLC.ReadTag(NPR_gel_Flow_Alarm_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[107, 7].Value2 = NPR_gel_Flow_Alarm_Upper_Limit.Value;
                }
                if (myPLC.ReadTag(NPR_gel_Flow_Warning_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[108, 7].Value2 = NPR_gel_Flow_Warning_Lower_Limit.Value;
                }
                if (myPLC.ReadTag(NPR_gel_Flow_Warning_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[109, 7].Value2 = NPR_gel_Flow_Warning_Upper_Limit.Value;
                }
                if (myPLC.ReadTag(Density_Alarm_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[110, 7].Value2 = Density_Alarm_Lower_Limit.Value;
                }
                if (myPLC.ReadTag(Density_Alarm_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[111, 7].Value2 = Density_Alarm_Upper_Limit.Value;
                }
                if (myPLC.ReadTag(Density_Warning_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[112, 7].Value2 = Density_Warning_Lower_Limit.Value;
                }
                if (myPLC.ReadTag(Density_Warning_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[113, 7].Value2 = Density_Warning_Upper_Limit.Value;
                }
                if (myPLC.ReadTag(NPR_gel_temperature_Alarm_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[114, 7].Value2 = NPR_gel_temperature_Alarm_Lower_Limit.Value;
                }
                if (myPLC.ReadTag(NPR_gel_temperature_Alarm_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[115, 7].Value2 = NPR_gel_temperature_Alarm_Upper_Limit.Value;
                }
                if (myPLC.ReadTag(NPR_gel_temperature_Warning_Lower_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[116, 7].Value2 = NPR_gel_temperature_Warning_Lower_Limit.Value;
                }
                if (myPLC.ReadTag(NPR_gel_temperature_Warning_Upper_Limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[117, 7].Value2 = NPR_gel_temperature_Warning_Upper_Limit.Value;
                }

                // Parameter timer
                if (myPLC.ReadTag(Heat_up_time) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[118, 7].Value2 = Heat_up_time.Value;
                }
                if (myPLC.ReadTag(Delay_shut_off_auto_cleaning) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[120, 7].Value2 = Delay_shut_off_auto_cleaning.Value;
                }

                // Manual
                if (myPLC.ReadTag(Pump_out_material_tank_manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[122, 7].Value2 = Pump_out_material_tank_manual.Value;
                }
                if (myPLC.ReadTag(Material_pump_manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[123, 7].Value2 = Material_pump_manual.Value;
                }
                if (myPLC.ReadTag(Waste_water_pump_manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[124, 7].Value2 = Waste_water_pump_manual.Value;
                }
                if (myPLC.ReadTag(Infeed_transport_manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[125, 7].Value2 = Infeed_transport_manual.Value;
                }
                //if (myPLC.ReadTag(Transport_UV_manual) == ResultCode.E_SUCCESS)
                //{
                //excelRange.Cells[126, 7].Value2 = Transport_UV_manual.Value;
                //}
                if (myPLC.ReadTag(Enable_lift_lower_manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[127, 7].Value2 = Enable_lift_lower_manual.Value;
                }
                if (myPLC.ReadTag(Roller_coater_transport_manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[128, 7].Value2 = Roller_coater_transport_manual.Value;
                }
                if (myPLC.ReadTag(Heating_Zone_1_Manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[129, 7].Value2 = Heating_Zone_1_Manual.Value;
                }
                if (myPLC.ReadTag(Exhaust_air_heating_zone_1_Manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[130, 7].Value2 = Exhaust_air_heating_zone_1_Manual.Value;
                }
                if (myPLC.ReadTag(Recirculating_air_1_heating_zone_1_Manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[131, 7].Value2 = Recirculating_air_1_heating_zone_1_Manual.Value;
                }
                if (myPLC.ReadTag(Recirculating_air_2_heating_zone_1_Manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[132, 7].Value2 = Recirculating_air_2_heating_zone_1_Manual.Value;
                }
                if (myPLC.ReadTag(Dryer_transport) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[133, 7].Value2 = Dryer_transport.Value;
                }
                if (myPLC.ReadTag(Heating_Zone_2_Manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[134, 7].Value2 = Heating_Zone_2_Manual.Value;
                }
                if (myPLC.ReadTag(Exhaust_air_heating_zone_2_Manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[135, 7].Value2 = Exhaust_air_heating_zone_2_Manual.Value;
                }
                if (myPLC.ReadTag(Recirculating_air_1_heating_zone_2_Manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[136, 7].Value2 = Recirculating_air_1_heating_zone_2_Manual.Value;
                }
                if (myPLC.ReadTag(Recirculating_air_2_heating_zone_2_Manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[137, 7].Value2 = Recirculating_air_2_heating_zone_2_Manual.Value;
                }
                if (myPLC.ReadTag(Exhaust_air_1_cooling_Manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[138, 7].Value2 = Exhaust_air_1_cooling_Manual.Value;
                }
                if (myPLC.ReadTag(Exhaust_air_2_cooling_Manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[139, 7].Value2 = Exhaust_air_2_cooling_Manual.Value;
                }
                if (myPLC.ReadTag(Recirculating_air_1_cooling_Manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[140, 7].Value2 = Recirculating_air_1_cooling_Manual.Value;
                }
                if (myPLC.ReadTag(Recirculating_air_2_cooling_Manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[141, 7].Value2 = Recirculating_air_2_cooling_Manual.Value;
                }
                if (myPLC.ReadTag(Scraper_oscillation_roller_top_Manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[142, 7].Value2 = Scraper_oscillation_roller_top_Manual.Value;
                }
                if (myPLC.ReadTag(Coating_Manual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[143, 7].Value2 = Coating_Manual.Value;
                }

                // Cleaning 1
                if (myPLC.ReadTag(Cleaning1_Seq1_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[148, 7].Value2 = Cleaning1_Seq1_cleaning_medium.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq1_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[149, 7].Value2 = Cleaning1_Seq1_cleaning_top.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq1_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[150, 7].Value2 = Cleaning1_Seq1_cleaning_bottom.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq1_Supply) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[151, 7].Value2 = Cleaning1_Seq1_Supply.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq1_Backflow) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[152, 7].Value2 = Cleaning1_Seq1_Backflow.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq1_duration) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[153, 7].Value2 = Cleaning1_Seq1_duration.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq2_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[154, 7].Value2 = Cleaning1_Seq2_cleaning_medium.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq2_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[155, 7].Value2 = Cleaning1_Seq2_cleaning_top.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq2_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[156, 7].Value2 = Cleaning1_Seq2_cleaning_bottom.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq2_Supply) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[157, 7].Value2 = Cleaning1_Seq2_Supply.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq2_Backflow) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[158, 7].Value2 = Cleaning1_Seq2_Backflow.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq2_duration) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[159, 7].Value2 = Cleaning1_Seq2_duration.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq3_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[160, 7].Value2 = Cleaning1_Seq3_cleaning_medium.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq3_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[161, 7].Value2 = Cleaning1_Seq3_cleaning_top.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq3_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[162, 7].Value2 = Cleaning1_Seq3_cleaning_bottom.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq3_Supply) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[163, 7].Value2 = Cleaning1_Seq3_Supply.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq3_Backflow) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[164, 7].Value2 = Cleaning1_Seq3_Backflow.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq3_duration) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[165, 7].Value2 = Cleaning1_Seq3_duration.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq4_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[166, 7].Value2 = Cleaning1_Seq4_cleaning_medium.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq4_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[167, 7].Value2 = Cleaning1_Seq4_cleaning_top.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq4_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[168, 7].Value2 = Cleaning1_Seq4_cleaning_bottom.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq4_Supply) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[169, 7].Value2 = Cleaning1_Seq4_Supply.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq4_Backflow) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[170, 7].Value2 = Cleaning1_Seq4_Backflow.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq4_duration) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[171, 7].Value2 = Cleaning1_Seq4_duration.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq5_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[172, 7].Value2 = Cleaning1_Seq5_cleaning_medium.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq5_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[173, 7].Value2 = Cleaning1_Seq5_cleaning_top.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq5_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[174, 7].Value2 = Cleaning1_Seq5_cleaning_bottom.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq5_Supply) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[175, 7].Value2 = Cleaning1_Seq5_Supply.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq5_Backflow) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[176, 7].Value2 = Cleaning1_Seq5_Backflow.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq5_duration) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[177, 7].Value2 = Cleaning1_Seq5_duration.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq6_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[178, 7].Value2 = Cleaning1_Seq6_cleaning_medium.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq6_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[179, 7].Value2 = Cleaning1_Seq6_cleaning_top.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq6_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[180, 7].Value2 = Cleaning1_Seq7_cleaning_bottom.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq6_Supply) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[181, 7].Value2 = Cleaning1_Seq6_Supply.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq6_Backflow) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[182, 7].Value2 = Cleaning1_Seq6_Backflow.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq6_duration) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[183, 7].Value2 = Cleaning1_Seq6_duration.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq7_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[184, 7].Value2 = Cleaning1_Seq7_cleaning_medium.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq7_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[185, 7].Value2 = Cleaning1_Seq7_cleaning_top.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq7_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[186, 7].Value2 = Cleaning1_Seq7_cleaning_bottom.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq7_Supply) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[187, 7].Value2 = Cleaning1_Seq7_Supply.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq7_Backflow) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[188, 7].Value2 = Cleaning1_Seq7_Backflow.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq7_duration) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[189, 7].Value2 = Cleaning1_Seq7_duration.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq8_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[190, 7].Value2 = Cleaning1_Seq8_cleaning_medium.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq8_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[191, 7].Value2 = Cleaning1_Seq8_cleaning_top.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq8_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[192, 7].Value2 = Cleaning1_Seq8_cleaning_bottom.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq8_Supply) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[193, 7].Value2 = Cleaning1_Seq8_Supply.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq8_Backflow) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[194, 7].Value2 = Cleaning1_Seq8_Backflow.Value;
                }
                if (myPLC.ReadTag(Cleaning1_Seq8_duration) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[195, 7].Value2 = Cleaning1_Seq8_duration.Value;
                }

                // Cleaning 2
                if (myPLC.ReadTag(Cleaning2_Seq1_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[200, 7].Value2 = Cleaning2_Seq1_cleaning_medium.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq1_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[201, 7].Value2 = Cleaning2_Seq1_cleaning_top.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq1_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[202, 7].Value2 = Cleaning2_Seq1_cleaning_bottom.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq1_Supply) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[203, 7].Value2 = Cleaning2_Seq1_Supply.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq1_Backflow) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[204, 7].Value2 = Cleaning2_Seq1_Backflow.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq1_duration) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[205, 7].Value2 = Cleaning2_Seq1_duration.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq2_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[206, 7].Value2 = Cleaning2_Seq2_cleaning_medium.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq2_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[207, 7].Value2 = Cleaning1_Seq2_cleaning_top.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq2_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[208, 7].Value2 = Cleaning2_Seq2_cleaning_bottom.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq2_Supply) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[209, 7].Value2 = Cleaning2_Seq2_Supply.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq2_Backflow) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[210, 7].Value2 = Cleaning2_Seq2_Backflow.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq2_duration) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[211, 7].Value2 = Cleaning2_Seq2_duration.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq3_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[212, 7].Value2 = Cleaning2_Seq3_cleaning_medium.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq3_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[213, 7].Value2 = Cleaning2_Seq3_cleaning_top.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq3_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[214, 7].Value2 = Cleaning2_Seq3_cleaning_bottom.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq3_Supply) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[215, 7].Value2 = Cleaning2_Seq3_Supply.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq3_Backflow) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[216, 7].Value2 = Cleaning2_Seq3_Backflow.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq3_duration) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[217, 7].Value2 = Cleaning2_Seq3_duration.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq4_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[218, 7].Value2 = Cleaning2_Seq4_cleaning_medium.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq4_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[219, 7].Value2 = Cleaning2_Seq4_cleaning_top.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq4_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[220, 7].Value2 = Cleaning2_Seq4_cleaning_bottom.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq4_Supply) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[221, 7].Value2 = Cleaning2_Seq4_Supply.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq4_Backflow) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[222, 7].Value2 = Cleaning2_Seq4_Backflow.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq4_duration) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[223, 7].Value2 = Cleaning2_Seq4_duration.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq5_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[224, 7].Value2 = Cleaning2_Seq5_cleaning_medium.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq5_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[225, 7].Value2 = Cleaning2_Seq5_cleaning_top.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq5_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[226, 7].Value2 = Cleaning2_Seq5_cleaning_bottom.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq5_Supply) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[227, 7].Value2 = Cleaning2_Seq5_Supply.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq5_Backflow) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[228, 7].Value2 = Cleaning2_Seq5_Backflow.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq5_duration) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[229, 7].Value2 = Cleaning2_Seq5_duration.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq6_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[230, 7].Value2 = Cleaning2_Seq6_cleaning_medium.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq6_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[231, 7].Value2 = Cleaning2_Seq6_cleaning_top.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq6_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[232, 7].Value2 = Cleaning2_Seq6_cleaning_bottom.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq6_Supply) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[233, 7].Value2 = Cleaning2_Seq6_Supply.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq6_Backflow) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[234, 7].Value2 = Cleaning2_Seq6_Backflow.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq6_duration) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[235, 7].Value2 = Cleaning2_Seq6_duration.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq7_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[236, 7].Value2 = Cleaning2_Seq7_cleaning_medium.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq7_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[237, 7].Value2 = Cleaning2_Seq7_cleaning_top.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq7_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[238, 7].Value2 = Cleaning2_Seq7_cleaning_bottom.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq7_Supply) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[239, 7].Value2 = Cleaning2_Seq7_Supply.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq7_Backflow) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[240, 7].Value2 = Cleaning2_Seq7_Backflow.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq7_duration) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[241, 7].Value2 = Cleaning2_Seq7_duration.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq8_cleaning_medium) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[242, 7].Value2 = Cleaning2_Seq8_cleaning_medium.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq8_cleaning_top) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[243, 7].Value2 = Cleaning2_Seq8_cleaning_top.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq8_cleaning_bottom) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[244, 7].Value2 = Cleaning2_Seq8_cleaning_bottom.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq8_Supply) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[245, 7].Value2 = Cleaning2_Seq8_Supply.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq8_Backflow) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[246, 7].Value2 = Cleaning2_Seq8_Backflow.Value;
                }
                if (myPLC.ReadTag(Cleaning2_Seq8_duration) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[247, 7].Value2 = Cleaning2_Seq8_duration.Value;
                }

                // UV expose MAchine UV
                if (myPLC.ReadTag(UV_Transport_set_value) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[248, 7].Value2 = UV_Transport_set_value.Value;
                }
                if (myPLC.ReadTag(UV_Transport_offset) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[249, 7].Value2 = UV_Transport_offset.Value;
                }
                if (myPLC.ReadTag(UV_Transport_actual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[250, 7].Value2 = UV_Transport_actual.Value;
                }
                if (myPLC.ReadTag(Number_of_boards_untill_next_measurement) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[251, 7].Value2 = Number_of_boards_untill_next_measurement.Value;
                }
                if (myPLC.ReadTag(Number_of_boards_untill_next_measurement_actual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[252, 7].Value2 = Number_of_boards_untill_next_measurement_actual.Value;
                }
                if (myPLC.ReadTag(UV_measurement_start) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[253, 7].Value2 = UV_measurement_start.Value;
                }
                if (myPLC.ReadTag(UV_transport_on_off) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[254, 7].Value2 = UV_transport_on_off.Value;
                }
                if (myPLC.ReadTag(automatic_measurement_on_off) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[255, 7].Value2 = automatic_measurement_on_off.Value;
                }
                if (myPLC.ReadTag(Preselection_UV1) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[256, 7].Value2 = Preselection_UV1.Value;
                }
                if (myPLC.ReadTag(Preselection_UV2) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[257, 7].Value2 = Preselection_UV2.Value;
                }

                // UV machine - Ventilation
                if (myPLC.ReadTag(UV1_Exhaust_set) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[258, 7].Value2 = UV1_Exhaust_set.Value;
                }
                if (myPLC.ReadTag(UV1_Exhaust_actual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[259, 7].Value2 = UV1_Exhaust_actual.Value;
                }
                if (myPLC.ReadTag(UV2_Exhaust_set) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[260, 7].Value2 = UV2_Exhaust_set.Value;
                }
                if (myPLC.ReadTag(UV2_Exhaust_actual) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[261, 7].Value2 = UV2_Exhaust_actual.Value;
                }

                // UV Manual General
                if (myPLC.ReadTag(UV_lamp1_on_off) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[262, 7].Value2 = UV_lamp1_on_off.Value;
                }
                if (myPLC.ReadTag(UV_lamp2_on_off) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[263, 7].Value2 = UV_lamp2_on_off.Value;
                }
                if (myPLC.ReadTag(Exhaust_air_UV_lamp1_on_off) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[264, 7].Value2 = Exhaust_air_UV_lamp1_on_off.Value;
                }
                if (myPLC.ReadTag(Exhaust_air_UV_lamp2_on_off) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[265, 7].Value2 = Exhaust_air_UV_lamp2_on_off.Value;
                }

                // UV parameter Calibration
                if (myPLC.ReadTag(teach_max_UV_lamp1_on_off) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[266, 7].Value2 = teach_max_UV_lamp1_on_off.Value;
                }
                if (myPLC.ReadTag(teach_min_UV_lamp1_on_off) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[267, 7].Value2 = teach_min_UV_lamp1_on_off.Value;
                }
                if (myPLC.ReadTag(teach_max_UV_lamp1_calibration_value) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[268, 7].Value2 = teach_max_UV_lamp1_calibration_value.Value;
                }
                if (myPLC.ReadTag(teach_min_UV_lamp1_calibration_value) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[269, 7].Value2 = teach_min_UV_lamp1_calibration_value.Value;
                }
                if (myPLC.ReadTag(UV_lamp1_sensor_max_values) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[270, 7].Value2 = UV_lamp1_sensor_max_values.Value;
                }
                if (myPLC.ReadTag(UV_lamp1_sensor_min_values) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[271, 7].Value2 = UV_lamp1_sensor_min_values.Value;
                }
                if (myPLC.ReadTag(actual_lamp1_value) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[272, 7].Value2 = actual_lamp1_value.Value;
                }
                if (myPLC.ReadTag(teach_max_UV_lamp2_on_off) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[273, 7].Value2 = teach_max_UV_lamp2_on_off.Value;
                }
                if (myPLC.ReadTag(teach_min_UV_lamp2_on_off) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[274, 7].Value2 = teach_min_UV_lamp2_on_off.Value;
                }
                if (myPLC.ReadTag(teach_max_UV_lamp2_calibration_value) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[275, 7].Value2 = teach_max_UV_lamp2_calibration_value.Value;
                }
                if (myPLC.ReadTag(teach_min_UV_lamp2_calibration_value) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[276, 7].Value2 = teach_min_UV_lamp2_calibration_value.Value;
                }
                if (myPLC.ReadTag(UV_lamp2_sensor_max_values) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[277, 7].Value2 = UV_lamp2_sensor_max_values.Value;
                }
                if (myPLC.ReadTag(UV_lamp2_sensor_min_values) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[278, 7].Value2 = UV_lamp2_sensor_min_values.Value;
                }
                if (myPLC.ReadTag(actual_lamp2_value) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[279, 7].Value2 = actual_lamp2_value.Value;
                }

                // UV Parameter Configuration
                if (myPLC.ReadTag(UV_radiation_alarm_lower_limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[280, 7].Value2 = UV_radiation_alarm_lower_limit.Value;
                }
                if (myPLC.ReadTag(UV_radiation_alarm_upper_limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[281, 7].Value2 = UV_radiation_alarm_upper_limit.Value;
                }
                if (myPLC.ReadTag(UV_radiation_warning_lower_limit) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[282, 7].Value2 = UV_radiation_warning_lower_limit.Value;
                }
                if (myPLC.ReadTag(economic_life_time_UV_lamp1) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[283, 7].Value2 = economic_life_time_UV_lamp1.Value;
                }
                if (myPLC.ReadTag(economic_life_time_UV_lamp2) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[284, 7].Value2 = economic_life_time_UV_lamp2.Value;
                }
                if (myPLC.ReadTag(radiation_switch_UV1_UV2) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[285, 7].Value2 = radiation_switch_UV1_UV2.Value;
                }
                if (myPLC.ReadTag(automatic_switch_UV1_UV2_on_off) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[286, 7].Value2 = automatic_switch_UV1_UV2_on_off.Value;
                }

                // UV Parameter Timers
                if (myPLC.ReadTag(Delay_time_switch_off_exhaust_air) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[288, 7].Value2 = Delay_time_switch_off_exhaust_air.Value;
                }
                if (myPLC.ReadTag(Delay_time_start_exhaust_air) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[289, 7].Value2 = Delay_time_start_exhaust_air.Value;
                }

                // UV Parameter Recipes
                if (myPLC.ReadTag(UV1ActValue) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[291, 7].Value2 = UV1ActValue.Value;
                }
                if (myPLC.ReadTag(UV2ActValue) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[292, 7].Value2 = UV2ActValue.Value;
                }
                if (myPLC.ReadTag(UV_lamp1_operating_grade) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[293, 7].Value2 = UV_lamp1_operating_grade.Value;
                }
                if (myPLC.ReadTag(UVsetValue1) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[294, 7].Value2 = UVsetValue1.Value;
                }
                if (myPLC.ReadTag(UVsetValue2) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[295, 7].Value2 = UVsetValue2.Value;
                }
                if (myPLC.ReadTag(UV_lamp2_operating_grade) == ResultCode.E_SUCCESS)
                {
                    excelRange.Cells[296, 7].Value2 = UV_lamp2_operating_grade.Value;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Can Trigger");
            }
        }

        // Trigger demo Tag data to get time connect
        private void triggerTagUpdate(object sender, ElapsedEventArgs e)
        {
            try
            {
                if (myPLC.ReadTag(Transport_set_value) == ResultCode.E_SUCCESS)
                {
                    TimeStamp.Text = Transport_set_value.TimeStamp.ToString();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Time null");
            }
        }

        // Connnect PLC function
        private void button5_Click(object sender, EventArgs e)
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

        // Auto trigger function
        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.DataSource != null) // Checking datasource hv already read data?
                {
                    for (int i = 6; i <= 366; i++) // For loop to Highlight in excel CE file.
                    {
                        if (dataGridView1.Rows[i - 1].Cells[4].Value != null && dataGridView1.Rows[i - 1].Cells[6].Value != null)
                        {
                            object temp1_value = dataGridView1.Rows[i - 1].Cells[4].Value;
                            object temp2_value = dataGridView1.Rows[i - 1].Cells[6].Value;
                            if (temp1_value.Equals(temp2_value) == false)// Adjustment actual value
                            {
                                dataGridView1.Rows[i - 1].Cells[6].Style.BackColor = Color.LightGreen;
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Please read CE file");
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Display Trigger False");
            }
        }

        // Button exit program
        private void Cancle_Click(object sender, EventArgs e)
        {
            if (myPLC.IsConnected == true)
            {
                myPLC.Disconnect();
            }
            this.Close();
        }

        // PLC disconnect and stop trigger timer
        private void button7_Click(object sender, EventArgs e)
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
            if (tagSampleTimer != null)
            {
                tagSampleTimer.Stop();
            }
        }

        // extra click function
        private void panel1_Paint(object sender, PaintEventArgs e) { }
        private void label9_Click(object sender, EventArgs e) { }
    }
}
// Source Code Main tab
