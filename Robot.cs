//
// (C) Copyright 2009 by Autodesk, Inc. All rights reserved.
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM 'AS IS' AND WITH ALL ITS FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE. AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using RobotOM;

namespace RobotSDKSample
{
    public partial class Robot : Form
    {
        // main reference to Robot Application object
        private IRobotApplication iapp;
        // user numbers for main structure elements
        private int startBar = 1, startNode = 1, beam1 = 2, beam2 = 3;
        // helper flags
        private bool geometryCreated = false, loadsGenerated = false;
        public Robot()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // connect to Robot
            iapp = new RobotApplicationClass();

            if (iapp.Project.IsActive == 0)
            {
                // we need an opened project to get names of available supports and bar sections
                // create a new project if not connected to existing one
                iapp.Project.New(IRobotProjectType.I_PT_FRAME_2D);
            }

            // fill combo-boxes with names of bar sections available in Robot
            IRobotNamesArray inames = iapp.Project.Structure.Labels.GetAvailableNames(IRobotLabelType.I_LT_BAR_SECTION);
            for (int i = 1; i < inames.Count; ++i)
            {
                comboColumns.Items.Add(inames.Get(i));
                comboBeams.Items.Add(inames.Get(i));
            }
            comboBeams.SelectedIndex = 0;
            comboColumns.SelectedIndex = 0;

            // fill combo-boxes with names of supports available in Robot
            inames = iapp.Project.Structure.Labels.GetAvailableNames(IRobotLabelType.I_LT_SUPPORT);
            for (int i = 1; i < inames.Count; ++i)
            {
                comboSupportLeft.Items.Add(inames.Get(i));
                comboSupportRight.Items.Add(inames.Get(i));
            }
            comboSupportLeft.SelectedIndex = 0;
            comboSupportRight.SelectedIndex = 0;
        }

        // create the structure geometry
        private void createGeometry(object sender, EventArgs e)
        {
            // switch Interactive flag off to avoid any questions that need user interaction in Robot
            iapp.Interactive = 0;

            // create a new project of type Frame 2D
            iapp.Project.New(IRobotProjectType.I_PT_FRAME_2D);

            // create nodes
            double x = 0, y = 0;
            double h = System.Convert.ToDouble(editH.Text);
            double l = System.Convert.ToDouble(editL.Text);
            double alpha = System.Convert.ToDouble(editA.Text);
            IRobotNodeServer inds = iapp.Project.Structure.Nodes;
            int n1 = startNode;
            inds.Create(n1, x, 0, y);
            inds.Create(n1 + 1, x, 0, y + h);
            inds.Create(n1 + 2, x + (l / 2), 0, y + h + Math.Tan(alpha * (Math.PI / 180)) * l / 2);
            inds.Create(n1 + 3, x + l, 0, y + h);
            inds.Create(n1 + 4, x + l, 0, 0);

            // create bars
            IRobotBarServer ibars = iapp.Project.Structure.Bars;
            int b1 = startBar;
            ibars.Create(b1, n1, n1 + 1);
            ibars.Create(b1 + 1, n1 + 1, n1 + 2); beam1 = b1 + 1;
            ibars.Create(b1 + 2, n1 + 2, n1 + 3); beam2 = b1 + 2;
            ibars.Create(b1 + 3, n1 + 3, n1 + 4);

            // set selected bar section label to columns
            RobotSelection isel = iapp.Project.Structure.Selections.Create(IRobotObjectType.I_OT_BAR);
            isel.AddOne(b1);
            isel.AddOne(b1 + 3);
            ibars.SetLabel(isel, IRobotLabelType.I_LT_BAR_SECTION, comboColumns.Text);

            // set selected bar section label to beams
            isel.Clear();
            isel.AddOne(b1 + 1);
            isel.AddOne(b1 + 2);
            ibars.SetLabel(isel, IRobotLabelType.I_LT_BAR_SECTION, comboBeams.Text);

            // set selected support label to nodes
            IRobotNode ind = (IRobotNode)inds.Get(n1);
            ind.SetLabel(IRobotLabelType.I_LT_SUPPORT, comboSupportLeft.Text);
            ind = (IRobotNode)inds.Get(n1 + 4);
            ind.SetLabel(IRobotLabelType.I_LT_SUPPORT, comboSupportRight.Text);

            geometryCreated = true;
            loadsGenerated = false;

            // switch Interactive flag on to allow user to work with Robot GUI
            iapp.Interactive = 1;

            // get the focus back
            this.Activate();
        }

        // helper method that creates a new load record in given load case
        private void create_concentrated_load(IRobotSimpleCase isc, double pos, double val)
        {
            // create a new force concentrated load record
            IRobotLoadRecord2 ilr = (IRobotLoadRecord2)isc.Records.Create(IRobotLoadRecordType.I_LRT_BAR_FORCE_CONCENTRATED);

            // define the values of load record
            ilr.SetValue((short)IRobotBarForceConcentrateRecordValues.I_BFCRV_REL, 1.0);
            ilr.SetValue((short)IRobotBarForceConcentrateRecordValues.I_BFCRV_X, pos);
            ilr.SetValue((short)IRobotBarForceConcentrateRecordValues.I_BFCRV_FZ, val);

            // apply load record to both beams of the structure
            ilr.Objects.AddOne(beam1);
            ilr.Objects.AddOne(beam2);
        }

        // generate loads
        private void generateLoads(object sender, EventArgs e)
        {
            if (!geometryCreated)
            {
                // geometry must be created first
                createGeometry(sender, e);
            }

            // switch Interactive flag off to avoid any questions that need user interaction in Robot
            iapp.Interactive = 0;

            if (loadsGenerated)
            {
                // remove all existing load cases
                RobotSelection iallCases = iapp.Project.Structure.Selections.CreateFull(IRobotObjectType.I_OT_CASE);
                iapp.Project.Structure.Cases.DeleteMany(iallCases);
            }

            // get reference to load cases server
            IRobotCaseServer icases = (IRobotCaseServer)iapp.Project.Structure.Cases;

            // get first available (free) user number for load case
            int c1 = icases.FreeNumber;

            if (checkDeadLoad.Checked)
            {
                // create dead load case
                IRobotSimpleCase isc1 = (IRobotSimpleCase)icases.CreateSimple(c1, "Dead load", IRobotCaseNature.I_CN_PERMANENT, IRobotCaseAnalizeType.I_CAT_STATIC_LINEAR);

                // define dead load record for the entire structure
                IRobotLoadRecord2 ilr1 = (IRobotLoadRecord2)isc1.Records.Create(IRobotLoadRecordType.I_LRT_DEAD);
                ilr1.SetValue((short)IRobotDeadRecordValues.I_DRV_ENTIRE_STRUCTURE, (double)1);
                ++c1;
            }

            // create live load case
            IRobotSimpleCase isc = (IRobotSimpleCase)icases.CreateSimple(c1, "Live load", IRobotCaseNature.I_CN_ACCIDENTAL, IRobotCaseAnalizeType.I_CAT_STATIC_LINEAR);
            // convert force value from kN to N and change the direction
            double val = -1000 * System.Convert.ToDouble(editIntensity1.Text);
            // add force concentrated load records in 5 points on the beams
            create_concentrated_load(isc, 0.0, 0.5 * val);
            create_concentrated_load(isc, 0.25, val);
            create_concentrated_load(isc, 0.5, val);
            create_concentrated_load(isc, 0.75, val);
            create_concentrated_load(isc, 1.0, 0.5 * val);

            // create live load case
            isc = (IRobotSimpleCase)icases.CreateSimple(c1 + 1, "Exploitation load", IRobotCaseNature.I_CN_EXPLOATATION, IRobotCaseAnalizeType.I_CAT_STATIC_LINEAR);
            // convert force value from kN to N and change the direction
            val = -1000 * System.Convert.ToDouble(editIntensity2.Text);
            // add force concentrated load records in 5 points on the beams
            create_concentrated_load(isc, 0.0, 0.5 * val);
            create_concentrated_load(isc, 0.25, val);
            create_concentrated_load(isc, 0.5, val);
            create_concentrated_load(isc, 0.75, val);
            create_concentrated_load(isc, 1.0, 0.5 * val);

            // create wind load applied to columns
            isc = (IRobotSimpleCase)icases.CreateSimple(c1 + 2, "Wind load", IRobotCaseNature.I_CN_WIND, IRobotCaseAnalizeType.I_CAT_STATIC_LINEAR);
            IRobotLoadRecord2 ilr = (IRobotLoadRecord2)isc.Records.Create(IRobotLoadRecordType.I_LRT_BAR_UNIFORM);
            ilr.SetValue((short)IRobotBarUniformRecordValues.I_BURV_PZ, 1000 * System.Convert.ToDouble(editIntensity3.Text));
            ilr.SetValue((short)IRobotBarUniformRecordValues.I_BURV_LOCAL, 1.0);
            ilr.Objects.AddOne(beam1 - 1);
            ilr.Objects.AddOne(beam2 + 1);

            // create combinations
            IRobotCaseCombination icc = (IRobotCaseCombination)icases.CreateCombination(c1 + 3, "Comb ULS", IRobotCombinationType.I_CBT_ULS, IRobotCaseNature.I_CN_EXPLOATATION, IRobotCaseAnalizeType.I_CAT_COMB);
            icc.CaseFactors.New(c1, System.Convert.ToDouble(editUls1.Text));
            icc.CaseFactors.New(c1 + 1, System.Convert.ToDouble(editUls2.Text));
            icc.CaseFactors.New(c1 + 2, System.Convert.ToDouble(editUls3.Text));

            icc = (IRobotCaseCombination)icases.CreateCombination(c1 + 4, "Comb SLS", IRobotCombinationType.I_CBT_SLS, IRobotCaseNature.I_CN_EXPLOATATION, IRobotCaseAnalizeType.I_CAT_COMB);
            icc.CaseFactors.New(c1, System.Convert.ToDouble(editSls1.Text));
            icc.CaseFactors.New(c1 + 1, System.Convert.ToDouble(editSls2.Text));
            icc.CaseFactors.New(c1 + 2, System.Convert.ToDouble(editSls3.Text));

            loadsGenerated = true;

            // switch Interactive flag on to allow user to work with Robot GUI
            iapp.Interactive = 1;

            // get the focus back
            this.Activate();
        }

        private void tabControl2_Selected(object sender, TabControlEventArgs e)
        {
            // fill combo-box with names of all load cases defined in the structure
            comboLoadCase.Items.Clear();
            IRobotCollection icases = iapp.Project.Structure.Cases.GetAll();
            for (int i = 1; i <= icases.Count; ++i)
            {
                IRobotCase ic = (IRobotCase)icases.Get(i);
                int idx = comboLoadCase.Items.Add(ic.Name);
            }
            // select the first item
            if (comboLoadCase.Items.Count > 0)
            {
                comboLoadCase.SelectedIndex = 0;
            }
        }

        private void getResults(object sender, EventArgs e)
        {
            if (iapp.Project.Structure.Results.Available == 0)
            {
                iapp.Project.CalcEngine.Calculate();
            }
            dataGridView1.Rows.Clear();
            dataGridView1.Rows.Add(12);
            dataGridView1.Rows[0].Cells[0].Value = "FX (N)";
            dataGridView1.Rows[1].Cells[0].Value = "FY (N)";
            dataGridView1.Rows[2].Cells[0].Value = "FZ (N)";
            dataGridView1.Rows[3].Cells[0].Value = "MX (Nm)";
            dataGridView1.Rows[4].Cells[0].Value = "MY (Nm)";
            dataGridView1.Rows[5].Cells[0].Value = "MZ (Nm)";
            dataGridView1.Rows[6].Cells[0].Value = "UX (m)";
            dataGridView1.Rows[7].Cells[0].Value = "UY (m)";
            dataGridView1.Rows[8].Cells[0].Value = "UZ (m)";
            dataGridView1.Rows[9].Cells[0].Value = "RX (Rad)";
            dataGridView1.Rows[10].Cells[0].Value = "RY (Rad)";
            dataGridView1.Rows[11].Cells[0].Value = "RZ (Rad)";

            int idx = comboLoadCase.SelectedIndex + 1;
            IRobotCase ic = (IRobotCase)iapp.Project.Structure.Cases.GetAll().Get(idx);
            int case_num = ic.Number;
            for (int n = startNode; n < startNode+5; ++n)
            {
                IRobotReactionData ireact = iapp.Project.Structure.Results.Nodes.Reactions.Value(n, case_num);
                IRobotDisplacementData idisp = iapp.Project.Structure.Results.Nodes.Displacements.Value(n, case_num);
                dataGridView1.Rows[0].Cells[n].Value = ireact.FX.ToString("####0.00");
                dataGridView1.Rows[1].Cells[n].Value = ireact.FY.ToString("####0.00");
                dataGridView1.Rows[2].Cells[n].Value = ireact.FZ.ToString("####0.00");
                dataGridView1.Rows[3].Cells[n].Value = ireact.MX.ToString("####0.00");
                dataGridView1.Rows[4].Cells[n].Value = ireact.MY.ToString("####0.00");
                dataGridView1.Rows[5].Cells[n].Value = ireact.MZ.ToString("####0.00");
                dataGridView1.Rows[6].Cells[n].Value = idisp.UX.ToString("####0.00");
                dataGridView1.Rows[7].Cells[n].Value = idisp.UY.ToString("####0.00");
                dataGridView1.Rows[8].Cells[n].Value = idisp.UZ.ToString("####0.00");
                dataGridView1.Rows[9].Cells[n].Value = idisp.RX.ToString("####0.00");
                dataGridView1.Rows[10].Cells[n].Value = idisp.RY.ToString("####0.00");
                dataGridView1.Rows[11].Cells[n].Value = idisp.RZ.ToString("####0.00");
            }

        
        //************************************
        // Read all bars ,test
                    IRobotApplication robApp = new RobotApplication();
            IRobotCollection bar_col = robApp.Project.Structure.Bars.GetAll();
            for (int i = 0; i < bar_col.Count; i++)
            {
                   IRobotBar bar = (IRobotBar)bar_col.Get(i+1);

                // or
               // IRobotBar bar = (IRobotBar)robApp.Project.Structure.Bars.Get(i+1);

            //}
        //-----------------------------
        //----------------------------
        // find bar nodes 
            //Declare the main variable representing the Robot application and connect it to the currently running instance of Robot.  
//Dim robapp As IRobotApplication
//Set robapp = New RobotApplication
  //Get the collection of all bars from the structure.  
//Set bar_col = robapp.Project.Structure.Bars.GetAll()
  //Iterate for consecutive bars from the collection.  
//For i = 1 To bar_col.Count
//    //Get the object representing the following (i-th) bar in the collection.    
//Set bar = bar_col.Get(i)
    //Read required attributes of the i-th bar.    
int  bar_num = bar.Number;
int start_node_num = bar.StartNode;
int end_node_num = bar.EndNode;
    //Declare variables defining individual nodes.    
//IRobotNode start_node,end_node ;
                
    //Get the node with start_node_num number from the server.    
 var start_node = robApp.Project.Structure.Nodes.Get(start_node_num);
     //Read the values of coordinates x, y, z for the node with start_node_num number.    
double start_node_x = start_node.X;     // check from internet 
double start_node_y = start_node.Y;
double start_node_z = start_node.Z;
    //Get the node with end_node_num from the server and read its coordinates.    
var end_node = robApp.Project.Structure.Nodes.Get(end_node_num);
double end_node_x = end_node.X;
double end_node_y = end_node.Y;
double end_node_z = end_node.Z;
    //Free all declared variable references.    
//Set bar = Nothing
//Set start_node = Nothing
//Set end_node = Nothing
  //Repeat the operation of reading data for the next bar from the collection.  
//Next
        //************************************
        // determine if it is beam or column 
        //************************************
        //----------------------------
                bool isBeam,isCol;

        // beam
        if (   start_node_z==end_node_z ) {
            //MY    ;
            isBeam=true;   isCol=false;
        }
        // col
        if (   start_node_z!=end_node_z )  {
            //FX    ;
            isBeam=false;   isCol=true;
        }
         //----------------------------
        // get bar length       
if (isCol)
{
                //Function ColumnHeight()
//Dim RobApp As IRobotApplication
//Set RobApp = New RobotApplication

 RobotBarServer Barserver ;
RobotBar bar2 ;
 var BarServer = robApp.Project.Structure.Bars;

int  ColNum  ;
ColNum = Val(Sheets("CHS").Range("n18"));

 var bar2 = BarServer.Get(ColNum);
Sheets("CHS").Range("n21") = bar2.Length * 1000;

  double  L=bar2.Length;
//End Function
}
if (isCol)
{
     //=================
//====================================================
// Extreme-value FZ
//====================================================
RobotExtremeParams EPFZ = new RobotExtremeParams();
EPFZ.ValueType = IRobotExtremeValueType.I_EVT_FORCE_BAR_FZ;
//EPMZ.Selection.Set(IRobotObjectType.I_OT_BAR, barEP);
//EPMZ.Selection.Set(IRobotObjectType.I_OT_CASE, loadEP);
EPFZ.BarDivision = 11;
RobotExtremeValue FZmax = robApp.Project.Structure.Results.Extremes.MaxValue(EPFZ);
RobotExtremeValue FZmin = robApp.Project.Structure.Results.Extremes.MinValue(EPFZ);
double FZ;
if (Math.Abs(FZmax.Value) > Math.Abs(FZmin.Value))
{
    FZ = FZmax.Value;
}
else
{
    FZ = FZmin.Value;
}
    FZ=FZ/1000;
MessageBox.Show(FZ.ToString(), "FZ");

//===============================
}
if( isBeam)
{
     //=================
//====================================================
// Extreme-value MY
//====================================================
RobotExtremeParams EPMY = new RobotExtremeParams();
EPMY.ValueType = IRobotExtremeValueType.I_EVT_FORCE_BAR_MY;
//EPMZ.Selection.Set(IRobotObjectType.I_OT_BAR, barEP);
//EPMZ.Selection.Set(IRobotObjectType.I_OT_CASE, loadEP);
EPMY.BarDivision = 11;
RobotExtremeValue MYmax = robApp.Project.Structure.Results.Extremes.MaxValue(EPMY);
RobotExtremeValue MYmin = robApp.Project.Structure.Results.Extremes.MinValue(EPMY);
double MY;
if (Math.Abs(MYmax.Value) > Math.Abs(MYmin.Value))
{
    MY = MYmax.Value;
}
else
{
    MY = MYmin.Value;
}
    MY=MY/1000;
MessageBox.Show(MY.ToString(), "MY");
}
//===============================
//-------------------------
    //design and calculate section by design formulas
    //************************************
    if (isBeam)     // set formulas
    {
        Mu=MY ;       
        steel beam : Zx=Mu/(phi*Fy) --->  Wply --->eshtal table  ----> IPE section
    }
    if (isCol)
    {
        Load=FZ;
    steel column : Pcr=(pi^2*E*I)/L^2=Load  --->  I=b^4/12  ---> b section
    }
        double b=;
        double h=;
//-------------------------

    //----------------------------
    //defining a steel rectangular custom section 
    //RobotLabel RLabel ;
        
    //RobotBarSectionData RLabelData ;
    RobotBarSectionNonstdData RLabelNSData ;

               string my_section = b.ToString() +"x"+ h.ToString();

     var RLabel = robApp.Project.Structure.Labels.Create(IRobotLabelType.I_LT_BAR_SECTION, my_section);
      var RLabelData = RLabel.Data;
   
    
    var RLabelData.Type = I_BST_NS_RECT;        // check from internet 
    var RLabelData.ShapeType = I_BSST_RECT_FILLED;
    
     var RLabelNSData = RLabelData.CreateNonstd(0);
    
    RLabelNSData.SetValue I_BSNDV_RECT_B,b; //0.2
    RLabelNSData.SetValue I_BSNDV_RECT_H,h; //0.4
    
    RLabelData.MaterialName = "S355";   // Selected material
    
    robApp.Project.Structure.Labels.Store (RLabel);

//----------------------------
                // introduction -------
             // create bars
            IRobotBarServer ibars = iapp.Project.Structure.Bars;
            int b1 = startBar;
            ibars.Create(b1, n1, n1 + 1);
            ibars.Create(b1 + 1, n1 + 1, n1 + 2); beam1 = b1 + 1;
            ibars.Create(b1 + 2, n1 + 2, n1 + 3); beam2 = b1 + 2;
            ibars.Create(b1 + 3, n1 + 3, n1 + 4);
    //-------------------------
    ////Read all sections
    // // fill combo-boxes with names of bar sections available in Robot
    //        IRobotNamesArray inames = iapp.Project.Structure.Labels.GetAvailableNames(IRobotLabelType.I_LT_BAR_SECTION);
    //        for (int i = 1; i < inames.Count; ++i)
    //        {
    //            comboColumns.Items.Add(inames.Get(i));
    //            comboBeams.Items.Add(inames.Get(i));
    //        }
    //        comboBeams.SelectedIndex = 0;
    //        comboColumns.SelectedIndex = 0;
    ////-------------------------
    //Assign section to element
     // set selected bar section label to columns
            RobotSelection isel = iapp.Project.Structure.Selections.Create(IRobotObjectType.I_OT_BAR);
            isel.AddOne(b1);
            isel.AddOne(b1 + 3);
            ibars.SetLabel(isel, IRobotLabelType.I_LT_BAR_SECTION, my_section);     // Which bar??
            
            //// set selected bar section label to beams
            //isel.Clear();
            //isel.AddOne(b1 + 1);
            //isel.AddOne(b1 + 2);
            //ibars.SetLabel(isel, IRobotLabelType.I_LT_BAR_SECTION, comboBeams.Text);
//-------------------------
//************************************
    }
    }
    }
}
