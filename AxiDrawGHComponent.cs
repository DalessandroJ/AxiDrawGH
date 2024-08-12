using System;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System.Windows.Forms.VisualStyles;
using Rhino.UI;
using System.Windows.Forms;
using System.Runtime.Remoting.Messaging;
using System.IO.Ports;
using Rhino.Display;
using Rhino.DocObjects;
using Rhino.Input.Custom;
using System.Drawing;

namespace AxiDrawGH
{
    public class AxiDrawGHComponent : GH_Component
    {

        public AxiDrawGHComponent()
          : base("AxiDraw Control", "AxiDraw",
              "Sends plotting commands to the AxiDraw API from inside of Grasshopper. This component launches a command prompt window which holds the information for the plotter. Once plotting, you can close Grasshopper and Rhino, but closing the command prompt window will stop (not pause) the AxiDraw. Important: this component will draw whatever visible curves there are.",
              "Extra", "AxiDraw")
        {
        }

        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Config", "Config", "The filepath to the AxiDraw Configuration file you want to use.", GH_ParamAccess.item);
            pManager.AddTextParameter("Filepath", "Filepath", "OPTIONAL Folder to save the SVG file to. If left blank, file will be saved to your temp folder and deleted soon after the file begins plotting.", GH_ParamAccess.item);
            pManager.AddTextParameter("Name", "Name", "OPTIONAL Name for the SVG file. If left blank, the file will be named with a timestamp. Do not include an extension.", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Run", "Run", "Connect a button to this, click it to execute the plot. When run, uses Rhino's ViewCapture.CaptureToSVG to write ANY VISIBLE CURVES in the ACTIVE VIEWPORT to a .SVG file.", GH_ParamAccess.item);
            pManager.AddRectangleParameter("Paper Rectangle", "Paper", "A rectangle the actual size and orientation of the paper you are plotting on. All the curves to plot should be inside this.", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Draw Speed", "Draw", "Pen down speed. The maximum speed the pen travels at when drawing. 1 - 110", GH_ParamAccess.item, 25);
            pManager.AddIntegerParameter("Move Speed", "Move", "Pen up speed. The maximum speed the pen travels at when moving in the up position. 1 - 110", GH_ParamAccess.item, 75);
            pManager.AddIntegerParameter("Acceleration", "Accel", "Acceleration value. How fast it takes to get to max speed. 1- 100", GH_ParamAccess.item, 75);
            pManager.AddIntegerParameter("Lowering Rate", "Lower", "Pen lowering rate. How fast the pen lowers. 1- 100", GH_ParamAccess.item, 50);
            pManager.AddIntegerParameter("Raising Rate", "Raise", "Pen raising rate. How fast the pen raises. 1 - 100", GH_ParamAccess.item, 75);
            pManager.AddBooleanParameter("Const. Speed", "Const.", "Constant speed makes the pen down movement always the max speed. Pen up travel is still affected by acceleration.", GH_ParamAccess.item, false);
            pManager.AddBooleanParameter("Reorder", "Reorder", "Attempts to reorder the SVG to plot more efficiently with AxiDraw's built in reordering.", GH_ParamAccess.item, false);

            pManager[1].Optional = true;    //filepath is optional
            pManager[2].Optional = true;    //name is optional
        }
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("AxiCli Command", "CMD", "The AxiCli command that was sent to the AxiDraw API.", GH_ParamAccess.item);
        }


        protected override void SolveInstance(IGH_DataAccess DA)
        {
            //-------------------------------------------------------------------------------
            //set up inputs
            //-------------------------------------------------------------------------------
            #region boilerplate
            string config = string.Empty;
            DA.GetData(0, ref config);

            string customPath = string.Empty;
            DA.GetData(1, ref customPath);

            string customName = string.Empty;
            DA.GetData(2, ref customName);

            bool run = new bool();
            DA.GetData(3, ref run);

            Rectangle3d border = new Rectangle3d();
            DA.GetData(4, ref border);

            int penDownSpeed = new int();
            DA.GetData(5, ref penDownSpeed);

            int penUpSpeed = new int();
            DA.GetData(6, ref penUpSpeed);

            int acceleration = new int();
            DA.GetData(7, ref acceleration);

            int penLowerSpeed = new int();
            DA.GetData(8, ref penLowerSpeed);

            int penRaiseSpeed = new int();
            DA.GetData(9, ref penRaiseSpeed);

            bool constantSpeed = new bool();
            DA.GetData(10, ref constantSpeed);

            bool reordering = new bool();
            DA.GetData(11, ref reordering);

            #endregion boilerplate

            //-------------------------------------------------------------------------------
            //sanity checks for everything that needs to be checked
            //-------------------------------------------------------------------------------
            #region sanity

            //just check for config file. checking if a filepath is valid is too annoying
            if (config == null)
            {
                AddRuntimeMessage(Grasshopper.Kernel.GH_RuntimeMessageLevel.Error, "You must supply a config file.");
                return;
            }
            //check for a border
            if (border.IsValid == false)
            {
                AddRuntimeMessage(Grasshopper.Kernel.GH_RuntimeMessageLevel.Error, "You must supply a rectangular border (Size of the paper you wish to use).");
                return;
            }
            //pendownspeed stuff
            if (penDownSpeed < 1)
            {
                AddRuntimeMessage(Grasshopper.Kernel.GH_RuntimeMessageLevel.Warning, "Pen Down Speed cannot be less than 1, so it was set to 1.");
                penDownSpeed = 1;
            }
            if (penDownSpeed > 110)
            {
                AddRuntimeMessage(Grasshopper.Kernel.GH_RuntimeMessageLevel.Warning, "Pen Down Speed cannot be more than 110, so it was set to 110.");
                penDownSpeed = 110;
            }
            //penupspeed stuff
            if (penUpSpeed < 1)
            {
                AddRuntimeMessage(Grasshopper.Kernel.GH_RuntimeMessageLevel.Warning, "Pen Up Speed cannot be less than 1, so it was set to 1.");
                penUpSpeed = 1;
            }
            if (penUpSpeed > 110)
            {
                AddRuntimeMessage(Grasshopper.Kernel.GH_RuntimeMessageLevel.Warning, "Pen Up Speed cannot be more than 110, so it was set to 110.");
                penUpSpeed = 110;
            }
            //acceleration stuff
            if (acceleration < 1)
            {
                AddRuntimeMessage(Grasshopper.Kernel.GH_RuntimeMessageLevel.Warning, "Acceleration cannot be less than 1, so it was set to 1.");
                acceleration = 1;
            }
            if (acceleration > 100)
            {
                AddRuntimeMessage(Grasshopper.Kernel.GH_RuntimeMessageLevel.Warning, "Acceleration cannot be more than 100, so it was set to 100.");
                acceleration = 100;
            }
            //penlowerspeed stuff
            if (penLowerSpeed < 1)
            {
                AddRuntimeMessage(Grasshopper.Kernel.GH_RuntimeMessageLevel.Warning, "Pen Lowering Speed cannot be less than 1, so it was set to 1.");
                penLowerSpeed = 1;
            }
            if (penLowerSpeed > 100)
            {
                AddRuntimeMessage(Grasshopper.Kernel.GH_RuntimeMessageLevel.Warning, "Pen Lowering Speed cannot be more than 100, so it was set to 100.");
                penLowerSpeed = 100;
            }
            //penraisespeed stuff
            if (penRaiseSpeed < 1)
            {
                AddRuntimeMessage(Grasshopper.Kernel.GH_RuntimeMessageLevel.Warning, "Pen Raising Speed cannot be less than 1, so it was set to 1.");
                penRaiseSpeed = 1;
            }
            if (penRaiseSpeed > 100)
            {
                AddRuntimeMessage(Grasshopper.Kernel.GH_RuntimeMessageLevel.Warning, "Pen Raising Speed cannot be more than 100, so it was set to 100.");
                penRaiseSpeed = 100;
            }

            //unpreview the Paper Rectangle source
            (this.Params.Input[4].Sources[0].Attributes.GetTopLevel.DocObject as GH_Component).Hidden = true;

            #endregion sanity

            //-------------------------------------------------------------------------------
            //RUN IT
            //-------------------------------------------------------------------------------

            //run if true
            if (!run && !pending)
            {
                return;
            }
            if (!pending)
            {
                pending = true;
                return;
            }
            pending = false;
            {

                //-------------------------------------------------------------------------------
                //Creating the temp / filepath here
                //-------------------------------------------------------------------------------
                #region create filepath

                //create the Filepath
                string directory;
                if (Params.Input[1].SourceCount == 0) //if no filepath is supplied
                {
                    //use temporary file directory
                    directory = Path.GetTempPath();
                }
                else
                {
                    //use the custom filepath
                    directory = customPath;
                }

                //create the file name
                string filename;
                if (Params.Input[2].SourceCount == 0) //if no name is supplied
                {
                    //use timestamp by default
                    filename = DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss");
                    filename += ".svg";
                }
                else
                {
                    //use the custom file name
                    filename = customName;
                    filename += ".svg";
                }

                //concatenate the filepath
                string filepath = Path.Combine(directory, filename);
                #endregion create filepath

                //-------------------------------------------------------------------------------
                //Creating the SVG via viewCapture.CaptureSVG
                //-------------------------------------------------------------------------------
                #region svg capture

                //get points for the window size
                Point3d p0 = border.Corner(0);
                Point3d p1 = border.Corner(2);

                //set DPI (idk i just use 300??) and size (to the rectangle size)
                int dpi = 300;
                Size size = new Size((int)(border.Width * dpi), (int)(border.Height * dpi));

                //get the top view (this assumes you are plotting on the XY plane). hope your active view is your top view
                RhinoView view = Rhino.RhinoDoc.ActiveDoc.Views.ActiveView;

                //create view capture settings
                ViewCaptureSettings settings = new ViewCaptureSettings(view, size, dpi);
                settings.SetWindowRect(p0, p1);

                //capture to svg
                System.Xml.XmlDocument svgxml = ViewCapture.CaptureToSvg(settings);

                //save the svg to correct path without opening a dialog
                svgxml.Save(filepath);

                #endregion svg capture

                //-------------------------------------------------------------------------------
                //applying all the options and creating the axicli command
                //-------------------------------------------------------------------------------
                #region axicli cmd

                //the command line command for axidraw that will get concatenated with all the options
                string cmdText = "axicli " + "\"" + filepath + "\"" + " --config " + "\"" + config + "\"";
                //string cmdText = "/K python -m axicli " + "\"" + filepath + "\"" + " --config " + "\"" + config + "\"";
                //i havent quite figured out what needs to prepend the command. but this works on my machine

                //penDownSpeed
                string pds = " -s" + penDownSpeed;
                cmdText = string.Concat(cmdText, pds);

                //penUpSpeed
                string pus = " -S" + penUpSpeed;
                cmdText = string.Concat(cmdText, pus);

                //acceleration
                string acc = " -a" + acceleration;
                cmdText = string.Concat(cmdText, acc);

                //penLowerSpeed
                string pls = " -r" + penLowerSpeed;
                cmdText = string.Concat(cmdText, pls);

                //penRaiseSpeed
                string prs = " -R" + penRaiseSpeed;
                cmdText = string.Concat(cmdText, prs);

                //reordering
                if (reordering == true)
                {
                    string r = " -G3";  //reorder all, ignoring groups
                    cmdText = string.Concat(cmdText, r);
                }

                //constant speed
                if (constantSpeed == true)
                {
                    string cs = " -C";
                    cmdText = string.Concat(cmdText, cs);
                }

                //reporting time elapsed
                string rt = " -T";
                cmdText = string.Concat(cmdText, rt);
                #endregion axicli cmd

                //-------------------------------------------------------------------------------
                //asking the user, starting cmd process, sending command, plotting
                //-------------------------------------------------------------------------------
                #region cmd prompt

                //have to ask this to give user an out if they accidentally ran it
                if (
                  System.Windows.Forms.MessageBox.Show
                  ("Plot it?", "AxiDrawGH",
                  System.Windows.Forms.MessageBoxButtons.YesNo,
                  System.Windows.Forms.MessageBoxIcon.None,
                  System.Windows.Forms.MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.Yes)
                {
                    //wish i could get some animated ascii art in the cmd window or something
                    Process process = new Process();
                    ProcessStartInfo startInfo = new ProcessStartInfo
                    {
                        WindowStyle = ProcessWindowStyle.Normal,
                        FileName = "cmd.exe",
                        Arguments = $"/K {cmdText}" // using /K will execute the arguments and keep window open
                    };
                    process.StartInfo = startInfo;
                    process.Start();
                    System.Threading.Thread.Sleep(100);
                    SetWindowText(process.MainWindowHandle, "PLOTTING...CLOSE THIS WINDOW TO STOP THE PLOT");
                }
                #endregion cmd prompt

                //-------------------------------------------------------------------------------
                //printing the axicli command
                //-------------------------------------------------------------------------------

                //get the axicli command
                command = cmdText;

                //wait 5 seconds for the cli to load the svg in before deleting the temp file (IF no custom filepath was supplied)
                if (Params.Input[1].SourceCount == 0)
                {
                    System.Threading.Thread.Sleep(5000);
                    File.Delete(filepath);
                }
            }

            //show the Paper Rectangle source
            (this.Params.Input[4].Sources[0].Attributes.GetTopLevel.DocObject as GH_Component).Hidden = false;

            DA.SetData(0, command);
        }

        private bool pending = false;

        string command = string.Empty;

        [DllImport("user32.dll")]
        static extern int SetWindowText(IntPtr hWnd, string text);


        //-------------------------------------------------------------------------------
        //icon and guid
        //-------------------------------------------------------------------------------
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                return Properties.Resources.AxiDrawGHIcon;
            }
        }
        public override Guid ComponentGuid
        {
            get { return new Guid("adf1257b-3db7-49c7-8e42-a8c1081d1d64"); }
        }
    }
}