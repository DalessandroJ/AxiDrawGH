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

namespace AxiDrawGH
{
    public class AxiDrawGHComponent : GH_Component
    {

        public AxiDrawGHComponent()
          : base("AxiDraw Control", "AxiDraw",
              "Sends plotting commands to the AxiDraw API from inside of Grasshopper by way of a custom SVG exporter. This component launches a command prompt window which holds the information for the plotter. Once plotting, you can close Grasshopper and Rhino, but closing the command prompt window will stop (not pause) the AxiDraw.",
              "Extra", "AxiDraw")
        {
        }

        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Config", "Config", "The filepath to the AxiDraw Configuration file you want to use.", GH_ParamAccess.item);
            pManager.AddTextParameter("Filepath", "Filepath", "Folder to save the SVG file to.", GH_ParamAccess.item);
            pManager.AddTextParameter("Name", "Name", "OPTIONAL Name for the SVG file. If left blank, the file will be named with a timestamp. Do not include an extension.", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Run", "Run", "Connect a button to this, click it to execute the plot. Do not use a boolean toggle.", GH_ParamAccess.item);
            pManager.AddRectangleParameter("Border", "Border", "Rectangle to use as a border. The actual size of the paper you are plotting on.", GH_ParamAccess.item);
            pManager.AddCurveParameter("Curves", "Curves", "Curves to plot. Each branch in the tree is its own layer, accessed through its branch index.", GH_ParamAccess.tree);
            pManager.AddIntegerParameter("Layer", "Layer", "OPTIONAL layer to plot. Branch index to use in Curves datatree. If left empty, all layers are plotted together.", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Draw Speed", "Draw", "Pen down speed. The maximum speed the pen travels at when drawing. 1 - 110", GH_ParamAccess.item, 25);
            pManager.AddIntegerParameter("Move Speed", "Move", "Pen up speed. The maximum speed the pen travels at when moving in the up position. 1 - 110", GH_ParamAccess.item, 75);
            pManager.AddIntegerParameter("Acceleration", "Accel", "Acceleration value. How fast it takes to get to max speed. 1- 100", GH_ParamAccess.item, 75);
            pManager.AddIntegerParameter("Lowering Rate", "Lower", "Pen lowering rate. How fast the pen lowers. 1- 100", GH_ParamAccess.item, 50);
            pManager.AddIntegerParameter("Raising Rate", "Raise", "Pen raising rate. How fast the pen raises. 1 - 100", GH_ParamAccess.item, 75);
            pManager.AddBooleanParameter("Const. Speed", "Const.", "Constant speed makes the pen down movement always the max speed. Pen up travel is still affected by acceleration.", GH_ParamAccess.item, false);
            pManager.AddBooleanParameter("Reorder", "Reorder", "Attempts to reorder the SVG to plot more efficiently with AxiDraw's built in reordering.", GH_ParamAccess.item, false);
            pManager.AddNumberParameter("Polyline Resolution", "Poly Res", "Some curve types such as Ellipses, Arcs, and similar segments of PolyCurves are automatically converted to polylines. This is the target segment length for those curves that get automatically converted to polylines.", GH_ParamAccess.item);

            //pManager[1].Optional = true;    //filepath is optional
            pManager[2].Optional = true;    //name is optional
            pManager[6].Optional = true;    //layer is optional
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

            GH_Structure<GH_Curve> curves = new GH_Structure<GH_Curve>();
            DA.GetDataTree(5, out curves);

            int layer = new int();
            DA.GetData(6, ref layer);

            int penDownSpeed = new int();
            DA.GetData(7, ref penDownSpeed);

            int penUpSpeed = new int();
            DA.GetData(8, ref penUpSpeed);

            int acceleration = new int();
            DA.GetData(9, ref acceleration);

            int penLowerSpeed = new int();
            DA.GetData(10, ref penLowerSpeed);

            int penRaiseSpeed = new int();
            DA.GetData(11, ref penRaiseSpeed);

            bool constantSpeed = new bool();
            DA.GetData(12, ref constantSpeed);

            bool reordering = new bool();
            DA.GetData(13, ref reordering);

            Double res = new Double();
            DA.GetData(14, ref res);
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
            //check for at least one curve in any of the tree branches
            if (curves.DataCount < 1)
            {
                AddRuntimeMessage(Grasshopper.Kernel.GH_RuntimeMessageLevel.Error, "You must supply at least one curve to plot.");
                return;
            }

            //check the specified layer, if any, for any curves
            if (Params.Input[4].SourceCount != 0)
            {
                if (layer < 0 || layer >= curves.Branches.Count)
                {
                    AddRuntimeMessage(Grasshopper.Kernel.GH_RuntimeMessageLevel.Error, "Curves DataTree does not contain a layer with the specified index.");
                    return;
                }
                if (curves.Branches[layer].Count == 0)
                {
                    AddRuntimeMessage(Grasshopper.Kernel.GH_RuntimeMessageLevel.Error, "There are no curves on this layer.");
                    return;
                }
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

            if(res < 1.0/2032.0)
            {
                AddRuntimeMessage(Grasshopper.Kernel.GH_RuntimeMessageLevel.Warning, "Polyline resolution cannot be less than the AxiDraw resolution, so it was set to your Rhino document tolerance.");
                res = Rhino.RhinoDoc.ActiveDoc.ModelAbsoluteTolerance;
            }

            #endregion sanity

            //-------------------------------------------------------------------------------
            //calculate info: border coords, create svg string list, write header info to it
            //-------------------------------------------------------------------------------
            #region setup

            //get the coordinates of the x and y min and max of the border
            double bxMin = border.Corner(0).X, bxMax = border.Corner(2).X, byMin = border.Corner(0).Y, byMax = border.Corner(2).Y;

            //get the width and the height as well
            double bWidth = Math.Abs(border.Width), bHeight = Math.Abs(border.Height);

            //create string list that will eventually be used to write all lines to an svg file
            List<string> svgStrings = new List<string>();

            //add the first two lines of the header that are always the same
            svgStrings.Add("<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"no\"?>");
            svgStrings.Add("<!DOCTYPE svg PUBLIC \"-//W3C//DTD SVG 1.1//EN\" \"http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd\">");

            //calculate your remapped width and height
            int sWidth = (int) bWidth * 72;
            int sHeight = (int) bHeight * 72;

            //build and add the third line of the header that includes the calculated svg width and height
            string h3 = $"<svg version=\"1.1\" width=\"{sWidth}pt\" height=\"{sHeight}pt\" viewBox=\"0 0 {sWidth} {sHeight}\" overflow=\"visible\" xmlns=\"http://www.w3.org/2000/svg\">";
            svgStrings.Add(h3);

            #endregion setup

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
                //looping through every curve in selection and write it as SVG object in a string
                //-------------------------------------------------------------------------------

                if (Params.Input[4].SourceCount == 0)    //no layers specified, put all the curves into one plot
                {
                    for (int i = 0; i < curves.Branches.Count; i++) //for each branch of the data tree
                    {
                        for (int j = 0; j < curves.Branches[i].Count; j++) //for each GH_Curve object in the current branch
                        {
                            //WRITE THE CURVE TO THE SVG STRING LIST
                            string currentCurve = curveToSvgObj(curves.Branches[i][j], bxMin, bxMax, byMin, byMax, sWidth, sHeight, res);
                            svgStrings.Add(currentCurve);
                        }
                    }
                }
                else    //use the layer int to choose the datatree branch
                {
                    for (int i = 0; i < curves.Branches[layer].Count; i++) //for each object in the specified branch
                    {
                        //WRITE THE CURVE TO THE SVG STRING LIST
                        string currentCurve = curveToSvgObj(curves.Branches[layer][i], bxMin, bxMax, byMin, byMax, sWidth, sHeight, res);
                        svgStrings.Add(currentCurve);
                    }
                }

                svgStrings.Add("</svg>");

                //-------------------------------------------------------------------------------
                //Creating the temp / filepath here
                //-------------------------------------------------------------------------------
                #region create filepath

                //create the Filepath
                string directory;
                //if (Params.Input[1].SourceCount == 0) //if no filepath is supplied
                //{
                //    //use temporary file directory
                //    directory = System.IO.Path.GetTempPath();
                //}
                //else
                //{
                    //use the custom filepath
                    directory = customPath;
                //}

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
                //Creating the SVG here
                //-------------------------------------------------------------------------------
                File.WriteAllLines(filepath, svgStrings);

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
                    //REMOVED THIS BLOCK BECAUSE IT CAUSES WEIRD USB COMMUNICATION ERRORS TO HAVE THIS PROCESS START AND END SO QUICKLY WITH THE NEXT ONE
                    //and suposedly axicli will raise the pen initially...
                    
                    //oskay says axicli always raises pen first anyways, but not in my experience
                    //Process RaisePen = new Process();
                    //ProcessStartInfo startInfo1 = new ProcessStartInfo
                    //{
                    //    WindowStyle = ProcessWindowStyle.Hidden,
                    //    FileName = "cmd.exe",
                    //    Arguments = "/C axicli -m manual -M raise_pen"
                    //};
                    //RaisePen.StartInfo = startInfo1;
                    //RaisePen.Start();
                    //System.Threading.Thread.Sleep(500);
                    //RaisePen.Close();
                    //RaisePen.Dispose();

                    //wish i could get some animated ascii art in the cmd window or something
                    Process process = new Process();
                    ProcessStartInfo startInfo = new ProcessStartInfo
                    {
                        WindowStyle = ProcessWindowStyle.Normal,
                        FileName = "cmd.exe",
                        Arguments = cmdText
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

                //wait 2 seconds for the cli to load the svg in before deleting the temp file (IF no custom filepath was supplied)
                //if (Params.Input[1].SourceCount == 0)
                //{
                //    System.Threading.Thread.Sleep(2000);
                //    File.Delete(filepath);
                //}
            }
            DA.SetData(0, command);
        }

        private bool pending = false;

        string command = string.Empty;

        [DllImport("user32.dll")]
        static extern int SetWindowText(IntPtr hWnd, string text);

        //-------------------------------------------------------------------------------
        //helper functions
        //-------------------------------------------------------------------------------
        #region functions
        //function to remap a coordinate from y-up inchspace to y-down svgspace
        public double[] remapCoord(double inchX, double inchY, double bxMin, double bxMax, double byMin, double byMax, int sWidth, int sHeight)
        {
            //remap x
            double svgX = 0 + (inchX - bxMin) * (sWidth - 0) / (bxMax - bxMin);
            //remap y, its inverted
            double svgY = sHeight + (inchY - byMin) * (0 - sHeight) / (byMax - byMin);

            return new double[] {Math.Round(svgX, 5), Math.Round(svgY, 5)}; //round to 5 decimals, plenty
        }

        //-------------------------------------------------------------------------------

        //function to convert a GH_Curve to an SVG path string
        //this is where everything happens. i should split this into more basic functions 
        //but i am still figuring out certain curve types conversion to svg strings (arcs, ellipses)
        public string curveToSvgObj(GH_Curve crv, double bxMin, double bxMax, double byMin, double byMax, int sWidth, int sHeight, double res)
        {
            //string to build
            string p = string.Empty;

            //handle things differently if it is line/circle/polyline
            //anything else is treated as a nurbs curve, including polycurves
            if (crv.Value.IsLinear()) //it is a line
            {
                //convert to a line
                Line ln = new Line();
                GH_Convert.ToLine(crv, ref ln, GH_Conversion.Both);

                //remap start pt to svg space
                double[] startpt = remapCoord(ln.From.X, ln.From.Y, bxMin, bxMax, byMin, byMax, sWidth, sHeight);
                double[] endpt = remapCoord(ln.To.X, ln.To.Y, bxMin, bxMax, byMin, byMax, sWidth, sHeight);

                //add the svg coords to path string
                p += $"<path d=\"M{startpt[0]}, {startpt[1]} L{endpt[0]}, {endpt[1]}\"";
            }
            else if (crv.Value.IsCircle()) //it is a circle
            {
                //convert to a circle
                Circle circ = new Circle();
                GH_Convert.ToCircle(crv, ref circ, GH_Conversion.Both);

                //remap center and radius
                double[] cen = remapCoord(circ.Center.X, circ.Center.Y, bxMin, bxMax, byMin, byMax, sWidth, sHeight);
                double rad = circ.Radius * 72;

                //add the svg coords to path string
                p += $"<circle cx=\"{cen[0]}\" cy=\"{cen[1]}\" r=\"{rad}\"";
            }
            else if (crv.Value.IsPolyline()) //it can be represented as a polyline
            {
                //convert to a polyline
                crv.Value.TryGetPolyline(out Polyline poly);

                //add the beginning to the path string
                p += "<path d=\"";

                //loop through polyline points, remap coords, and add to the string
                for(int i = 0; i < poly.Count; i++)
                {
                    if (i == 0) p += "M"; //add M on the first go round
                    else p += "L"; //otherwise add L

                    //remap coords of current pt in the polyline
                    double[] currentPt = remapCoord(poly[i].X, poly[i].Y, bxMin, bxMax, byMin, byMax, sWidth, sHeight);

                    //add svg coords to path string
                    p += $"{currentPt[0]}, {currentPt[1]}";

                    if(i != poly.Count - 1 || poly.IsClosed) { p += " "; }
                }

                if (poly.IsClosed) //close path if necessary
                {
                    p += "z";
                }

                p += "\""; //add that final quotation mark to close path before attributes start
            }
            else if (crv.Value.IsArc() || crv.Value.IsEllipse())
            {
                //I just turn it into a polyline, because i havent figured this SVG arc or ellipse stuff out

                //add the beginning to the path string
                p += "<path d=\"";

                double len = crv.Value.GetLength();
                double target = res;
                int segs = (int)(len / target);

                double[] paramz = crv.Value.DivideByCount(segs + 1, true);

                for (int i = 0; i < paramz.Length; i++)
                {
                    Point3d pt = crv.Value.PointAt(paramz[i]);
                    double[] remapped = remapCoord(pt.X, pt.Y, bxMin, bxMax, byMin, byMax, sWidth, sHeight);

                    if (i == 0) p += "M"; //add M on the first go round
                    else p += "L"; //otherwise add L

                    //add svg coords to path string
                    p += $"{remapped[0]}, {remapped[1]}";

                    if (i != paramz.Length - 1 || crv.Value.IsClosed) { p += " "; }
                }

                if (crv.Value.IsClosed) //close path if necessary
                {
                    p += "z";
                }

                p += "\""; //add that final quotation mark to close path before attributes start

                #region svgarc
                //convert to an arc
                Arc arc = new Arc();
                GH_Convert.ToArc(crv, ref arc, GH_Conversion.Both);

                //get the radius, cen, start point, end point
                double rad = arc.Radius;
                Point3d cen = arc.Center;
                Point3d startPt = arc.StartPoint;
                Point3d endPt = arc.EndPoint;

                //flip world plane so you always measure clockwise angles
                Plane pln = Plane.WorldXY;
                pln.Flip();

                //SOMETHING IS GOING ON HERE vvvv

                //measure the angle of both endpoints, clockwise from x axis (like SVG do)
                Vector3d cen2start = new Vector3d(startPt - cen);
                double CWangle1 = Vector3d.VectorAngle(Vector3d.XAxis, cen2start, pln);
                Vector3d cen2end = new Vector3d(endPt - cen);
                double CWangle2 = Vector3d.VectorAngle(Vector3d.XAxis, cen2end, pln);

                //then find which is less and more
                double startAngle = Math.Min(CWangle1, CWangle2);
                double sweepAngle = Math.Max(CWangle1, CWangle2);

                double startX = cen[0] + rad * Math.Cos(startAngle);
                double startY = cen[1] + rad * Math.Sin(startAngle);
                double endX = cen[0] + rad * Math.Cos(sweepAngle);
                double endY = cen[1] + rad * Math.Sin(sweepAngle);
                int flagArc = (sweepAngle - startAngle > Math.PI) ? 1 : 0;
                int flagSweep = (sweepAngle > 0) ? 1 : 0;

                //Now(?) remap to svg space (or should i do this prior to the angle checking?)
                double[] SVGstartPt = remapCoord(startX, startY, bxMin, bxMax, byMin, byMax, sWidth, sHeight);
                double[] SVGendPt = remapCoord(endX, endY, bxMin, bxMax, byMin, byMax, sWidth, sHeight);

                //add the SVG coords to the path string
                p += "<path d=\"";
                p += $"M{SVGstartPt[0]} {SVGstartPt[1]} A {rad * 72} {rad * 72} 0 {flagArc} {flagSweep} {SVGendPt[0]} {SVGendPt[1]}\"";
                #endregion svgarc
            }
            else if (crv.Value.ToString() == "Rhino.Geometry.PolyCurve") //it is a polycurve
            {
                PolyCurve pcrv = (PolyCurve)crv.Value;

                //get the segment count of the polycurve
                int c = pcrv.SegmentCount;

                //separate each segment into the points to rebuild as a polyline
                List<Point3d> pts = new List<Point3d>();
                for (int i = 0; i < c; i++)
                {
                    Curve seg = pcrv.SegmentCurve(i);
                    if (seg.IsLinear()) //if the segment is a line
                    {
                        pts.Add(seg.PointAtStart);
                        pts.Add(seg.PointAtEnd);
                    }
                    else if (seg.IsPolyline()) //if the segment is a polyline
                    {
                        Polyline poly = new Polyline();
                        seg.TryGetPolyline(out poly);
                        pts.AddRange(poly.ToArray());
                    }
                    else //otherwise its a nurbs curve, or a arc, or an ellipse segment, so we target length divide and create a point list
                    {
                        //for arc and ellipse, i dont have a better way yet
                        //but for nurbs curve segments this is really wasteful and heavily weighs down the file with a lot of points
                        //depending on your target polyline resolution

                        double len = seg.GetLength();
                        double target = res;
                        int segs = (int)(len / target);
                        double[] paramz = seg.DivideByCount(segs + 1, true);

                        for (int j = 0; j < paramz.Length; j++)
                        {
                            pts.Add(seg.PointAt(paramz[j]));
                        }
                    }
                }

                //remove duplicates since endpts of neighboring segments are coincident
                List<Point3d> uniques = new List<Point3d>() { pts[0] };
                for (int i = 1; i < pts.Count; i++)
                {
                    if (pts[i].EpsilonEquals(pts[i - 1], Rhino.RhinoDoc.ActiveDoc.ModelAbsoluteTolerance) == false)
                    {
                        uniques.Add(pts[i]);
                    }
                }

                //construct the polyline svg string

                //add the beginning to the path string
                p += "<path d=\"";

                //remap coords and add to path string
                for (int i = 0; i < uniques.Count; i++)
                {
                    double[] remapped = remapCoord(uniques[i].X, uniques[i].Y, bxMin, bxMax, byMin, byMax, sWidth, sHeight);

                    if (i == 0) p += "M"; //add M on the first go round
                    else p += "L"; //otherwise add L

                    //add svg coords to path string
                    p += $"{remapped[0]}, {remapped[1]}";

                    if (i != uniques.Count - 1 || crv.Value.IsClosed) { p += " "; }
                }

                if (crv.Value.IsClosed) //close path if necessary
                {
                    p += "z";
                }

                p += "\""; //add that final quotation mark to close path before attributes start
            }
            else //otherwise i guess it can be represented by a nurbs curve right?
            {
                //convert to nurbs curve
                NurbsCurve nc = crv.Value.ToNurbsCurve();

                //get the bezier spans
                List<BezierCurve> bcs = new List<BezierCurve>();
                for (int i = 0; i < nc.SpanCount; i++)
                {
                    BezierCurve bc = nc.ConvertSpanToBezier(i);
                    bcs.Add(bc);
                }

                //get the control points
                List<Point2d> pts = new List<Point2d>();
                for (int i = 0; i < bcs.Count; i++)
                {
                    for(int j = 0; j < bcs[i].ControlVertexCount; j++)
                    {
                        pts.Add(bcs[i].GetControlVertex2d(j));
                    }
                }

                //remove duplicates since span endpts are coincident
                List<Point2d> uniques = new List<Point2d>() { pts[0] };
                for (int i = 1; i < pts.Count; i++)
                {
                    if (pts[i].EpsilonEquals(pts[i - 1], Rhino.RhinoDoc.ActiveDoc.ModelAbsoluteTolerance) == false)
                    {
                        uniques.Add(pts[i]);
                    }
                }

                //add the beginning to the path string
                p += "<path d=\"";

                //loop through control points, remap coords, and add to the string
                for(int i = 0; i < uniques.Count; i++)
                {
                    if (i == 0) p += "M"; //add M on the first go round

                    //otherwise add C but only every 3rd coord starting at index 1
                    //SVG is a pretty cool system that just makes sense (TM)
                    if ((i + 2) % 3 == 0) p += "C";

                    //remap coords of current control point
                    double[] currentCpt = remapCoord(uniques[i].X, uniques[i].Y, bxMin, bxMax, byMin, byMax, sWidth, sHeight);

                    //add SVG coords to path string
                    p += $"{currentCpt[0]}, {currentCpt[1]}";

                    if (i != uniques.Count - 1 || nc.IsClosed) { p += " "; }
                }

                if (nc.IsClosed) //close path if necessary
                {
                    p += "z";
                }

                p += "\""; //add that final quotation mark to close path before attributes start
            }

            //add the end attributes to path string
            p += " stroke=\"#000000\" stroke-width=\"0.25\" fill=\"none\" />";

            return p;
        }
        #endregion functions

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