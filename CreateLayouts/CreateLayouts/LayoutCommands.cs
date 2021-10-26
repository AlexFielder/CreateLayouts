// (C) Copyright 2021 by Alex Fielder 
//
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace CreateLayouts
{
    public class LayoutCommands
    {
        /// <summary>
        /// Sets the plotter we want to use.
        /// </summary>
        public static string Plotterused = "";
        /// <summary>
        /// Sets the paper sheet size.
        /// </summary>
        public static string Papersize = "";
        /// <summary>
        /// Our collection of Main Viewports
        /// </summary>
        public static Dictionary<string, string> MainVports;
        /// <summary>
        /// 
        /// </summary>
        public static List<Vports> Newvports = new List<Vports>();
        /// <summary>
        /// Creates a list of new viewports that we can then add to our new Layout objects.
        /// </summary>
        /// <param name="vportitem">The Vports object that we added data to earlier.</param>
        /// <param name="ISPOSP">ISP or OSP</param>
        /// <param name="BlockId">The ObjectId of the block we're using</param>
        /// <param name="BlockNo">The Block Number we're using</param>
        /// <param name="UCSName">The name of the UCS this block is using.</param>
        /// <returns>Returns a list of New Viewports populated from the inserted blocks.</returns>
        public static List<Vports> SetupNewViewports(Vports vportitem, int ISPOSP, ObjectId BlockId, int BlockNo, string UCSName)
        {
            /* Each viewport added to the Othervports list has the blockid of the block (and therefore sheet) that it relates to.
             It should be relatively easy now to add the correct viewports to the correct sheet! 
             */
            string FullorPartial = vportitem.FullorPartial;
            string whatfloor = vportitem.FloorName;
            string VisProp = vportitem.vpVisProp;
            #region "Keyplan, North Arrow, Address"

            Vports vport = new Vports();
            vport.vpName = "KP";
            vport.FloorName = vportitem.FloorName;
            vport.vpNumber = BlockNo;
            vport.vpScale = "1";
            vport.vpPSCentrePoint = new Point3d(1137.5, 65, 0);
            vport.vpHeight = 40;
            vport.vpWidth = 53;
            vport.blockid = BlockId;
            Newvports.Add(vport);
            vport = null; // need to empty vport after each viewport definition.
            // north arrow.
            vport = new Vports(); // and then reassign it.
            vport.vpName = "NA";
            vport.vpScale = "1";
            vport.vpPSCentrePoint = new Point3d(1091, 65, 0);
            vport.vpHeight = 40;
            vport.vpWidth = 40;
            vport.vpMSCentrePoint = new Point2d(21536134.8934, 405722.8536);
            vport.blockid = BlockId;
            Newvports.Add(vport);
            vport = null;
            // Address viewport.
            vport = new Vports(); // and then reassign it.
            vport.vpName = "AD";
            vport.vpScale = "1";
            vport.vpPSCentrePoint = new Point3d(1063.6, 15, 0);
            vport.vpHeight = 20;
            vport.vpWidth = 51.5;
            vport.vpMSCentrePoint = new Point2d(21536253.2646, 406428.7391);
            vport.blockid = BlockId;
            Newvports.Add(vport);
            vport = null;
            #endregion
            #region "BoM & Notes"

            if (VisProp == "A3")
            {
                if (FullorPartial == "DC")
                {
                    // A3 Notes Viewport.
                    vport = new Vports();
                    vport.vpName = "A3 Notes Viewport";
                    vport.vpScale = "1";
                    vport.vpPSCentrePoint = new Point3d(1117.5, 240, 0);
                    vport.vpHeight = 70;
                    vport.vpWidth = 93;
                    if (ISPOSP == 1)
                    {
                        vport.vpMSCentrePoint = AssignFloorView(whatfloor);
                    }
                    vport.vpVisProp = VisProp;
                    vport.blockid = BlockId;
                    Newvports.Add(vport);
                    vport = null;

                    // A3 Annex E Viewport.
                    vport = new Vports();
                    vport.vpName = "A3 Annex E Viewport";
                    vport.vpScale = "1";
                    vport.vpPSCentrePoint = new Point3d(1117.5, 158, 0);
                    vport.vpHeight = 80;
                    vport.vpWidth = 93;
                    if (ISPOSP == 1) // 1 for ISP
                    {
                        vport.vpMSCentrePoint = AssignFloorView(whatfloor);
                    }
                    vport.vpVisProp = VisProp;
                    vport.blockid = BlockId;
                    Newvports.Add(vport);
                    vport = null;
                }
                else if (FullorPartial == "ELEVATION" | FullorPartial == "PARTIAL")
                {
                    // A3 Elevation Notes Viewport.
                    vport = new Vports();
                    vport.vpName = "A3 Notes Viewport";
                    vport.vpScale = "1";
                    vport.vpPSCentrePoint = new Point3d(1117.5, 240, 0);
                    vport.vpHeight = 70;
                    vport.vpWidth = 93;
                    if (ISPOSP == 1)
                    {
                        vport.vpMSCentrePoint = new Point2d(21537874.5, 401500);
                    }
                    else
                    {
                        vport.vpMSCentrePoint = new Point2d(21537765.1392, 399655.319);
                    }
                    vport.vpVisProp = VisProp;
                    vport.blockid = BlockId;
                    Newvports.Add(vport);
                    vport = null;
                }
                else
                {
                    // A3 Notes Viewport.
                    vport = new Vports();
                    vport.vpName = "A3 Notes Viewport";
                    vport.vpScale = "1";
                    vport.vpPSCentrePoint = new Point3d(1117.5, 240, 0);
                    vport.vpHeight = 70;
                    vport.vpWidth = 93;
                    if (ISPOSP == 1)
                    {
                        vport.vpMSCentrePoint = AssignFloorView(whatfloor);
                    }
                    else
                    {
                        vport.vpMSCentrePoint = new Point2d(21537797.3077, 399658.8806);
                    }
                    vport.vpVisProp = VisProp;
                    vport.blockid = BlockId;
                    Newvports.Add(vport);
                    vport = null;

                    // A3 Annex E Viewport.
                    vport = new Vports();
                    vport.vpName = "A3 Annex E Viewport";
                    vport.vpScale = "1";
                    vport.vpPSCentrePoint = new Point3d(1117.5, 158, 0);
                    vport.vpHeight = 80;
                    vport.vpWidth = 93;
                    if (ISPOSP == 1) // 1 for ISP
                    {
                        vport.vpMSCentrePoint = AssignFloorView(whatfloor);
                    }
                    else
                    {
                        vport.vpMSCentrePoint = new Point2d(21542213.0171, 405527.7864);
                    }
                    vport.vpVisProp = VisProp;
                    vport.blockid = BlockId;
                    Newvports.Add(vport);
                    vport = null;
                }
            }
            else if (VisProp == "A2")
            {
                // A0 Notes Viewport.
                vport = new Vports();
                vport.vpName = "A2 Notes Viewport";
                vport.vpScale = "1";
                vport.vpPSCentrePoint = new Point3d(1117.5, 363, 0);
                vport.vpHeight = 70;
                vport.vpWidth = 93;
                if (ISPOSP == 1)
                {
                    vport.vpMSCentrePoint = new Point2d(21537874.5, 401470.484);
                }
                else
                {
                    vport.vpMSCentrePoint = new Point2d(21537765.1392, 399655.319);
                }
                vport.vpVisProp = VisProp;
                vport.blockid = BlockId;
                Newvports.Add(vport);
                vport = null;

                // A0 Annex E Viewport.
                vport = new Vports();
                vport.vpName = "A2 Annex E Viewport";
                vport.vpScale = "1";
                vport.vpPSCentrePoint = new Point3d(1117.5, 281, 0);
                vport.vpHeight = 80;
                vport.vpWidth = 93;
                if (ISPOSP == 1) // 1 for ISP
                {
                    vport.vpMSCentrePoint = new Point2d(21537823.894, 405519.6273);
                }
                else
                {
                    vport.vpMSCentrePoint = new Point2d(21542202.8471, 405522.6326);
                }
                vport.vpVisProp = VisProp;
                vport.blockid = BlockId;
                Newvports.Add(vport);
                vport = null;
            }
            else if (VisProp == "A1")
            {
                // A0 Notes Viewport.
                vport = new Vports();
                vport.vpName = "A1 Notes Viewport";
                vport.vpScale = "1";
                vport.vpPSCentrePoint = new Point3d(1117.5, 537, 0);
                vport.vpHeight = 70;
                vport.vpWidth = 93;
                if (ISPOSP == 1)
                {
                    vport.vpMSCentrePoint = new Point2d(21537874.5, 401470.484);
                }
                else
                {
                    vport.vpMSCentrePoint = new Point2d(21537765.1392, 399655.319);
                }
                vport.vpVisProp = VisProp;
                vport.blockid = BlockId;
                Newvports.Add(vport);
                vport = null;

                // A0 Annex E Viewport.
                vport = new Vports();
                vport.vpName = "A1 Annex E Viewport";
                vport.vpScale = "1";
                vport.vpPSCentrePoint = new Point3d(1117.5, 455, 0);
                vport.vpHeight = 80;
                vport.vpWidth = 93;
                if (ISPOSP == 1) // 1 for ISP
                {
                    vport.vpMSCentrePoint = new Point2d(21537823.894, 405519.6273);
                }
                else
                {
                    vport.vpMSCentrePoint = new Point2d(21542202.8471, 405522.6326);
                }
                vport.vpVisProp = VisProp;
                vport.blockid = BlockId;
                Newvports.Add(vport);
                vport = null;
            }
            else if (VisProp == "A0")
            {
                // A0 Notes Viewport.
                vport = new Vports();
                vport.vpName = "A0 Notes Viewport";
                vport.vpScale = "1";
                vport.vpPSCentrePoint = new Point3d(1117.5, 784, 0);
                vport.vpHeight = 70;
                vport.vpWidth = 93;
                if (ISPOSP == 1)
                {
                    vport.vpMSCentrePoint = new Point2d(21537874.5, 401470.484);
                }
                else
                {
                    vport.vpMSCentrePoint = new Point2d(21537765.1392, 399655.319);
                }
                vport.vpVisProp = VisProp;
                vport.blockid = BlockId;
                Newvports.Add(vport);
                vport = null;

                // A0 Annex E Viewport.
                vport = new Vports();
                vport.vpName = "A0 Annex E Viewport";
                vport.vpScale = "1";
                vport.vpPSCentrePoint = new Point3d(1117.5, 702, 0);
                vport.vpHeight = 80;
                vport.vpWidth = 93;
                if (ISPOSP == 1) // 1 for ISP
                {
                    vport.vpMSCentrePoint = new Point2d(21537823.894, 405519.6273);
                }
                else
                {
                    vport.vpMSCentrePoint = new Point2d(21542202.8471, 405522.6326);
                }
                vport.vpVisProp = VisProp;
                vport.blockid = BlockId;
                Newvports.Add(vport);
                vport = null;
            }
            #endregion
            #region "Main Viewport"
            if (BlockId == vportitem.blockid)
            {
                vport = new Vports();
                vport.vpName = "Main";
                vport.vpNumber = BlockNo;
                vport.blockid = BlockId;
                vport.vpVisProp = VisProp;
                vport.Xaxis = vportitem.Xaxis;
                vport.Yaxis = vportitem.Yaxis;
                vport.UCSName = UCSName;
                if (VisProp == "A0")
                {
                    vport.vpPSCentrePoint = new Point3d(538, 435.625, 0);
                    vport.vpHeight = 766;
                    vport.vpWidth = 1066;
                }
                else if (VisProp == "A1")
                {
                    vport.vpPSCentrePoint = new Point3d(702, 304.5, 0);
                    vport.vpHeight = 519;
                    vport.vpWidth = 718;
                }
                else if (VisProp == "A2")
                {
                    vport.vpPSCentrePoint = new Point3d(835, 217.5, 0);
                    vport.vpHeight = 345;
                    vport.vpWidth = 471;
                }
                else if (VisProp == "A3")
                {
                    vport.vpPSCentrePoint = new Point3d(922.5, 156, 0);
                    vport.vpHeight = 222;
                    vport.vpWidth = 297;
                }
                Newvports.Add(vport);
            }
            #endregion
            return Newvports;
        }
        private static Point2d AssignFloorView(string whatfloor)
        {
            throw new NotImplementedException();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="finished"></param>
        public static void DeleteAllLayouts(Boolean finished)
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database tempDb = doc.Database;
            LayoutManager tempLoMan = LayoutManager.Current;
            using (Transaction Trans = tempDb.TransactionManager.StartTransaction())
            {
                DBDictionary LoDict = (DBDictionary)Trans.GetObject(tempDb.LayoutDictionaryId, OpenMode.ForRead);
                foreach (DictionaryEntry de in LoDict)
                {
                    if (de.Key.ToString() == "Model")
                    {
                        break;
                    }
                    else if (de.Key.ToString() == "Layout1")
                    {
                        if (finished == true)
                        {
                            tempLoMan.DeleteLayout(de.Key.ToString());
                        }
                    }
                    else if (de.Key.ToString() != "Layout1")
                    {
                        if (finished == false)
                        {
                            tempLoMan.DeleteLayout(de.Key.ToString());
                        }
                    }
                }
                Trans.Commit(); //if we don't commit, then the transaction doesn't do anything!
            }
        }
        /// <summary>
        /// Adds our new viewports to each freshly created layout.
        /// </summary>
        /// <param name="LayoutName">the name of the new Layout to add viewports to.</param>
        /// <param name="BlockId">the ObjectId of the Paperspace block we're adding viewports to.</param>
        /// <param name="newvport">A Vports object containing everything we captured to correctly create/position our new viewport</param>
        public static void AddNewVports(String LayoutName, ObjectId BlockId, List<Vports> newvport)
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            //Boolean FirstTime = true;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable blktbl = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord blktblrec =
                    (BlockTableRecord)tr.GetObject(blktbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite);

                foreach (ObjectId oid in blktbl)
                {
                    BlockTableRecord oblktblrec = (BlockTableRecord)tr.GetObject(oid, OpenMode.ForWrite);
                    if (oblktblrec.IsLayout)
                    {
                        Layout olayout = (Layout)tr.GetObject(oblktblrec.LayoutId, OpenMode.ForWrite);
                        if (olayout.LayoutName == LayoutName)
                        {
                            #region "Setup New Viewports"

                            foreach (Vports Stdvport in newvport)
                            {
                                if (Stdvport.blockid == BlockId)
                                {
                                    //go to newest layout using an AutoCAD Variable
                                    AcadApp.SetSystemVariable("TILEMODE", 0);
                                    // then change to the one we actually need.
                                    LayoutManager lytmgr = LayoutManager.Current;
                                    lytmgr.CurrentLayout = LayoutName;
                                    //switch to paperspace
                                    ed.SwitchToPaperSpace();
                                    //make a new viewport
                                    Viewport acVport = new Viewport();
                                    acVport.SetDatabaseDefaults();
                                    acVport.CenterPoint = Stdvport.vpPSCentrePoint;
                                    acVport.Width = Stdvport.vpWidth;
                                    acVport.Height = Stdvport.vpHeight;
                                    acVport.ViewCenter = Stdvport.vpMSCentrePoint;
                                    Double ScaleVal = Convert.ToDouble(Stdvport.vpScale);
                                    acVport.CustomScale = ScaleVal;
                                    // Add the new object to the block table record and the transaction
                                    blktblrec.AppendEntity(acVport);
                                    tr.AddNewlyCreatedDBObject(acVport, true);
                                    // Change the view direction 
                                    acVport.ViewDirection = new Vector3d(0, 0, 1);
                                    // Enable the viewport
                                    acVport.On = true;
                                    // zoom to extents!
                                    db.UpdateExt(true);
                                    // add something here to determine which keyplan we need based on the Stdvport.Floorname variable!
                                    // Activate model space in the viewport
                                    if (Stdvport.vpName == "KP")
                                    {
                                        //the next line results in ecannotchangeactiveviewport error.
                                        ed.SwitchToModelSpace();
                                        // Set the new viewport current via an imported ObjectARX function
                                        acedSetCurrentVPort(acVport.UnmanagedObject);
                                        // will zoom to each keyplan in turn!
                                        ObjectId kpid = (from KeyValuePair<string, ObjectId> kvp in MyCommands.kpdict
                                                         where kvp.Key.Contains(Convert.ToString(Stdvport.FloorName))
                                                         select kvp.Value).Single();
                                        ZoomCommands.ZoomToEntity(kpid);
                                        //switch back to Paperspace
                                        doc.Editor.SwitchToPaperSpace();
                                    }
                                    else if (Stdvport.vpName == "Main")
                                    {
                                        ed.SwitchToModelSpace();
                                        //set the new viewport current!
                                        acedSetCurrentVPort(acVport.UnmanagedObject);
                                        //attempting a new approach:
                                        ViewportTable vporttbl = (ViewportTable)tr.GetObject(db.ViewportTableId, OpenMode.ForRead);
                                        ViewportTableRecord acvporttblrec =
                                            (ViewportTableRecord)tr.GetObject(vporttbl["*Active"], OpenMode.ForWrite);
                                        acvporttblrec.SetUcs(UCSTools.GetCurrentUCSId(Stdvport.UCSName));
                                        UcsTableRecord vportcurrentucs = (UcsTableRecord)tr.GetObject(acvporttblrec.UcsName, OpenMode.ForRead);
                                        //This shows the correct ucs name but it doesn't affect the viewport!
                                        AcadApp.ShowAlertDialog("The current UCS is: " + vportcurrentucs.Name);
                                        acedVportTableRecords2Vports();
                                        //this next line might do what we need - we just need to figure out what doubles to use!
                                        acVport.ViewDirection = new Vector3d(Stdvport.Xaxis.X, Stdvport.Yaxis.Y, 1);
                                        //add the ucs change to the transaction queue
                                        tr.TransactionManager.QueueForGraphicsFlush();
                                        //regenerate the view
                                        ed.Regen();
                                        //Zoom the new viewport to the correct modelspace location.
                                        ZoomCommands.ZoomToEntity(Stdvport.blockid);
                                        //switch back to paperspace
                                        ed.SwitchToPaperSpace();
                                    }
                                    acVport = null;
                                }
                            }
                            #endregion
                            //remove the first viewport that was created on each layout.
                            //RemoveExistingViewport(oblktblrec.LayoutId);
                        }

                    }
                }
                tr.Commit();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="jobdetail"></param>
        /// <param name="BlockId"></param>
        /// <param name="newvport"></param>
        public static void CreateNewLayouts(List<Jobdetails> jobdetail, ObjectId BlockId, List<Vports> newvport)
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            using (DocumentLock doclock = doc.LockDocument())
            {
                Database db = doc.Database;
                //call GetBlockId twice otherwise we get a fatal error!
                ObjectId block = BlockHelperClass.GetBlockId(db, "Drawing Border");
                block = BlockHelperClass.GetBlockId(db, "Drawing Border");
                ObjectId LayoutblockId = ObjectId.Null;
                Layouts LayoutList = new Layouts();
                //ObjectId ltid = ObjectId.Null;
                Autodesk.AutoCAD.DatabaseServices.TransactionManager trm = db.TransactionManager;

                // this whole routine can be slimmed/refined using LINQ!
                using (Transaction tr = trm.StartTransaction())
                {
                    String layoutname = (from Jobdetails jdetails in jobdetail
                                         select jdetails.SiteNo + "-" + jdetails.BuildingNo + "-01").First();
                    //int vpno = (from Vports vp in newvport
                    //            where vp.vpName == "Main"
                    //                select vp.vpNumber).Single();
                    int vpno = (from Vports vp in newvport
                                where vp.blockid == BlockId && vp.vpName == "Main"
                                select vp.vpNumber).Single();
                    string Visprop = (from Vports vp in newvport
                                      where vp.blockid == BlockId && vp.vpName == "Main"
                                      select vp.vpVisProp).Single();
                    layoutname = String.Concat(layoutname, " ", vpno.ToString(), "of", MyCommands.blockdict.Count);
                    LayoutManager lman = LayoutManager.Current;
                    try
                    {
                        InsertBlockJig BlkClass = new InsertBlockJig();
                        /*we were checking that the KeyValuePair Value matched the BlockId - 
                         But as we're now using LINQ to get only the blockid we require, we no longer need to do this!
                         */
                        ObjectId ltid = lman.CreateLayout(layoutname);
                        Layout lt = (Layout)tr.GetObject(ltid, OpenMode.ForWrite);
                        {
                            lt.UpgradeOpen();
                            lt.Initialize();
                            //change to the one we need.
                            LayoutManager lytmgr = LayoutManager.Current;
                            lytmgr.CurrentLayout = layoutname;
                            //begin configuring plot settings
                            PlotInfo plotinf = new PlotInfo();
                            plotinf.Layout = lt.ObjectId;
                            PlotSettings plotset = new PlotSettings(lt.ModelType);
                            plotset.CopyFrom(lt);
                            PlotSettingsValidator plotsetvdr = PlotSettingsValidator.Current;
                            PlotInfoValidator plotvdr = new PlotInfoValidator();
                            Plotterused = GetPlotter(Visprop);
                            Papersize = GetPaper(Visprop);
                            plotsetvdr.SetPlotConfigurationName(plotset, Plotterused, Papersize);
                            plotsetvdr.SetPlotType(plotset, Autodesk.AutoCAD.DatabaseServices.PlotType.Extents);
                            plotsetvdr.SetUseStandardScale(plotset, true);
                            plotsetvdr.SetPlotRotation(plotset, PlotRotation.Degrees090);
                            plotsetvdr.SetPlotPaperUnits(plotset, PlotPaperUnit.Millimeters);
                            plotsetvdr.SetPlotCentered(plotset, true);
                            Point2d offset = new Point2d(0, 0);
                            #region "Set offset"
                            if (Plotterused == "New_Style_old_trafford")
                            {
                                if (Papersize == "A0")
                                {
                                    offset = new Point2d(10, 0.05);
                                }
                                else if (Papersize == "A1")
                                {
                                    offset = new Point2d(10, 0.04);
                                }
                                else //papersize is A2
                                {
                                    offset = new Point2d(10, 0.57);
                                }
                            }
                            else if (Plotterused == "New_Style_Goodison_PS")
                            {
                                if (Papersize == "A0")
                                {
                                    offset = new Point2d(10, -0.47);
                                }
                                else if (Papersize == "A1")
                                {
                                    offset = new Point2d(10, 0.05);
                                }
                                else //papersize is A2
                                {
                                    offset = new Point2d(10, 0.73);
                                }
                            }
                            else //plotter is Anfield
                            {
                                //this offset is reversed due to the printer
                                //offset = new Point2d(10, 0.58);
                                offset = new Point2d(0.58, 10);
                            }
                            #endregion
                            plotsetvdr.SetPlotOrigin(plotset, offset);
                            //override the current settings?
                            plotinf.OverrideSettings = plotset;
                            //validate the new settings
                            plotvdr.Validate(plotinf);
                            //sets the layout plot settings
                            lt.CopyFrom(plotset);
                            //Open the BlockTableRecord for this Layout.
                            BlockTableRecord lbtr = (BlockTableRecord)tr.GetObject(lt.BlockTableRecordId, OpenMode.ForWrite, false);
                            BlockReference blkref = new BlockReference(Point3d.Origin, block);
                            lbtr.AppendEntity(blkref);
                            tr.AddNewlyCreatedDBObject(blkref, true);
                            LayoutblockId = blkref.ObjectId;
                            // add attribute references to the newly created block reference.
                            BlockTableRecord bd = (BlockTableRecord)tr.GetObject(block, OpenMode.ForWrite);
                            foreach (ObjectId attid in bd)
                            {
                                Entity ent = (Entity)tr.GetObject(attid, OpenMode.ForRead);
                                if (ent is AttributeDefinition)
                                {
                                    AttributeDefinition ad = (AttributeDefinition)ent;
                                    AttributeReference ar = new AttributeReference();
                                    ar.SetAttributeFromBlock(ad, blkref.BlockTransform);
                                    blkref.AttributeCollection.AppendAttribute(ar);
                                    tr.AddNewlyCreatedDBObject(ar, true);
                                    trm.QueueForGraphicsFlush();
                                }
                            }
                            lt.TabOrder = 1;
                            lt.DowngradeOpen();
                            ed.Regen();
                        }
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show("Error Message is: " + ex.Message + "\n\r" + "The StackTrace is:" + ex.StackTrace);
                    }
                    tr.Commit();
                    tr.Dispose();
                    DeleteAllLayouts(true);
                    //once the blocks are inserted we need to setup their attributes using
                    // the insertblock routine we used earlier for amending the viewport blocks.
                    InsertBlockJig blkclass = new InsertBlockJig();
                    blkclass.InsertBlock(block, 0, LayoutblockId, true, jobdetail, vpno.ToString());
                    //add our new viewports to the new layout.
                    AddNewVports(layoutname, BlockId, newvport);
                }
            }
        }

        private static string GetPaper(string Visprop)
        {
            if (Visprop == "A0")
            {
                Papersize = "ISO A0 - 841 x 1189 mm.";
            }
            else if (Visprop == "A1")
            {
                Papersize = "ISO A1 - 594 x 841 mm. (landscape)";
            }
            else if (Visprop == "A2")
            {
                Papersize = "ISO A2 - 420 x 594 mm. (landscape)";
            }
            else // Visprop must be A3
            {
                Papersize = "A3";
            }
            return Papersize;
        }

        private static string GetPlotter(string Visprop)
        {
            //this is so we alternate plotting of each large scale sheet!
            if (Visprop != "A3")
            {
                if (Plotterused == "New_Style_old_trafford.pc3")
                {
                    Plotterused = "New_Style_Goodison_PS.pc3";
                }
                else if (Plotterused == "New_Style_Goodison_PS.pc3")
                {
                    Plotterused = "New_Style_old_trafford.pc3";
                }
            }
            else
            {
                Plotterused = "New_Style_Anfield.pc3";
            }
            return Plotterused;
        }
        /// <summary>
        /// Utilises an imported ObjectARX Function to set a viewport current.
        /// </summary>
        /// <param name="AcDbVport">The Viewport we wish to set active.</param>
        /// <returns>Returns a boolean value based on the outcome.</returns>
        [DllImport("acad.exe", CallingConvention = CallingConvention.Cdecl,
            EntryPoint = "?acedSetCurrentVPort@@YA?AW4ErrorStatus@Acad@@PBVAcDbViewport@@@Z")]
        extern static private int acedSetCurrentVPort(IntPtr AcDbVport);
        [DllImport("acad.exe", CallingConvention = CallingConvention.Cdecl,
            EntryPoint = "?acedVportTableRecords2Vports@@YA?AW4ErrorStatus@Acad@@XZ")]
        internal static extern ErrorStatus acedVportTableRecords2Vports();
        /// <summary>
        /// 
        /// </summary>
        /// <param name="db"></param>
        /// <returns></returns>
        public static void SortLayouts(Database db)
        {
            {
                SortedDictionary<int, string> layAndTab = new SortedDictionary<int, string>();
                string[] ldata = null;
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    DBDictionary layDict = (DBDictionary)db.LayoutDictionaryId.GetObject(OpenMode.ForRead);
                    foreach (DBDictionaryEntry entry in layDict)
                    {
                        Layout lay = (Layout)entry.Value.GetObject(OpenMode.ForRead);
                        layAndTab.Add(lay.TabOrder, lay.LayoutName);
                    }
                    ldata = (from n in layAndTab
                             where (n.Value != "Model")
                             select n.Value).ToArray();
                    Array.Sort(ldata);
                    trans.Commit();
                }
                // then make sure our new layouts match this list.
                for (int j = 0; j < ldata.Count(); j++)
                {
                    LayoutManager lman = LayoutManager.Current;
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        DBDictionary layDict = (DBDictionary)db.LayoutDictionaryId.GetObject(OpenMode.ForWrite);
                        foreach (DBDictionaryEntry entry in layDict)
                        {
                            Layout lay = (Layout)entry.Value.GetObject(OpenMode.ForWrite);
                            if (lay.LayoutName != "Model")
                            {
                                if (lay.LayoutName == ldata[j])
                                {
                                    if (lay.TabOrder != j + 1)
                                    {
                                        lay.TabOrder = j + 1;
                                    }
                                    break;
                                }
                            }
                        }
                        //commit the changes otherwise nothing happens!
                        tr.Commit();
                    }
                }
            }
        }
        /// <summary>
        /// Removes the default viewport created when you make a new layout.
        /// </summary>
        public static void RemoveExistingViewport(ObjectId LayoutId)
        {
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                Layout layout = tr.GetObject(LayoutId, OpenMode.ForRead) as Layout;
                Viewport vp = tr.GetObject(layout.GetViewports()[1], OpenMode.ForWrite) as Viewport;
                vp.Erase();
                tr.Commit();
            }

            // Deletes any custom UCSs still in the drawing.
        }
    }
}