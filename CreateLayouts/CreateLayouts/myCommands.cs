// (C) Copyright 2021 by Alex Fielder 
//
using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Windows.Forms;
using Autodesk.AutoCAD.Colors;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.PlottingServices;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(CreateLayouts.MyCommands))]

namespace CreateLayouts
{
    // This class is instantiated by AutoCAD for each document when
    // a command is called by the user the first time in the context
    // of a given document. In other words, non static data in this class
    // is implicitly per-document!
    public class MyCommands
    {
        [DllImport("acad.exe", CallingConvention = CallingConvention.Cdecl,
 EntryPoint = "?acedSetCurrentVPort@@YA?AW4ErrorStatus@Acad@@PBVAcDbViewport@@@Z")]
        extern static private int acedSetCurrentVPort(IntPtr AcDbVport);

        /// <summary>
        /// Boolean that states whether we inserted new blocks or not.
        /// </summary>
        public bool Blockswereinserted { get; set; }
        /// <summary>
        /// A string denoting Visibility of something which I forget.
        /// </summary>
        public string VisProp;
        //if we fail to new our dictionaries or lists we'll get a fatal error!
        /// <summary>
        /// a list of Jobdetails
        /// </summary>
        public static List<Jobdetails> detailslist = new List<Jobdetails>();
        /// <summary>
        /// a list of Layouts
        /// </summary>
        public static List<Layouts> LayoutList = new List<Layouts>();
        /// <summary>
        /// a list of Vports
        /// </summary>
        public static List<Vports> vport = new List<Vports>();
        /// <summary>
        /// A list for holding existing viewport details.
        /// </summary>
        public static List<Vports> ExistingvportList = new List<Vports>();
        /// <summary>
        /// A Dictionary that holds details of the new blocks we've inserted.
        /// </summary>
        public static Dictionary<int, ObjectId> blockdict = new Dictionary<int, ObjectId>();
        /// <summary>
        /// Contains an Array of available Floor values.
        /// </summary>
        public static ArrayList Floors = new ArrayList();
        /// <summary>
        /// Contains a Dictionary of Strings, ObjectIds relating to the keyplans required for our Tool.
        /// </summary>
        public static Dictionary<string, ObjectId> kpdict = new Dictionary<string, ObjectId>();
        Database db = HostApplicationServices.WorkingDatabase;
        Document dwg = AcadApp.DocumentManager.MdiActiveDocument;
        Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
        int ISPOSP = 0; // default to zero = ISP.
        String BlockName = "";

        [CommandMethod("CreateLayouts")]
        public void CreateLayouts()
        {
            // Before we change anything, we should save some details about the drawing - don't want to break anything!
            //need to check whether we've already run the tool otherwise we'll get an error when we try to save the layer state.
            int i = 0;
            int floorcount = 0;

            switch (MessageBox.Show("About to delete all layouts, (and Custom UCS entries) are you sure you wish to do this?", "Delete Layouts?", MessageBoxButtons.YesNo))
            {
                case DialogResult.Yes:
                    LayoutCommands.DeleteAllLayouts(false);
                    EverythingIsNew();
                    Document tmpdoc = AcadApp.DocumentManager.MdiActiveDocument;
                    LayerStateManager lyrStMan = tmpdoc.Database.LayerStateManager;
                    Blockswereinserted = false;
                    string tmplyrstName = "test";
                    if (lyrStMan.HasLayerState(tmplyrstName))
                    {
                        lyrStMan.DeleteLayerState(tmplyrstName);
                    }
                    // save the current layer states before we change/add anything.
                    LayerCommands layerclass = new LayerCommands();
                    layerclass.SaveLayerstate();
                    /*instead of deleting our custom ucs entries, I think I need to instead compare their details to those of the blocks 
                     that may or may not have already been inserted.*/
                    //UCSTools.RemoveCustomUCS();
                    break;
                case DialogResult.No:
                    goto theend;
                case DialogResult.Cancel:
                    goto theend;
            }
            SelectJobType:
            switch (MessageBox.Show("Is this an ISP Job?", "ISP?", MessageBoxButtons.YesNo))
            {
                case DialogResult.Yes:
                    ISPOSP = 0;
                    BlockName = "ISPVIEWPORT";
                    break;
                case DialogResult.No:
                    switch (MessageBox.Show("Is this an OSP Job?", "OSP?", MessageBoxButtons.YesNo))
                    {
                        case DialogResult.Yes:
                            ISPOSP = 1;
                            BlockName = "OSPVIEWPORT";
                            break;
                        case DialogResult.No:
                            //Application.Exit;
                            break;
                    }
                    break;
            }
            if (BlockName == "")
            {
                goto SelectJobType;
            }
            Boolean InsertVPBlock = true;
            Boolean FirstTime = true;
            //this will count the number of blocks already in the drawing prior to inserting more.
            Dictionary<int, ObjectId> origblockdict = GetListofInsertedBlocks(BlockName, true);
            do
            {
                switch (MessageBox.Show("Would you like to insert a viewport block?", "Insert block", MessageBoxButtons.YesNo))
                {
                    case DialogResult.Yes:
                        InsertBlockJig BlkClass = new InsertBlockJig();
                        BlkClass.RotationAngleUse = false;
                        /*need to add this here in order to account for already inserted blocks! 
                         but need we also should only need to do this once as after that blockdict
                         will have a complete list of available blocks.*/
                        if (FirstTime)
                        {
                            blockdict = GetListofInsertedBlocks(BlockName, false);
                            FirstTime = false;
                        }
                        if (blockdict.Count > 0)
                        {
                            if (blockdict.Count == 1)
                            {
                                i = 0;
                            }
                            else
                            {
                                i = blockdict.Count;
                            }
                            //zoom to the last block we inserted.
                            ZoomCommands.ZoomToEntity(blockdict[i]);
                            blockdict.Add(blockdict.Count + 1, BlkClass.StartInsert(BlockName, ISPOSP));
                        }
                        else
                        {
                            blockdict.Add(i, BlkClass.StartInsert(BlockName, ISPOSP));
                        }
                        if (blockdict.Count > origblockdict.Count)
                        {
                            Blockswereinserted = true;
                        }
                        else
                        {
                            Blockswereinserted = false;
                        }
                        BlkClass = null;
                        db.UpdateExt(true);
                        //zoom extents
                        ZoomCommands.ZoomWindow(ed, db.Extmin, db.Extmax);
                        i++;
                        break;
                    case DialogResult.No:
                        InsertVPBlock = false;
                        if (BlockName == "")
                        { goto SelectJobType; }
                        else
                        {
                            blockdict = GetListofInsertedBlocks(BlockName, false);
                        }
                        if (blockdict.Count > origblockdict.Count)
                        {
                            Blockswereinserted = true;
                        }
                        else
                        {
                            Blockswereinserted = false;
                        }
                        //ExistingvportList = RetrieveExistingViewportDetails(blockdict);
                        break;
                }
            } while (InsertVPBlock != false);
            switch (MessageBox.Show("Should we continue to the layout creation?", "Create New Layouts?", MessageBoxButtons.YesNo))
            {
                case DialogResult.Yes:
                    //db.UpdateExt(true);
                    ZoomCommands.ZoomWindow(ed, db.Extmin, db.Extmax);
                    /*as we have by this point collected the ObjectIds of the blocks we inserted it should be easy to sort them into groups
                     and assign a keyplan block accordingly. */
                    KeyplanCommands kpclass = new KeyplanCommands();
                    //if we don't new this here, we get multiple keyplans we don't need?!
                    kpdict = new Dictionary<string, ObjectId>();
                    List<Keyplans> keyplans = new List<Keyplans>();
                    foreach (string floor in Floors)
                    {
                        foreach (KeyValuePair<int, ObjectId> kp in blockdict)
                        {
                            Vports keyplan = (from Vports vp in ExistingvportList
                                              where vp.blockid == kp.Value
                                              select vp).First();
                            if (keyplan.FloorName == floor)
                            {
                                Keyplans newkp = new Keyplans();
                                newkp.groupint = floorcount;
                                newkp.oid = keyplan.blockid;
                                keyplans.Add(newkp);
                            }
                        }
                        floorcount++;
                    }

                    kpdict = kpclass.DrawKeyplans(keyplans, floorcount);
                    //kpdict = kpclass.DrawKeyplans();
                    if (blockdict.Count == 0)
                    {
                        MessageBox.Show("Cannot continue, there aren't any blocks inserted!");
                        break;
                    }
                    SetupNewLayoutDetails();
                    BeginCreateLayouts(kpdict);
                    LayoutCommands.DeleteAllLayouts(true);
                    //Sort the layouts
                    LayoutCommands.SortLayouts(db);
                    break;
                case DialogResult.No:
                    break;
            }
            EverythingIsNull();
            db.Dispose();
            ed.Regen();
            theend:
            return;
        }

        private void EverythingIsNew()
        {
            throw new NotImplementedException();
        }

        private void EverythingIsNull()
        {
            throw new NotImplementedException();
        }

        private void BeginCreateLayouts(Dictionary<string, ObjectId> kpdict)
        {
            throw new NotImplementedException();
        }

        private void SetupNewLayoutDetails()
        {
            throw new NotImplementedException();
        }

        private Dictionary<int, ObjectId> GetListofInsertedBlocks(string blockName, bool v)
        {
            throw new NotImplementedException();
        }

        //// The CommandMethod attribute can be applied to any public  member 
        //// function of any public class.
        //// The function should take no arguments and return nothing.
        //// If the method is an intance member then the enclosing class is 
        //// intantiated for each document. If the member is a static member then
        //// the enclosing class is NOT intantiated.
        ////
        //// NOTE: CommandMethod has overloads where you can provide helpid and
        //// context menu.

        //// Modal Command with localized name
        ////[CommandMethod("MyGroup", "MyCommand", "MyCommandLocal", CommandFlags.Modal)]
        //[CommandMethod("MyCommand")]
        //public void MyCommand() // This method can have any name
        //{
        //    // Put your command code here
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Editor ed;
        //    if (doc != null)
        //    {
        //        ed = doc.Editor;
        //        ed.WriteMessage("Hello, this is your first command.");

        //    }
        //}

        //// Modal Command with pickfirst selection
        //[CommandMethod("MyGroup", "MyPickFirst", "MyPickFirstLocal", CommandFlags.Modal | CommandFlags.UsePickSet)]
        //public void MyPickFirst() // This method can have any name
        //{
        //    PromptSelectionResult result = Application.DocumentManager.MdiActiveDocument.Editor.GetSelection();
        //    if (result.Status == PromptStatus.OK)
        //    {
        //        // There are selected entities
        //        // Put your command using pickfirst set code here
        //    }
        //    else
        //    {
        //        // There are no selected entities
        //        // Put your command code here
        //    }
        //}

        //// Application Session Command with localized name
        //[CommandMethod("MyGroup", "MySessionCmd", "MySessionCmdLocal", CommandFlags.Modal | CommandFlags.Session)]
        //public void MySessionCmd() // This method can have any name
        //{
        //    // Put your command code here
        //}

        //// LispFunction is similar to CommandMethod but it creates a lisp 
        //// callable function. Many return types are supported not just string
        //// or integer.
        //[LispFunction("MyLispFunction", "MyLispFunctionLocal")]
        //public int MyLispFunction(ResultBuffer args) // This method can have any name
        //{
        //    // Put your command code here

        //    // Return a value to the AutoCAD Lisp Interpreter
        //    return 1;
        //}

    }

}
