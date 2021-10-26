// (C) Copyright 2021 by Alex Fielder 
//
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace CreateLayouts
{
    internal class InsertBlockJig
    {
        /// <summary>
        /// The Masterscale of our blocks.
        /// </summary>
        public static double masterscale;
        /// <summary>
        /// 
        /// </summary>
        public static string masterscalestr;
        /// <summary>
        /// 
        /// </summary>
        public static Point3d BlockPos = new Point3d(0, 0, 0);
        /// <summary>
        /// 
        /// </summary>
        public static Scale3d BlockScale = new Scale3d(1);
        /// <summary>
        /// 
        /// </summary>
        public static Double BlockAngle = 0;
        /// <summary>
        /// 
        /// </summary>
        public static Point3d OldBlockPos = new Point3d(0, 0, 0);
        /// <summary>
        /// 
        /// </summary>
        public static int LastView = 1;
        /// <summary>
        /// 
        /// </summary>
        public List<Vports> vports = new List<Vports>();
        /// <summary>
        /// 
        /// </summary>
        public class InsertJig : EntityJig
        {
            /// <summary>
            /// 
            /// </summary>
            public Point3d position;
            /// <summary>
            /// 
            /// </summary>
            public double blockRotation;
            /// <summary>
            /// 
            /// </summary>
            public double m_initialscale = AssignScaleValue();
            /// <summary>
            /// 
            /// </summary>
            public double m_resultantscale;
            /// <summary>
            /// 
            /// </summary>
            public Point3d m_scalepoint;
            /// <summary>
            /// 
            /// </summary>
            public Matrix3d m_ucs;
            /// <summary>
            /// 
            /// </summary>
            private int vJigPromptCounter;
            private bool vRotationAngleUse = false;
            private double vRotationAngleValue = 0.0;
            //private double vScaleValue = AssignScaleValue();
            private double vScaleValue = 1.0;
            /// <summary>
            /// 
            /// </summary>
            /// <returns></returns>
            public static double AssignScaleValue()
            {
                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                PromptKeywordOptions keywordOptions = new PromptKeywordOptions("\nSelect a scale for the new block:");
                keywordOptions.Keywords.Add("1:20");
                keywordOptions.Keywords.Add("1:50");
                keywordOptions.Keywords.Add("1:100");
                keywordOptions.Keywords.Add("1:200");
                keywordOptions.Keywords.Add("Custom");
                keywordOptions.Keywords.Default = "1:50";
                PromptResult getWhichEntityResult = ed.GetKeywords(keywordOptions);
                if (getWhichEntityResult.StringResult == "Custom")
                    getWhichEntityResult = ed.GetString("\nEnter a Custom/Non-Standard Scale:");
                masterscalestr = getWhichEntityResult.StringResult;
                masterscale = ConvertScaleStrToDouble(getWhichEntityResult.StringResult);
                return masterscale;
            }
            /// <summary>
            /// 
            /// </summary>
            /// <param name="BlockId"></param>
            /// <param name="Position"></param>
            /// <param name="Normal"></param>
            public InsertJig(ObjectId BlockId, Point3d Position, Vector3d Normal)
                : base(new BlockReference(Position, BlockId))
            {
                BlockReference.Normal = Normal;
                position = Position;
            }
            /// <summary>
            /// 
            /// </summary>
            /// <param name="prompts"></param>
            /// <returns></returns>
            protected override SamplerStatus Sampler(JigPrompts prompts)
            {
                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                if (vJigPromptCounter == 0)
                {
                    //vScaleValue = AssignScaleValue();
                    vScaleValue = m_initialscale; // the default scale is set to 1:50!
                    this.BlockReference.ScaleFactors = new Scale3d(vScaleValue, vScaleValue, vScaleValue);
                    // And eval rotation here...
                    if (vRotationAngleUse) this.BlockReference.Rotation = vRotationAngleValue;
                    // Maybe still check the counter no (1) if the rotation was  set...

                    JigPromptPointOptions jigOpts = new JigPromptPointOptions();
                    jigOpts.DefaultValue = new Point3d(0, 0, 0);
                    jigOpts.UserInputControls = UserInputControls.Accept3dCoordinates | UserInputControls.NullResponseAccepted;
                    jigOpts.Message = "\nInsertion point: ";
                    PromptPointResult res = prompts.AcquirePoint(jigOpts);

                    //Point3d BlkRefPoint = BlockReference.Position;
                    Point3d curPoint = res.Value;
                    if (position.DistanceTo(curPoint) > 1.0e-4)  // 1.0e-4 make it 10 units (depending on drawing...)
                        // if (this.BlockReference.Position != curPoint)  // 1.0e-4 make it 10 units (depending on drawing...)
                        position = curPoint;
                    else
                        return SamplerStatus.NoChange;

                    if (res.Status == PromptStatus.Cancel)
                        return SamplerStatus.Cancel;
                    else
                        return SamplerStatus.OK;
                }

                else
                {
                    if (vJigPromptCounter == 1)
                    {
                        // Activate Rotate Jig...Just check if there already was a rotation...??
                        // JigPromptAngleOptions
                        JigPromptAngleOptions jigOpts = new JigPromptAngleOptions();
                        jigOpts.BasePoint = this.BlockReference.Position;
                        jigOpts.DefaultValue = 0;
                        jigOpts.Message = "Get angle: ";
                        jigOpts.Cursor = CursorType.RubberBand;
                        jigOpts.UseBasePoint = true;
                        jigOpts.UserInputControls = UserInputControls.Accept3dCoordinates | UserInputControls.NullResponseAccepted; // Not Yet..
                        PromptDoubleResult res = prompts.AcquireAngle(jigOpts);

                        if (res.Status == PromptStatus.OK)
                        {
                            double curAngle = res.Value;
                            if (this.BlockReference.Rotation != curAngle)
                            {
                                blockRotation = curAngle;
                                return SamplerStatus.OK;
                            }
                            else
                                return SamplerStatus.NoChange;
                        }
                        else
                            if (res.Status == PromptStatus.None)
                        {
                            // .... ????
                            blockRotation = 0;
                            return SamplerStatus.OK;
                        }
                        else
                            return SamplerStatus.Cancel;
                    }
                    else
                    {
                        return SamplerStatus.NoChange;
                    }
                }

            }
            /// <summary>
            /// 
            /// </summary>
            /// <returns></returns>
            protected override bool Update()
            {
                try
                {
                    bool returnvalue = false;
                    switch (vJigPromptCounter)
                    {
                        case 0:
                            // Insert point Update...

                            if (this.BlockReference.Position.DistanceTo(position) > 1.0e-4)
                            {
                                this.BlockReference.Position = position;
                                returnvalue = true;
                            }
                            break;
                        case 1:
                            // Rotate Update...

                            if (this.BlockReference.Rotation != blockRotation)
                            {
                                this.BlockReference.Rotation = blockRotation;
                                returnvalue = true;
                            }
                            break;
                        case 2:
                            // Scale update...
                            Scale3d scaleresult = new Scale3d(m_resultantscale, m_resultantscale, m_resultantscale);
                            if (this.BlockReference.ScaleFactors != scaleresult)
                            {
                                Matrix3d trans = Matrix3d.Scaling(m_resultantscale - m_initialscale, m_scalepoint);
                                this.BlockReference.TransformBy(trans);
                                // the base scale becomes the previous resultantscale
                                // and the resultantscale is set to zero
                                m_initialscale = m_resultantscale;
                                m_resultantscale = 0.0;

                                returnvalue = true;
                            }
                            break;
                    }
                    return returnvalue;
                }
                catch (System.Exception)
                {
                    return false;
                }
            }
            /// <summary>
            /// 
            /// </summary>
            public bool RotationAngleUse
            {
                get { return vRotationAngleUse; }
                set { vRotationAngleUse = value; }
            }
            /// <summary>
            /// 
            /// </summary>
            public double RotationAngleValue
            {
                get { return vRotationAngleValue; }
                set { vRotationAngleValue = value; }
            }
            /// <summary>
            /// 
            /// </summary>
            public double ScaleValue
            {
                get { return vScaleValue; }
                set { if (value != 0) vScaleValue = value; }
            }
            /// <summary>
            /// 
            /// </summary>
            public int JigPromptCounter
            {
                get { return vJigPromptCounter; }
                set { vJigPromptCounter = value; }
            }
            /// <summary>
            /// 
            /// </summary>
            /// <returns></returns>
            public bool UpdateBlockRef()
            {
                try
                {
                    bool returnvalue = false;
                    switch (vJigPromptCounter)
                    {
                        case 0:
                            if (this.BlockReference.Position != position)
                            {
                                this.BlockReference.Position = position;
                                returnvalue = true;
                            }
                            break;
                        case 1:
                            if (this.BlockReference.Rotation != blockRotation)
                            {
                                this.BlockReference.Rotation = blockRotation;
                                returnvalue = true;
                            }
                            break;

                    }
                    return returnvalue;
                }
                catch (System.Exception)
                {
                    return false;
                }

            }
            /// <summary>
            /// 
            /// </summary>
            public BlockReference BlockReference
            {
                get { return base.Entity as BlockReference; }
            }


        }

        private double vRotationVal = 0.0;
        private bool vRotationUse = true;
        private double vScaleVal = 1.0;

        /// <summary>
        /// 
        /// </summary>
        public bool RotationAngleUse
        {
            get { return vRotationUse; }
            set { vRotationUse = value; }
        }
        /// <summary>
        /// 
        /// </summary>
        public double RotationAngleValue
        {
            get { return vRotationVal; }
            set { vRotationVal = value; }
        }
        /// <summary>
        /// 
        /// </summary>
        public double ScaleValue
        {
            get { return vScaleVal; }
            set
            {
                if (value != 0) vScaleVal = value;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="scalestr"></param>
        /// <returns></returns>
        public static double ConvertScaleStrToDouble(string scalestr)
        {
            int stringloc = scalestr.IndexOf(":");
            string tmpstr = scalestr.Substring(stringloc + 1);
            //string tmpstr = scalestr.TrimStart('1',':');
            double convertedstr = Convert.ToDouble(tmpstr);
            return convertedstr;
        }
        /// <summary>
        /// 
        /// </summary>
        public int lastview = 0;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Blockid"></param>
        /// <param name="ISPOSP"></param>
        /// <param name="JigBlockId"></param>
        /// <param name="isLayout"></param>
        /// <param name="jobdetail"></param>
        /// <param name="LayoutNo"></param>
        /// <returns></returns>
        public List<Vports> InsertBlock(ObjectId Blockid, int ISPOSP, ObjectId JigBlockId, bool isLayout, List<Jobdetails> jobdetail, string LayoutNo)
        {
            //this assumes that because the viewport block has been copied to the drawing, the other necessary blocks will have been too.
            Database db = HostApplicationServices.WorkingDatabase;
            Document dwg = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = dwg.Editor;
            Point3d BlockPos = new Point3d(0, 0, 0);
            Scale3d BlockScale = new Scale3d(1);
            Double BlockAngle = 0;
            PromptStringOptions prso = null;
            PromptKeywordOptions prko = null;
            PromptResult res = null;
            string update = null;
            Vector3d Xaxis = new Vector3d();
            Vector3d Yaxis = new Vector3d();
            // this next section adds the ISPviewport-frame block on top of the already-inserted ISPViewport block.
            if (ISPOSP == 0 & isLayout == false) // Viewport = ISP
            #region InsertISPViewportframe
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead, false) as BlockTable;
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject((from ObjectId btrid in bt
                                                                           where btrid == Blockid
                                                                           select btrid).First(), OpenMode.ForRead, false);
                    // normally we could filter out layouts, xrefs etc. but since we already know the objectid of the block 
                    // we're looking for we don't need to bother.
                    if (btr.IsLayout || btr.IsFromExternalReference || !btr.IsDynamicBlock || btr.IsFromOverlayReference || !btr.HasAttributeDefinitions)
                    {
                        return vports;
                    }
                    ObjectIdCollection blkrefIds = DBUtils.GetBlockReferenceIds(btr.ObjectId);
                    if (blkrefIds == null || blkrefIds.Count == 0)
                        return vports;
                    BlockReference blkref = (BlockReference)tr.GetObject((from ObjectId blkrefid in blkrefIds
                                                                          where blkrefid == JigBlockId
                                                                          select blkrefid).Single(), OpenMode.ForRead, false);
                    BlockPos = blkref.Position;
                    BlockScale = blkref.ScaleFactors;
                    BlockAngle = blkref.Rotation;
                    Vports vport = new Vports();
                    vport.blockid = blkref.ObjectId;
                    String ViewNo = null;
                    String Sheetsize = null;
                    //Dynamic Block commands:
                    DynamicBlockReferencePropertyCollection DynpropsCol = blkref.DynamicBlockReferencePropertyCollection;
                    foreach (DynamicBlockReferenceProperty DynpropsRef in DynpropsCol)
                    {
                        string DynPropsName = DynpropsRef.PropertyName.ToUpper();
                        switch (DynPropsName)
                        {
                            case "LOOKUP":
                                prko = new PromptKeywordOptions("\nWhat size sheet would you like to use?");
                                prko.Keywords.Add("A3");
                                prko.Keywords.Add("A2");
                                prko.Keywords.Add("A1");
                                prko.Keywords.Add("A0");
                                prko.Keywords.Default = "A3";
                                PromptResult dynpropres = ed.GetKeywords(prko);
                                if (dynpropres.Status == PromptStatus.OK)
                                {
                                    Sheetsize = dynpropres.StringResult;
                                    DynpropsRef.Value = dynpropres.StringResult;
                                    vport.PageSize = Sheetsize;
                                    vport.vpVisProp = Sheetsize;
                                }
                                break;
                        }
                    }
                    #region EditViewportBlockAtts
                    // fill out the attributes for this block.
                    //int i = 1;
                    AttributeCollection attrefids = blkref.AttributeCollection;

                    foreach (ObjectId attrefid in attrefids)
                    {
                        AttributeReference attref = tr.GetObject(attrefid, OpenMode.ForWrite, false) as AttributeReference;
                        //vport = new List<Vports>();
                        //vport.Capacity = i + 1;
                        switch (attref.Tag)
                        {
                            case "VIEWNO_NONVIS":
                                prso = new PromptStringOptions("\nWhat view number is this?");
                                prso.DefaultValue = Convert.ToString(MyCommands.blockdict.Count + 1);
                                //prso.DefaultValue = Convert.ToString(LastView);
                                LastView++;
                                // need to setup a count function to automatically add 1 to each viewno.
                                res = ed.GetString(prso);
                                if (res.Status == PromptStatus.OK)
                                {
                                    attref.UpgradeOpen();
                                    update = res.StringResult;
                                    ViewNo = update;
                                    attref.TextString = update;
                                    attref.DowngradeOpen();
                                    LastView = Convert.ToInt16(update) + 1;
                                    vport.vpName = ViewNo;
                                    vport.vpNumber = Convert.ToInt32(vport.vpName);
                                }
                                break;
                            case "FULLORPARTIAL":
                                prko = new PromptKeywordOptions("\nWhat kind of view does this view depict?");
                                prko.Keywords.Add("FULL");
                                prko.Keywords.Add("PARTIAL");
                                prko.Keywords.Add("DC");
                                prko.Keywords.Add("ELEVATION");
                                prko.Keywords.Default = "FULL";
                                res = ed.GetKeywords(prko);
                                if (res.Status == PromptStatus.OK)
                                {
                                    attref.UpgradeOpen();
                                    update = res.StringResult;
                                    attref.TextString = update;
                                    attref.DowngradeOpen();
                                    vport.FullorPartial = update;
                                    if (res.StringResult == "PARTIAL")
                                    {
                                        btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                                        BlockReference frameref = new BlockReference(BlockPos, bt["ISPVIEWPORT-FRAME"]);
                                        frameref.ScaleFactors = BlockScale;
                                        frameref.Rotation = BlockAngle;
                                        btr.AppendEntity(frameref);
                                        tr.AddNewlyCreatedDBObject(frameref, true);
                                        //Dynamic Block Settings:
                                        DynamicBlockReferencePropertyCollection frameDynpropsCol = frameref.DynamicBlockReferencePropertyCollection;
                                        foreach (DynamicBlockReferenceProperty frameDynpropsRef in frameDynpropsCol)
                                        {
                                            switch (frameDynpropsRef.PropertyName.ToUpper())
                                            {
                                                case "LOOKUP":
                                                    frameDynpropsRef.Value = Sheetsize;
                                                    break;
                                            }
                                        }
                                        // Add the attributes for this block!
                                        BlockTableRecord bd = (BlockTableRecord)tr.GetObject(bt["ISPVIEWPORT-FRAME"], OpenMode.ForRead);
                                        foreach (ObjectId attid in bd)
                                        {
                                            Entity ent = (Entity)tr.GetObject(attid, OpenMode.ForRead);
                                            if (ent is AttributeDefinition)
                                            {
                                                AttributeDefinition ad = (AttributeDefinition)ent;
                                                AttributeReference ar = new AttributeReference();
                                                ar.SetAttributeFromBlock(ad, frameref.BlockTransform);
                                                frameref.AttributeCollection.AppendAttribute(ar);
                                                tr.AddNewlyCreatedDBObject(ar, true);
                                            }
                                        }

                                        AttributeCollection frameattrefids = frameref.AttributeCollection;
                                        foreach (ObjectId frameattrefid in frameattrefids)
                                        {
                                            AttributeReference frameattref = tr.GetObject(frameattrefid, OpenMode.ForWrite, false) as AttributeReference;
                                            switch (frameattref.Tag)
                                            {
                                                case "VIEWNO_VIS":
                                                    frameattref.TextString = ViewNo;
                                                    break;
                                            }
                                        }
                                        // add new layers as required!
                                        Addlayers(blkref, frameref, Convert.ToInt32(ViewNo));
                                    }
                                }
                                break;
                            case "FLOOR":
                                prko = new PromptKeywordOptions("\nWhat floor does this view depict?");
                                prko.Keywords.Add("GROUND");
                                prko.Keywords.Add("FIRST");
                                prko.Keywords.Add("SECOND");
                                prko.Keywords.Add("THIRD");
                                prko.Keywords.Add("FOURTH");
                                prko.Keywords.Add("OTHER");
                                prko.Keywords.Default = "GROUND";
                                PromptResult getWhichFloorResult = ed.GetKeywords(prko);
                                if (getWhichFloorResult.StringResult == "OTHER")
                                {
                                    res = ed.GetString("\nEnter another floor that's not listed.");
                                }
                                else
                                {
                                    res = getWhichFloorResult;
                                }
                                if (res.Status == PromptStatus.OK)
                                {
                                    attref.UpgradeOpen();
                                    update = res.StringResult;
                                    attref.TextString = update;
                                    vport.FloorName = update;
                                    attref.DowngradeOpen();
                                    MyCommands.Floors.Add(update);
                                }
                                break;
                            case "XLOCATION":
                                Xaxis = BlockPos.GetVectorTo(attref.Position);
                                break;
                            case "YLOCATION":
                                Yaxis = BlockPos.GetVectorTo(attref.Position);
                                break;
                            case "VPSCALE":
                                attref.TextString = masterscalestr;
                                vport.vpScale = masterscalestr;
                                break;
                            case "VIEWCENTRE":
                                vport.vpMSCentrePoint = new Point2d(attref.Position.X, attref.Position.Y);
                                break;
                            case "":
                                break;
                        }
                    }
                    vports.Add(vport);
                    #endregion //EditViewportBlockAtts
                    #region NewUCS
                    ViewNo = "UCS_" + ViewNo;
                    UCSTools.GetorCreateUCS(ViewNo, BlockPos, Xaxis, Yaxis);
                    #endregion //NewUCS
                    tr.Commit();
                    /* can't return before committing the transaction otherwise 
                    the attribute changes will fail to
                    stick! */
                    return vports;
                }
                #endregion //InsertISPViewportframe
            }
            else if (ISPOSP == 1 & isLayout == false) // OSP Block required - less code needed here as we use less options.
            #region InsertOSPViewPortframe
            {

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead, false) as BlockTable;
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject((from ObjectId btrid in bt
                                                                           where btrid == Blockid
                                                                           select btrid).Single(), OpenMode.ForRead, false);
                    // normally we could filter out layouts, xrefs etc. but since we already know the objectid of the block 
                    // we're looking for we don't need to bother.
                    if (btr.IsLayout || btr.IsFromExternalReference || !btr.IsDynamicBlock || btr.IsFromOverlayReference || !btr.HasAttributeDefinitions)
                    {
                        return vports;
                    }
                    ObjectIdCollection blkrefIds = DBUtils.GetBlockReferenceIds(btr.ObjectId);
                    if (blkrefIds == null || blkrefIds.Count == 0)
                        return vports;
                    BlockReference blkref = (BlockReference)tr.GetObject((from ObjectId blkrefid in blkrefIds
                                                                          where blkrefid == JigBlockId
                                                                          select blkrefid).Single(), OpenMode.ForRead, false);
                    BlockPos = blkref.Position;
                    BlockScale = blkref.ScaleFactors;
                    BlockAngle = blkref.Rotation;
                    String ViewNo = null;
                    String Sheetsize = null;
                    //Dynamic Block commands:
                    DynamicBlockReferencePropertyCollection DynpropsCol = blkref.DynamicBlockReferencePropertyCollection;
                    foreach (DynamicBlockReferenceProperty DynpropsRef in DynpropsCol)
                    {
                        string DynPropsName = DynpropsRef.PropertyName.ToUpper();
                        switch (DynPropsName)
                        {
                            case "LOOKUP":
                                prko = new PromptKeywordOptions("\nWhat size sheet would you like to use?");
                                prko.Keywords.Add("A3");
                                prko.Keywords.Add("A2");
                                prko.Keywords.Add("A1");
                                prko.Keywords.Add("A0");
                                prko.Keywords.Default = "A3";
                                PromptResult dynpropres = ed.GetKeywords(prko);
                                if (dynpropres.Status == PromptStatus.OK)
                                {
                                    Sheetsize = dynpropres.StringResult;
                                    DynpropsRef.Value = dynpropres.StringResult;
                                }
                                break;
                        }
                    }
                    #region EditViewportBlockAtts
                    // fill out the attributes for this block.

                    AttributeCollection attrefids = blkref.AttributeCollection;
                    foreach (ObjectId attrefid in attrefids)
                    {
                        AttributeReference attref = tr.GetObject(attrefid, OpenMode.ForWrite, false) as AttributeReference;
                        switch (attref.Tag)
                        {
                            case "VIEWNO_NONVIS":
                                prso = new PromptStringOptions("\nWhat view number is this?");
                                res = ed.GetString(prso);
                                if (res.Status == PromptStatus.OK)
                                {
                                    attref.UpgradeOpen();
                                    update = res.StringResult;
                                    ViewNo = update;
                                    attref.TextString = update;
                                    attref.DowngradeOpen();
                                }
                                break;
                            case "FULLORPARTIAL":
                                prko = new PromptKeywordOptions("\nWhat kind of view will this view depict?");
                                prko.Keywords.Add("FULL");
                                prko.Keywords.Add("PARTIAL");
                                prko.Keywords.Default = "FULL";
                                res = ed.GetKeywords(prko);
                                if (res.Status == PromptStatus.OK)
                                {
                                    attref.UpgradeOpen();
                                    update = res.StringResult;
                                    attref.TextString = update;
                                    attref.DowngradeOpen();
                                    if (res.StringResult == "PARTIAL")
                                    {
                                        btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                                        BlockReference frameref = new BlockReference(BlockPos, bt["OSPVIEWPORT-FRAME"]);
                                        frameref.ScaleFactors = BlockScale;
                                        frameref.Rotation = BlockAngle;
                                        btr.AppendEntity(frameref);
                                        tr.AddNewlyCreatedDBObject(frameref, true);
                                        //Dynamic Block Settings:
                                        DynamicBlockReferencePropertyCollection frameDynpropsCol = frameref.DynamicBlockReferencePropertyCollection;
                                        foreach (DynamicBlockReferenceProperty frameDynpropsRef in frameDynpropsCol)
                                        {
                                            switch (frameDynpropsRef.PropertyName.ToUpper())
                                            {
                                                case "LOOKUP":
                                                    frameDynpropsRef.Value = Sheetsize;
                                                    break;
                                            }
                                        }
                                        // Add the attributes for this block!
                                        BlockTableRecord bd = (BlockTableRecord)tr.GetObject(bt["OSPVIEWPORT-FRAME"], OpenMode.ForRead);
                                        foreach (ObjectId attid in bd)
                                        {
                                            Entity ent = (Entity)tr.GetObject(attid, OpenMode.ForRead);
                                            if (ent is AttributeDefinition)
                                            {
                                                AttributeDefinition ad = (AttributeDefinition)ent;
                                                AttributeReference ar = new AttributeReference();
                                                ar.SetAttributeFromBlock(ad, frameref.BlockTransform);
                                                frameref.AttributeCollection.AppendAttribute(ar);
                                                tr.AddNewlyCreatedDBObject(ar, true);
                                            }
                                        }

                                        AttributeCollection frameattrefids = frameref.AttributeCollection;
                                        foreach (ObjectId frameattrefid in frameattrefids)
                                        {
                                            AttributeReference frameattref = tr.GetObject(frameattrefid, OpenMode.ForWrite, false) as AttributeReference;
                                            switch (frameattref.Tag)
                                            {
                                                case "VIEWNO_VIS":
                                                    frameattref.TextString = ViewNo;
                                                    break;
                                            }
                                        }
                                        // add new layers as required!
                                        Addlayers(blkref, frameref, Convert.ToInt32(ViewNo));
                                    }
                                }
                                break;
                            case "FLOOR":
                                attref.UpgradeOpen();
                                update = "GROUND";
                                attref.TextString = update;
                                attref.DowngradeOpen();
                                break;
                            case "XLOCATION":
                                Xaxis = BlockPos.GetVectorTo(attref.Position);
                                break;
                            case "YLOCATION":
                                Yaxis = BlockPos.GetVectorTo(attref.Position);
                                break;
                        }
                    }
                    #endregion //EditViewportBlockAtts
                    #region NewUCS
                    ViewNo = "UCS_" + ViewNo;
                    UCSTools.GetorCreateUCS(ViewNo, BlockPos, Xaxis, Yaxis);
                    #endregion //NewUCS
                    tr.Commit();
                    return vports;
                }
                #endregion //InsertOSPViewPortframe
            }
            else if (isLayout == true)
            #region "Layout Block"
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead, false) as BlockTable;
                    ObjectId btrid = (from ObjectId oid in bt
                                      where oid == Blockid
                                      select oid).First();
                    BlockTableRecord btr = tr.GetObject(btrid, OpenMode.ForWrite, false) as BlockTableRecord;
                    // normally we could filter out layouts, xrefs etc. but since we already know the objectid of the block 
                    // we're looking for we don't need to bother.
                    if (btr.IsLayout || btr.IsFromExternalReference || !btr.IsDynamicBlock || btr.IsFromOverlayReference || !btr.HasAttributeDefinitions)
                    {
                        return vports;
                    }
                    ObjectIdCollection blkrefIds = DBUtils.GetBlockReferenceIds(btr.ObjectId);
                    if (blkrefIds == null || blkrefIds.Count == 0)
                        return vports;
                    ObjectId borderrefid = (from ObjectId bid in blkrefIds
                                            where bid == JigBlockId
                                            select bid).First();
                    BlockReference borderref = tr.GetObject(borderrefid, OpenMode.ForRead, false) as BlockReference;
                    //get the AttributeCollection from the BlockReference
                    AttributeCollection attrefids = borderref.AttributeCollection;
                    int Sheetint = 0;
                    string sheetname = "";
                    string DrawnDate = "";
                    string Creator = "";
                    string SiteNo = "";
                    foreach (Jobdetails jdetails in jobdetail) // there should only be one entry?
                    {
                        sheetname = String.Concat(jdetails.SiteNo, "-", jdetails.BuildingNo, "-01");
                        DrawnDate = jdetails.Datestr;
                        Creator = jdetails.Initials;
                        SiteNo = jdetails.SiteNo;
                    }
                    foreach (ObjectId attrefid in attrefids)
                    {
                        AttributeReference attref = tr.GetObject(attrefid, OpenMode.ForWrite, false) as AttributeReference;
                        switch (attref.Tag)
                        {
                            case "DRAWING NUMBER":
                                attref.UpgradeOpen();
                                attref.TextString = sheetname;
                                attref.DowngradeOpen();
                                break;
                            case "~DATE":
                                attref.UpgradeOpen();
                                attref.TextString = DrawnDate;
                                attref.DowngradeOpen();
                                break;
                            case "REV":
                                attref.UpgradeOpen();
                                attref.TextString = "1";
                                attref.DowngradeOpen();
                                break;
                            case "SHEET":
                                foreach (Layouts layoutitem in MyCommands.LayoutList)
                                {
                                    //if (layoutitem.LayoutId == Blockid)
                                    //{
                                    Sheetint = layoutitem.LayoutNo;
                                    attref.UpgradeOpen();
                                    attref.TextString = Convert.ToString(Sheetint);
                                    attref.DowngradeOpen();
                                    //}
                                }
                                break;
                            case "~CREATOR":
                                attref.UpgradeOpen();
                                attref.TextString = Creator;
                                attref.DowngradeOpen();
                                break;
                            case "SCALE":
                                foreach (Vports vportitem in MyCommands.vport)
                                {
                                    if (Convert.ToUInt32(vportitem.vpName) == Sheetint)
                                    {
                                        string ScaleVal = "1:" + Convert.ToString(vportitem.vpScale);
                                        attref.UpgradeOpen();
                                        attref.TextString = ScaleVal;
                                        attref.DowngradeOpen();
                                    }
                                }

                                break;
                            case "WORK STREAM":
                                attref.UpgradeOpen();
                                attref.TextString = "COMMS AND SOMETHING";
                                attref.DowngradeOpen();
                                break;
                            case "~APPROVER":
                                attref.UpgradeOpen();
                                attref.TextString = "K. HETHERINGTON";
                                attref.DowngradeOpen();
                                break;
                        }
                    }
                    DynamicBlockReferencePropertyCollection DynpropsCol = borderref.DynamicBlockReferencePropertyCollection;
                    foreach (DynamicBlockReferenceProperty DynpropsRef in DynpropsCol)
                    {
                        string DynPropsName = DynpropsRef.PropertyName.ToUpper();
                        switch (DynPropsName)
                        {
                            case "VISIBILITY":
                                Vports vportitem = (from Vports v in MyCommands.ExistingvportList
                                                    where v.vpName == LayoutNo
                                                    select v).Single();
                                DynpropsRef.Value = vportitem.PageSize;
                                break;
                        }
                    }
                    //Without commiting these changes, to the db, nothing will stick!
                    tr.Commit();
                }
            }
            #endregion
            return vports;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="VPortBlkref"></param>
        /// <param name="VportFrameBlkref"></param>
        /// <param name="Numlayers"></param>
        public void Addlayers(BlockReference VPortBlkref, BlockReference VportFrameBlkref, int Numlayers)
        {
            Document Doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = Doc.Database;
            Editor ed = Doc.Editor;
            string LayerName = null;
            // define the layername we should first go looking for.
            for (int i = 1; i <= Numlayers; i++)
            {
                if (i == Numlayers)
                {
                    if (i < 10)
                    {
                        LayerName = "-02 KEYPLAN 00" + i;
                    }
                    else if ((i > 9) | (i < 100))
                    {
                        LayerName = "-02 KEYPLAN 0" + i;
                    }
                    else
                    {
                        LayerName = "-02 KEYPLAN " + i;
                    }
                }
            }
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                try
                {
                    if (lt.Has(LayerName))
                    {
                        // normally we'd delete the layer here, but if it already exists we should place the new block reference on it.
                        VPortBlkref.Layer = LayerName;
                        VportFrameBlkref.Layer = LayerName;
                    }
                    else
                    {
                        LayerTableRecord newltr = new LayerTableRecord();
                        newltr.Name = LayerName;
                        newltr.Color = Color.FromRgb(255, 0, 0);
                        newltr.LinetypeObjectId = GetLineTypeID(db, "Continuous");
                        lt.UpgradeOpen();
                        lt.Add(newltr);
                        //lt.DowngradeOpen();
                        tr.AddNewlyCreatedDBObject(newltr, true);
                        tr.Commit();
                        //change the layer of the newly inserted blockref.
                        VPortBlkref.Layer = LayerName;
                        VportFrameBlkref.Layer = LayerName;
                    }
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    tr.Abort();
                    System.Windows.Forms.MessageBox.Show(ex.ToString());
                }
                finally
                {
                    tr.Dispose();
                }
            }
        }

        private ObjectId GetLineTypeID(Database db, string LtypeName)
        {
            ObjectId id = ObjectId.Null;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LinetypeTable ltt = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;
                try
                {
                    LinetypeTableRecord lttr = (LinetypeTableRecord)tr.GetObject(ltt[LtypeName], OpenMode.ForRead);
                    if (ltt.Has(LtypeName))
                    {
                        id = lttr.ObjectId;
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("\nLinetype not found!");
                    }
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    tr.Abort();
                    System.Windows.Forms.MessageBox.Show(ex.ToString());
                }
                finally
                {
                    tr.Dispose();
                }
            }
            return id;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="JigBlockName"></param>
        /// <param name="ISPOSP"></param>
        /// <returns></returns>
        public ObjectId StartInsert(String JigBlockName, int ISPOSP)
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = HostApplicationServices.WorkingDatabase; // Application.DocumentManager.MdiActiveDocument.Database;
            ObjectId blockid = ObjectId.Null;
            try
            {
                Vector3d Normal = db.Ucsxdir.CrossProduct(db.Ucsydir);
                PromptResult res; // = ed.GetString(opts);
                if (JigBlockName != string.Empty)
                {
                    ObjectId block = BlockHelperClass.GetBlockId(db, JigBlockName);
                    block = BlockHelperClass.GetBlockId(db, JigBlockName);
                    if (block.IsNull)
                    {
                        ed.WriteMessage("\nBlock {0} not found.", JigBlockName);
                        return ObjectId.Null;
                    }
                    InsertJig jig = new InsertJig(block, Point3d.Origin, Normal.GetNormal());

                    //blockRotation  
                    if (vRotationUse == true)
                    {
                        jig.RotationAngleUse = true;
                        jig.RotationAngleValue = vRotationVal;
                    }
                    else
                    {
                        jig.RotationAngleUse = false;
                    }

                    jig.ScaleValue = vScaleVal;
                    jig.JigPromptCounter = 0; // Start with insertpoint

                    using (DocumentLock docLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
                    {
                        res = ed.Drag(jig); // First Jig option..[Insertpoint]

                        // Should it accept a rightclick as the 0-point of the drawing?
                        // or does that cancel the operation...?
                        if (res.Status == PromptStatus.OK)
                        {
                            jig.UpdateBlockRef(); // Check if it returns True!!!
                            // Here also the Rotation-Jig should Start...
                            // if Rotation was fixed this should have been set at first...
                            // Should Accept rightclick as 0-rotation
                            if (vRotationUse == true)
                            {
                                // Ready...end the Jig
                            }
                            else
                            {
                                // Start a new Jig for the rotation..
                                jig.JigPromptCounter = 1;
                                res = ed.Drag(jig);
                                if (res.Status == PromptStatus.OK)
                                {
                                    jig.UpdateBlockRef();
                                    //Set a bite to check if it should continue...(esc => no block placed!!)
                                }
                            }

                            // if jig.RotationAngleUse = false then do it...
                            // jig.RotationAngleValue = 1; // Or skip this altogether if it's a given rotation...
                            // No checking to be done by Jig(ger)..??

                            Autodesk.AutoCAD.DatabaseServices.TransactionManager tm = db.TransactionManager;


                            if (res.Status == PromptStatus.OK) // Second check needed if there was a rotation Jig..
                            {
                                using (Transaction tr = tm.StartTransaction())
                                {

                                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite, false);
                                    btr.AppendEntity(jig.BlockReference);
                                    tr.AddNewlyCreatedDBObject(jig.BlockReference, true);

                                    // Call a function to make the graphics display
                                    // (otherwise it will only do so when we Commit)
                                    // trying to add the attributes here!
                                    BlockTableRecord bd = (BlockTableRecord)tr.GetObject(block, OpenMode.ForWrite);
                                    foreach (ObjectId attid in bd)
                                    {
                                        Entity ent = (Entity)tr.GetObject(attid, OpenMode.ForRead);
                                        if (ent is AttributeDefinition)
                                        {
                                            AttributeDefinition ad = (AttributeDefinition)ent;
                                            AttributeReference ar = new AttributeReference();
                                            ar.SetAttributeFromBlock(ad, jig.BlockReference.BlockTransform);
                                            jig.BlockReference.AttributeCollection.AppendAttribute(ar);
                                            tr.AddNewlyCreatedDBObject(ar, true);
                                        }
                                    }
                                    //end attribute update?
                                    tm.QueueForGraphicsFlush();
                                    tr.Commit();
                                }
                            }
                        }
                    }
                    // Insert the viewportframe block (if ISP) here and update the attributes in the Viewport block.
                    blockid = jig.BlockReference.ObjectId;
                    List<Jobdetails> fakelist = null;
                    string blah = "";
                    MyCommands.vport = InsertBlock(block, ISPOSP, blockid, false, fakelist, blah);
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("\nThe Error was: " + ex.Message);
                return ObjectId.Null;
                //return false;
            }
            db.Dispose();
            return blockid;
        }
    }
}