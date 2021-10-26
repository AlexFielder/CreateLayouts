// (C) Copyright 2021 by Alex Fielder 
//
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace CreateLayouts
{
    public class KeyplanCommands
    {
        /// <summary>
        /// Creates an automated set of keyplans (and stores their ObjectIds) based on the input list of required keyplans
        /// generated according to what floor and VisProp each block has.
        /// </summary>
        /// <param name="keyplans">A List of integers and ObjectIds grouped according to what floor and VisProp each block has</param>
        /// <param name="floorcount">A count of the different floors used in this drawing - Helps us to group things together.</param>
        /// <returns>Returns a list of ObjectIds that can be zoomed to in each new Layout.</returns>
        public Dictionary<string, ObjectId> DrawKeyplans(List<Keyplans> keyplans, int floorcount)
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            //ArrayList kplist = new ArrayList();
            Dictionary<Extents3d, string> kplist = new Dictionary<Extents3d, string>();
            Dictionary<string, ObjectId> rectangids = new Dictionary<string, ObjectId>();
            //if we use the following values we get odd-shaped extents boxes!
            Point3d min = new Point3d(0, 0, 0);
            Point3d max = new Point3d(0, 0, 0);
            double lowestxvaluefound = 9999999999999;
            double lowestyvaluefound = 9999999999999;
            double highestxvaluefound = -9999999999999;
            double highestyvaluefound = -9999999999999;
            Extents3d tmpext = new Extents3d(min, max);
            Extents3d baseline = new Extents3d(min, max);
            //if we just create the ext object but don't assign a value, it might work better.
            for (int i = 0; i <= (MyCommands.Floors.Count - 1); i++)
            {
                //only find the keyplans on this floor
                var newDict = keyplans.Where(kp => kp.groupint == i).Select(kp => kp);
                foreach (Keyplans kp in newDict)
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        //convert the ObjectId into a blockreference.
                        BlockReference blkref = (BlockReference)tr.GetObject(kp.oid, OpenMode.ForRead);
                        //get the AttributeCollection from the BlockReference
                        AttributeCollection attrefids = blkref.AttributeCollection;
                        foreach (ObjectId attrefid in attrefids)
                        {
                            AttributeReference attref = tr.GetObject(attrefid, OpenMode.ForRead, false) as AttributeReference;
                            //compare the TextString of the attribute to the floor string of the group we want to add it to.
                            if (attref.TextString == MyCommands.Floors[i].ToString())
                            {
                                Extents3d ext = blkref.GeometricExtents;
                                min = ext.MinPoint;
                                max = ext.MaxPoint;
                                // no need to do the Z values as it should always be zero!
                                if (min.X < lowestxvaluefound) lowestxvaluefound = min.X;
                                if (min.Y < lowestyvaluefound) lowestyvaluefound = min.Y;
                                if (max.X > highestxvaluefound) highestxvaluefound = max.X;
                                if (max.Y > highestyvaluefound) highestyvaluefound = max.Y;
                                min = new Point3d(lowestxvaluefound, lowestyvaluefound, 0);
                                max = new Point3d(highestxvaluefound, highestyvaluefound, 0);
                                if (tmpext != baseline)
                                {
                                    tmpext.AddExtents(new Extents3d(min, max));
                                }
                                else
                                {
                                    tmpext = new Extents3d(min, max);
                                }
                                //scale the extents by 1% - we can maybe make this larger if needed.
                                Matrix3d trans = Matrix3d.Scaling(1.01, blkref.Position);
                                tmpext.TransformBy(trans);
                                tmpext.TransformBy(ed.CurrentUserCoordinateSystem.Inverse());
                                //reset our high and low values so we grab the high and low values of our blockreference:
                                lowestxvaluefound = 9999999999999;
                                lowestyvaluefound = 9999999999999;
                                highestxvaluefound = -9999999999999;
                                highestyvaluefound = -9999999999999;
                                break;
                            }
                        }
                        tr.Commit();
                    }
                }
                kplist.Add(tmpext, MyCommands.Floors[i].ToString());
                tmpext = baseline;
            }
            foreach (KeyValuePair<Extents3d, string> kp in kplist)
            {
                Extents3d newext = kp.Key;
                string floor = kp.Value;
                min = newext.MinPoint;
                max = newext.MaxPoint;
                //add  the new rectangle to the BlockTableRecord for Modelspace.
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Point3dCollection pts = new Point3dCollection();

                    Point3d ll = new Point3d(min.X, min.Y, min.Z);
                    Point3d ul = new Point3d(min.X, max.Y, min.Z);
                    Point3d ur = new Point3d(max.X, max.Y, max.Z);
                    Point3d lr = new Point3d(max.X, min.Y, max.Z);

                    pts.Add(ll);
                    pts.Add(ul);
                    pts.Add(ur);
                    pts.Add(lr);

                    Polyline3d rectang = new Polyline3d(Poly3dType.SimplePoly, pts, true);

                    BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    btr.AppendEntity(rectang);
                    tr.AddNewlyCreatedDBObject(rectang, true);
                    tr.Commit();
                    string kpname = floor;
                    rectangids.Add(kpname, rectang.ObjectId);
                }
                // need to regenerate drawing so we can zoom out.
                ed.Regen();
            }
            return rectangids;
        }       /// <summary>
                /// 
                /// </summary>
                /// <returns></returns>
        public Dictionary<string, ObjectId> DrawKeyplans()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            int grpnum;
            int i;
            Dictionary<string, ObjectId> rectangids = new Dictionary<string, ObjectId>();
            PromptStringOptions pso = new PromptStringOptions("Please enter the number of keyplans you require!");
            PromptResult res = ed.GetString(pso);
            if (res.Status == PromptStatus.OK)
            {
                grpnum = Convert.ToInt16(res.StringResult);
                for (i = 1; i <= grpnum; i++)
                {
                    //Group Newgroup = new Group("CH2M Layout Group", true);
                    //Polyline3d rectang = new Polyline3d();
                    // Get the window coordinates

                    PromptPointOptions ppo = new PromptPointOptions("\nSpecify first corner for keyplan rectangle:");

                    PromptPointResult ppr = ed.GetPoint(ppo);

                    if (ppr.Status != PromptStatus.OK)
                    {
                        MessageBox.Show("There was an error with your selection!");
                    }

                    Point3d min = ppr.Value;

                    PromptCornerOptions pco = new PromptCornerOptions("\nSpecify opposite corner for keyplan rectangle:", ppr.Value);

                    ppr = ed.GetCorner(pco);

                    if (ppr.Status != PromptStatus.OK)
                    {
                        MessageBox.Show("There was an error with your selection!");
                    }

                    Point3d max = ppr.Value;
                    //add  the new rectangle to the BlockTableRecord for Modelspace.
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        Point3dCollection pts = new Point3dCollection();

                        Point3d ll = new Point3d(min.X, min.Y, min.Z);
                        Point3d ul = new Point3d(min.X, max.Y, min.Z);
                        Point3d ur = new Point3d(max.X, max.Y, max.Z);
                        Point3d lr = new Point3d(max.X, min.Y, max.Z);

                        pts.Add(ll);
                        pts.Add(ul);
                        pts.Add(ur);
                        pts.Add(lr);

                        Polyline3d rectang = new Polyline3d(Poly3dType.SimplePoly, pts, true);

                        BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        btr.AppendEntity(rectang);
                        tr.AddNewlyCreatedDBObject(rectang, true);
                        tr.Commit();
                        string kpname = "k_" + Convert.ToString(i);
                        rectangids.Add(kpname, rectang.ObjectId);
                    }
                    // need to regenerate drawing so we can zoom out.
                    ed.Regen();
                }
            }
            return rectangids;
        }
    }
}