// (C) Copyright 2021 by Alex Fielder 
//
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Linq;

namespace CreateLayouts
{
    public class ZoomCommands
    {
        /// <summary>
        /// Zoom to a window specified by the user
        /// </summary>
        [CommandMethod("CH2MPlugins", "ZW", CommandFlags.Modal)]
        public static void ZoomWindow()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Get the window coordinates

            PromptPointOptions ppo = new PromptPointOptions("\nSpecify first corner:");

            PromptPointResult ppr = ed.GetPoint(ppo);

            if (ppr.Status != PromptStatus.OK)
            {
                return;
            }

            Point3d min = ppr.Value;

            PromptCornerOptions pco = new PromptCornerOptions("\nSpecify opposite corner: ", ppr.Value);

            ppr = ed.GetCorner(pco);

            if (ppr.Status != PromptStatus.OK)
            {
                return;
            }

            Point3d max = ppr.Value;

            // Call out helper function
            // [Change this to ZoomWin2 or WoomWin3 to
            // use different zoom techniques]

            ZoomWindow(ed, min, max);
        }



        //[CommandMethod("ZE")]
        /// <summary>
        /// Zoom to the extents of an entity
        /// </summary>
        /// <param name="GroupId"></param>
        public static void ZoomToEntity(ObjectId GroupId)
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Extract its extents

            Extents3d ext;
            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                Entity ent = (Entity)tr.GetObject(GroupId, OpenMode.ForRead);
                ext = ent.GeometricExtents;
                tr.Commit();
            }

            ext.TransformBy(ed.CurrentUserCoordinateSystem.Inverse());

            // Call our helper function
            // [Change this to ZoomWin2 or WoomWin3 to
            // use different zoom techniques]

            ZoomWindow(ed, ext.MinPoint, ext.MaxPoint);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Group"></param>
        public static void ZoomToGroup(ObjectIdCollection Group)
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            Extents3d ext;
            ArrayList zgLower = new ArrayList();
            ArrayList zgUpper = new ArrayList();
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId oid in Group)
                {
                    //add the Point3d for each entity's extents to and array.
                    Entity ent = (Entity)tr.GetObject(oid, OpenMode.ForRead);
                    ext = ent.GeometricExtents;
                    zgLower.Add(ext.MinPoint);
                    zgUpper.Add(ext.MaxPoint);
                }
                tr.Commit();
            }

            /* then sort each array accordingly and extract the highest value from zgmax and
             the lowest value from zgmin */
            Point3d zgmin = (from Point3d zgm in zgLower
                             orderby zgm descending
                             select zgm).Last();
            zgmin.TransformBy(ed.CurrentUserCoordinateSystem.Inverse());
            Point3d zgmax = (from Point3d zgu in zgUpper
                             orderby zgu ascending
                             select zgu).First();
            zgmax.TransformBy(ed.CurrentUserCoordinateSystem.Inverse());
            ZoomWindow(ed, zgmin, zgmax);
        }
        // Helper functions to zoom using different techniques
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ed"></param>
        /// <param name="min"></param>
        /// <param name="max"></param>
        public static void ZoomWindow(Editor ed, Point3d min, Point3d max)
        {
            Point2d min2d = new Point2d(min.X, min.Y);
            Point2d max2d = new Point2d(max.X, max.Y);
            ViewTableRecord view = new ViewTableRecord();
            view.CenterPoint = min2d + ((max2d - min2d) / 2.0);
            view.Height = max2d.Y - min2d.Y;
            view.Width = max2d.X - min2d.X;
            ed.SetCurrentView(view);
        }

        // Zoom via COM

        //private static void ZoomWin2(Editor ed, Point3d min, Point3d max)
        //{
        //    AcadApplication app = (AcadApplication)AcadApp.AcadApplication;

        //    double[] lower = new double[3] { min.X, min.Y, min.Z };
        //    double[] upper = new double[3] { max.X, max.Y, max.Z };

        //    app.ZoomWindow(lower, upper);
        //}

        // Zoom by sending a command

        //private static void ZoomWin3(Editor ed, Point3d min, Point3d max)
        //{
        //    string lower = min.ToString().Substring(1, min.ToString().Length - 2);
        //    string upper = max.ToString().Substring(1, max.ToString().Length - 2);

        //    string cmd = "_.ZOOM _W " + lower + " " + upper + " ";

        //    // Call the command synchronously using COM

        //    AcadApplication app = (AcadApplication)AcadApp.AcadApplication;

        //    app.ActiveDocument.SendCommand(cmd);

        //    // Could also use async command calling:
        //    //ed.Document.SendStringToExecute(
        //    //  cmd, true, false, true
        //    //);
        //}
    }
}