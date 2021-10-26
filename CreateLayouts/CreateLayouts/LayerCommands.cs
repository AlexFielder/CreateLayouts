// (C) Copyright 2021 by Alex Fielder 
//
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace CreateLayouts
{
    public class LayerCommands
    {
        //[CommandMethod ("SaveLayers")]
        /// <summary>
        /// Create/Save the "test" layerstate.
        /// </summary>
        public void SaveLayerstate()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            LayerStateManager lsm = db.LayerStateManager;
            string dwgname = "test";
            lsm.SaveLayerState(dwgname, LayerStateMasks.Frozen | LayerStateMasks.Locked | LayerStateMasks.On, ObjectId.Null);
        }
        //[CommandMethod("RestoreLayers")]
        /// <summary>
        /// restore the "test" layerstate.
        /// </summary>
        public void RestoreLayerstate()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            LayerStateManager lsm = db.LayerStateManager;
            string dwgname = "test";
            lsm.RestoreLayerState(dwgname, ObjectId.Null, 0, LayerStateMasks.Frozen | LayerStateMasks.Locked | LayerStateMasks.On);
        }
        //[CommandMethod("DeleteLayerstates")]
        /// <summary>
        /// Delete the "test" layerstate.
        /// </summary>
        public void Deletelayerstates()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            LayerStateManager lsm = db.LayerStateManager;
            string dwgname = "test";

            try
            {
                lsm.DeleteLayerState(dwgname);
                ed.WriteMessage("Layer state " + dwgname + " was deleted!");
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                MessageBox.Show(ex.StackTrace);
            }

        }
    }
}