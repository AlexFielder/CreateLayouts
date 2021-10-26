// (C) Copyright 2021 by Alex Fielder 
//
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace CreateLayouts
{
    internal class UCSTools
    {
        /// <summary>
        /// Removes any UCSTableRecords that have a name matching the string "UCS_*"
        /// </summary>
        public static void RemoveCustomUCS()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            using (Database tmpdb = doc.Database)
            {
                using (Transaction tmptr = tmpdb.TransactionManager.StartTransaction())
                {
                    UcsTable ucstbl = tmptr.GetObject(tmpdb.UcsTableId, OpenMode.ForRead) as UcsTable;
                    foreach (ObjectId UcsId in ucstbl)
                    {
                        UcsTableRecord tmpUcs = tmptr.GetObject(UcsId, OpenMode.ForRead) as UcsTableRecord;
                        if (Regex.IsMatch(tmpUcs.Name.ToString(), "UCS_*"))
                        {
                            UcsTableRecord HasUcs = (UcsTableRecord)tmptr.GetObject(ucstbl[tmpUcs.Name], OpenMode.ForWrite, true);
                            HasUcs.Erase();
                        }
                    }
                    tmptr.Commit();
                }
            }
        }

        /// <summary>
        /// Recreates our custom UCSTableRecord entries
        /// </summary>
        /// <param name="UCSName">The name of the UCSTableRecord Entry to erase/recreate</param>
        /// <param name="UCSOrigin">The Point3d Origin of the new UCSTableRecord</param>
        /// <param name="UCSXAxispt">The Vector3d XAxis of the new UCSTableRecord</param>
        /// <param name="UCSYAxispt">The Vector3d YAxis of the new UCSTableRecord</param>
        internal static string GetorCreateUCS(String UCSName, Point3d UCSOrigin, Vector3d UCSXAxispt, Vector3d UCSYAxispt)
        {
            ObjectId UcsId;
            //the return value for this function
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            using (Database db = doc.Database)
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    UcsTable lt = tr.GetObject(db.UcsTableId, OpenMode.ForRead) as UcsTable;
                    Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                    try
                    {
                        if (lt.Has(UCSName))
                        // we should then check whether it matches the origin, orientation of our blocks (because the blocks could have been moved!)
                        {
                            UcsTableRecord ucstblrec = (UcsTableRecord)tr.GetObject(lt[UCSName], OpenMode.ForRead);
                            if (ucstblrec.Origin != UCSOrigin)
                            {
                                //now we can either delete the ucs or move it to the correct location!
                                ucstblrec.UpgradeOpen();
                                ucstblrec.Origin = UCSOrigin;
                                ucstblrec.XAxis = UCSXAxispt;
                                ucstblrec.YAxis = UCSYAxispt;
                                ed.UpdateScreen();
                                db.Regenmode = true;
                                //tr.Commit();
                            }
                            else
                            {
                                // the existing UCS matches the origin of our block, we should still check the orientation though.
                                if (ucstblrec.XAxis != UCSXAxispt)
                                {
                                    ucstblrec.UpgradeOpen();
                                    //there should be no need to change the YAxis point as the UCS needs to maintain 90° between the two point.
                                    ucstblrec.XAxis = UCSXAxispt;
                                    //tr.Commit();
                                }
                            }
                        }
                        else
                        {
                            // we need to create the ucs for the new view.
                            UcsTableRecord ltr = new UcsTableRecord();
                            ltr.Name = UCSName;
                            ltr.Origin = UCSOrigin;
                            ltr.XAxis = UCSXAxispt;
                            ltr.YAxis = UCSYAxispt;
                            // it doesn't exist so add it, but first upgrade the open to write
                            lt.UpgradeOpen();
                            UcsId = lt.Add(ltr);
                            tr.AddNewlyCreatedDBObject(ltr, true);
                            db.UcsBase = UcsId;
                        }
                        ed.UpdateScreen();
                        db.Regenmode = true;
                        tr.Commit();
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    {
                        tr.Abort();
                        MessageBox.Show(ex.ToString());
                    }
                    finally
                    {
                        tr.Dispose();
                    }
                    return UCSName;
                }
            }
            //return UcsId;
        }

        private static bool EraseUCS(string UCSName)
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                UcsTable lt = tr.GetObject(db.UcsTableId, OpenMode.ForRead) as UcsTable;
                if (lt.Has(UCSName))
                {
                    UcsTableRecord HasUcs = (UcsTableRecord)tr.GetObject(lt[UCSName], OpenMode.ForWrite, true);
                    HasUcs.Erase();
                    tr.Commit();
                    return true;
                }
                else
                {
                    tr.Commit();
                    return false;
                }
            }
        }

        internal static bool SetCurrentUCS(int p)
        {
            bool ucsset = false;
            string UCSName = "UCS_" + Convert.ToString(p);
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            using (Database db = doc.Database)
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    UcsTable ucstbl = tr.GetObject(db.UcsTableId, OpenMode.ForRead) as UcsTable;
                    Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                    try
                    {

                        if (ucstbl.Has(UCSName))
                        // we should then set it current for this view?
                        {
                            UcsTableRecord ucstblrec = (UcsTableRecord)tr.GetObject(ucstbl[UCSName], OpenMode.ForRead);
                            Matrix3d ucsMat = Matrix3d.AlignCoordinateSystem(Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis, ucstblrec.Origin,
                                ucstblrec.XAxis, ucstblrec.YAxis, ucstblrec.XAxis.CrossProduct(ucstblrec.YAxis));
                            ed.CurrentUserCoordinateSystem = ucsMat;
                            ed.Regen();
                            ed.UpdateScreen();
                            db.Regenmode = true;
                        }
                        tr.Commit();
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    {
                        tr.Abort();
                        MessageBox.Show(ex.ToString());
                    }
                    finally
                    {
                        tr.Dispose();
                    }
                }
            }
            return ucsset;
        }

        internal static ObjectId GetCurrentUCSId(string UCSName)
        {
            ObjectId ucsid = ObjectId.Null;
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                UcsTable ucstbl = tr.GetObject(db.UcsTableId, OpenMode.ForRead) as UcsTable;
                Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                try
                {

                    if (ucstbl.Has(UCSName))
                    {
                        UcsTableRecord utr = (UcsTableRecord)tr.GetObject(ucstbl[UCSName], OpenMode.ForRead);
                        ucsid = utr.ObjectId;
                    }
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                tr.Commit();
            }
            return ucsid;
        }
    }
}