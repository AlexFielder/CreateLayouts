// (C) Copyright 2021 by Alex Fielder 
//
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;

namespace CreateLayouts
{
    public class BlockHelperClass
    {
        /// <summary>
        /// Get the ObjectId of the block we're looking for.
        /// </summary>
        /// <param name="db">The Current Database.</param>
        /// <param name="Name">The Name of the Block we're looking for.</param>
        /// <returns>Returns an ObjectId related to the Block Name.</returns>
        public static ObjectId GetBlockId(Database db, string Name)
        {
            ObjectId id = ObjectId.Null;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable blocks = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (blocks.Has(Name))
                {
                    id = blocks[Name];
                    if (id.IsErased == false)
                    {
                        foreach (ObjectId btrId in blocks)
                        {
                            if (!id.IsErased)
                            {
                                if (!id.IsEffectivelyErased)
                                {
                                    BlockTableRecord rec = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                                    if (string.Compare(rec.Name, Name, true) == 0)
                                    {
                                        id = btrId;
                                        break;
                                    }
                                }
                            }
                        }
                    }

                }
                else
                {
                    // we should add the block to the drawing?
                    BlockHelperClass.ImportBlocks();
                }
                tr.Commit();
            }
            return id;
        }
        /// <summary>
        /// Imports blocks from our base drawing - currently resides on the desktop of the user.
        /// This will need moving to the server once testing/coding is complete.
        /// </summary>
        public static void ImportBlocks()
        {
            //string ViewportBlockdwg = "C:\\users\\" + Environment.UserName + "\\Desktop\\VIEWPORT-Block.dwg";
            string ViewportBlockdwg = "C:\\Program Files\\CH2M HILL\\CreateLayouts For AutoCAD 2009\\VIEWPORT-Block.dwg";
            DocumentCollection dm =
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            Editor ed = dm.MdiActiveDocument.Editor;
            Database destDb = dm.MdiActiveDocument.Database;
            Database sourceDb = new Database(false, true);
            String sourceFileName;
            try
            {
                // Get name of DWG from which to copy blocks
                sourceFileName = ViewportBlockdwg;
                //ed.GetString("\nEnter the name of the source drawing: ");
                // Read the DWG into a side database
                sourceDb.ReadDwgFile(sourceFileName,
                                    System.IO.FileShare.Read,
                                    true,
                                    "");

                // Create a variable to store the list of block identifiers
                ObjectIdCollection blockIds = new ObjectIdCollection();

                Autodesk.AutoCAD.DatabaseServices.TransactionManager tm =
                  sourceDb.TransactionManager;

                using (Transaction myT = tm.StartTransaction())
                {
                    // Open the block table
                    BlockTable bt =
                        (BlockTable)tm.GetObject(sourceDb.BlockTableId,
                                                OpenMode.ForRead,
                                                false);

                    // Check each block in the block table
                    foreach (ObjectId btrId in bt)
                    {
                        BlockTableRecord btr =
                          (BlockTableRecord)tm.GetObject(btrId,
                                                        OpenMode.ForRead,
                                                        false);
                        // Only add named & non-layout blocks to the copy list
                        if (!btr.IsAnonymous && !btr.IsLayout)
                            blockIds.Add(btrId);
                        btr.Dispose();
                    }
                    bt.Dispose();
                    // No need to commit the transaction
                    myT.Dispose();
                }
                // Copy blocks from source to destination database
                IdMapping mapping = new IdMapping();
                sourceDb.WblockCloneObjects(blockIds,
                                            destDb.BlockTableId,
                                            mapping,
                                            DuplicateRecordCloning.Replace,
                                            false);
                ed.WriteMessage("\nCopied "
                                + blockIds.Count.ToString()
                                + " block definitions from "
                                + sourceFileName
                                + " to the current drawing.");
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("\nError during copy: " + ex.Message);
            }
            sourceDb.Dispose();
        }
    }
}