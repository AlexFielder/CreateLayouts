// (C) Copyright 2021 by Alex Fielder 
//
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Runtime.InteropServices;

namespace CreateLayouts
{
    public class DBUtils
    {
        // This is a managed workaround for getting a non-erased
        // SymbolTableRecord, when there are also erased ones with
        // the same name:

        public static ObjectId GetTableRecordId(ObjectId TableId, string Name)
        {
            ObjectId id = ObjectId.Null;
            using (Transaction tr = TableId.Database.TransactionManager.StartTransaction())
            {
                SymbolTable table = (SymbolTable)tr.GetObject(TableId, OpenMode.ForRead);
                if (table.Has(Name))
                {
                    id = table[Name];
                    if (!id.IsErased)
                        return id;
                    foreach (ObjectId recId in table)
                    {
                        if (!recId.IsErased)
                        {
                            SymbolTableRecord rec = (SymbolTableRecord)tr.GetObject(recId, OpenMode.ForRead);
                            if (string.Compare(rec.Name, Name, true) == 0)
                                return recId;
                        }
                    }
                }
            }
            return id;
        }

        // This is a much better/faster solution that P/Invokes
        // AcDbSymbolTableRecord::getAt() directly from managed code:


        public static class AcDbSymbolTable
        {
            // Acad::ErrorStatus AcDbSymbolTable::getAt(wchar_t const *, class AcDbObjectId &, bool)

            [System.Security.SuppressUnmanagedCodeSecurity]
            [DllImport("acdb17.dll", CallingConvention = CallingConvention.ThisCall, CharSet = CharSet.Unicode,
               EntryPoint = "?getAt@AcDbSymbolTable@@QBE?AW4ErrorStatus@Acad@@PB_WAAVAcDbObjectId@@_N@Z")]

            public static extern ErrorStatus getAt(IntPtr symbolTable, string name, out ObjectId id, bool getErased);
        }

        public static ObjectId GetSymbolTableRecordId(SymbolTable table, string name)
        {
            ObjectId id = ObjectId.Null;
            ErrorStatus es = AcDbSymbolTable.getAt(table.UnmanagedObject, name, out id, false);
            return id;
        }

        public static ObjectId GetSymbolTableRecordId(ObjectId TableId, string name)
        {
            using (Transaction tr = TableId.Database.TransactionManager.StartTransaction())
            {
                SymbolTable table = (SymbolTable)tr.GetObject(TableId, OpenMode.ForRead);
                try
                {
                    return GetSymbolTableRecordId(table, name);
                }
                finally
                {
                    tr.Commit();
                }
            }
        }
        public static ObjectIdCollection GetBlockReferenceIds(ObjectId btrId)
        {
            if (btrId.IsNull)
                throw new ArgumentException("null object id");
            ObjectIdCollection result = new ObjectIdCollection();
            using (Transaction trans = btrId.Database.TransactionManager.StartTransaction())
            {
                try
                {
                    // open the BlockTableRecord: 

                    BlockTableRecord btr = trans.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                    if (btr != null)
                    {
                        // Add the ids of all references to the BTR. Some dynamic blocks may
                        // reference the dynamic block directly rather than an anonymous block,
                        // so this will get those as well:

                        ObjectIdCollection blockRefIds = btr.GetBlockReferenceIds(true, false);
                        if (blockRefIds != null)
                        {
                            foreach (ObjectId id in blockRefIds)
                                result.Add(id);
                        }

                        // if this is not a dynamic block, we're done:

                        if (!btr.IsDynamicBlock)
                            return result;

                        // Get the ids of all anonymous block table records for the dynamic block

                        ObjectIdCollection anonBtrIds = btr.GetAnonymousBlockIds();
                        if (anonBtrIds != null)
                        {
                            foreach (ObjectId anonBtrId in anonBtrIds)
                            {
                                // get all references to each anonymous block:

                                BlockTableRecord rec = trans.GetObject(anonBtrId, OpenMode.ForRead) as BlockTableRecord;
                                if (rec != null)
                                {
                                    blockRefIds = rec.GetBlockReferenceIds(false, true);
                                    if (blockRefIds != null)
                                    {
                                        foreach (ObjectId id in blockRefIds)
                                            result.Add(id);
                                    }
                                }
                            }
                        }
                    }
                }
                finally
                {
                    trans.Commit();
                }
            }
            return result;
        }
    }
}