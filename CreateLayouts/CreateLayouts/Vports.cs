// (C) Copyright 2021 by Alex Fielder 
//
using Autodesk.AutoCAD.DatabaseServices;

namespace CreateLayouts
{
    public class Vports
    {
        internal ObjectId blockid;

        public string FloorName { get; internal set; }
    }
}