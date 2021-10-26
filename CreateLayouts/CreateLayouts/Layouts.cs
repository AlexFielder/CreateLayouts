// (C) Copyright 2021 by Alex Fielder 
//
using Autodesk.AutoCAD.DatabaseServices;

namespace CreateLayouts
{
    public class Layouts
    {
        /// <summary>
        /// 
        /// </summary>
        public string LayoutName { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public ObjectId LayoutId { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public int LayoutNo { get; set; }
    }
}