// (C) Copyright 2021 by Alex Fielder 
//
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace CreateLayouts
{
    public class Vports
    {
        /// <summary>
        /// Viewport name.
        /// </summary>
        public string vpName { get; set; }
        /// <summary>
        /// Viewport Custom Scale
        /// </summary>
        public string vpScale { get; set; }
        /// <summary>
        /// Viewport Paperspace centrepoint
        /// </summary>
        public Point3d vpPSCentrePoint { get; set; }
        /// <summary>
        /// Viewport Modelspace centrepoint
        /// </summary>
        public Point2d vpMSCentrePoint { get; set; }
        /// <summary>
        /// Full or Partial intended view
        /// </summary>
        public string FullorPartial { get; set; }
        /// <summary>
        /// Floor name
        /// </summary>
        public string FloorName { get; set; }
        /// <summary>
        /// Site Number
        /// </summary>
        public string SiteNo { get; set; }
        /// <summary>
        /// Building Number
        /// </summary>
        public string BldgNo { get; set; }
        /// <summary>
        /// Date Drawn
        /// </summary>
        public string DateDrawn { get; set; }
        /// <summary>
        /// Drawn By
        /// </summary>
        public string DrawnBy { get; set; }
        /// <summary>
        /// Viewport Visibility Property
        /// </summary>
        public string vpVisProp { get; set; }
        /// <summary>
        /// Viewport Height
        /// </summary>
        public double vpHeight { get; set; }
        /// <summary>
        /// Viewport Width
        /// </summary>
        public double vpWidth { get; set; }
        /// <summary>
        /// Page Size (A3/A2/A1/A0)
        /// </summary>
        public string PageSize { get; set; }
        /// <summary>
        /// Id of the Viewport Block - so we can find it again later
        /// </summary>
        public ObjectId blockid { get; set; }
        /// <summary>
        /// The alignment of the Block (rotation)
        /// </summary>
        public Vector3d Xaxis { get; set; }
        /// <summary>
        /// The alignment of the Block (rotation)
        /// </summary>
        public Vector3d Yaxis { get; set; }
        /// <summary>
        /// The viewport number
        /// </summary>
        public int vpNumber { get; set; }
        /// <summary>
        /// used to store the name of the custom UCS associated with this block
        /// </summary>
        public string UCSName { get; set; }
    }
}