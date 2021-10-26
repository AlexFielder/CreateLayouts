// (C) Copyright 2021 by Alex Fielder 
//
using System;

namespace CreateLayouts
{
    public class Jobdetails
    {
        /// <summary>
        /// 
        /// </summary>
        public String Datestr;
        /// <summary>
        /// 
        /// </summary>
        public String Initials;
        /// <summary>
        /// 
        /// </summary>
        public String SiteNo;
        /// <summary>
        /// 
        /// </summary>
        public String BuildingNo;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="m_datestr"></param>
        /// <param name="m_Initials"></param>
        /// <param name="m_Siteno"></param>
        /// <param name="m_Buildingno"></param>
        public void New(String m_datestr,
            String m_Initials,
            String m_Siteno,
            String m_Buildingno
            )
        {
            Datestr = m_datestr;
            Initials = m_Initials;
            SiteNo = m_Siteno;
            BuildingNo = m_Buildingno;
        }
    }
}