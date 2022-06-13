// (C) Copyright 2022 by 
//
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;

[assembly: ExtensionApplication(typeof(AutoCAD2023Plugin.ExportToExcelPlugin))]

namespace AutoCAD2023Plugin
{

    public class ExportToExcelPlugin : IExtensionApplication
    {

        void IExtensionApplication.Initialize()
        {   
            
        }

        void IExtensionApplication.Terminate()
        {
            
        }

    }

}
