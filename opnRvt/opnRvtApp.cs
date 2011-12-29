#region Namespaces

using System;
using System.Text;
using System.Linq;
using System.Xml;
using System.Reflection;
using System.ComponentModel;
using System.Collections;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Windows.Forms;
using System.IO;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Plumbing;

using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.UI.Events;

using Autodesk.Revit.Collections;
using Autodesk.Revit.Exceptions;
using Autodesk.Revit.Utility;

using RvtApplication = Autodesk.Revit.ApplicationServices.Application;
using RvtDocument = Autodesk.Revit.DB.Document;

#endregion

namespace opnRvt
{
    [Transaction(TransactionMode.Manual)] 
    [Regeneration(RegenerationOption.Manual)] 
    public class opnRvtApp : IExternalApplication
    {
        #region Cached Variables

        public static UIControlledApplication _cachedUiCtrApp;
         
        #endregion

        #region IExternalApplication Members

        public Result OnStartup(UIControlledApplication uiApp)
        {
            _cachedUiCtrApp = uiApp;

            try
            {
                RibbonPanel ribbonPanel = CreateRibbonPanel();

                //TODO: add you code below.


                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                MessageBox.Show( ex.ToString() );
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication uiApp)
        {
            try
            {

                //TODO: add you code below.


                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return Result.Failed;
            }
        }

        #endregion

        #region Local Methods
        
        private RibbonPanel CreateRibbonPanel()
        {
            RibbonPanel panel = _cachedUiCtrApp.CreateRibbonPanel("opnRvt");

            ////Default button:
            PushButtonData pbDataParamToCSV = new PushButtonData("ParamToCSV", "Export Family \nParameters", Assembly.GetExecutingAssembly().Location, "opnRvt.Parameters.ParamToCSV");
            PushButton pbParamToCSV = panel.AddItem(pbDataParamToCSV) as PushButton; 
            pbParamToCSV.ToolTip = "ParamToCSV";
            pbParamToCSV.LargeImage = BmpImageSource("opnRvt.Resources.ParamToCSV32x32.bmp");
            pbParamToCSV.Image = BmpImageSource("opnRvt.Resources.ParamToCSV16x16.bmp");

            ////More buttons:
            PushButtonData pbDataAddParameterToFamily = new PushButtonData("AddParameterToFamily", "Bind Parameters \nTo Family",
                Assembly.GetExecutingAssembly().Location, "opnRvt.Parameters.AddParameterToFamily");
            PushButton pbAddParameterToFamily = panel.AddItem(pbDataAddParameterToFamily) as PushButton;
            pbAddParameterToFamily.ToolTip = "AddParameterToFamily";
            pbAddParameterToFamily.LargeImage = BmpImageSource("opnRvt.Resources.AddParameterToFamily32x32.bmp");
            pbParamToCSV.Image = BmpImageSource("opnRvt.Resources.AddParameterToFamily16x16.bmp");

            return panel;
        }

        private System.Windows.Media.ImageSource BmpImageSource(string embeddedPath)
        {
            Stream stream = this.GetType().Assembly.GetManifestResourceStream(embeddedPath);
            var decoder = new System.Windows.Media.Imaging.BmpBitmapDecoder(stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);
            
            return decoder.Frames[0];
        }

        #endregion
    }
}
