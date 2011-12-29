using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit;
using System.Diagnostics;
using System.IO;
using System.Windows.Media;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI.Events;

namespace opnRvt.App
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Automatic)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public class Ribbon : IExternalApplication
    {
        // ExternalCommands assembly path
        static string AddInPath = typeof(Ribbon).Assembly.Location;
        // Button icons directory
        static string ButtonIconsFolder = Path.GetDirectoryName(AddInPath);

        #region Cached Variables

        public static UIControlledApplication _cachedUiCtrApp;

        #endregion

        #region IExternalApplication Members

        public Autodesk.Revit.UI.Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // create customer Ribbon Items
                CreateRibbonPanel(application);

                return Autodesk.Revit.UI.Result.Succeeded;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Ribbon Sample");

                return Autodesk.Revit.UI.Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
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

        private RibbonPanel CreateRibbonPanel(UIControlledApplication application)
        {
            // create a Ribbon panel which contains three stackable buttons and one single push button.
            string firstPanelName = "openRevit Add-Ins";
            RibbonPanel ribbonPanel = application.CreateRibbonPanel(firstPanelName);

            //Create First button:
            PushButtonData pbDataParamToCSV = new PushButtonData("ParamToCSV", "Export Family \nParameters", 
                AddInPath, "opnRvt.Parameters.ParamToCSV");
            PushButton pbParamToCSV = ribbonPanel.AddItem(pbDataParamToCSV) as PushButton;
            pbParamToCSV.ToolTip = "Export Family Parameters To CSV File";
            pbParamToCSV.LargeImage = new BitmapImage(new Uri(Path.Combine(ButtonIconsFolder, "Resources\\export-excel-icon.png"), UriKind.Absolute));
            pbParamToCSV.Image = new BitmapImage(new Uri(Path.Combine(ButtonIconsFolder, "Resources\\export-excel-icon-s.png"), UriKind.Absolute));

            //Create Second button:
            PushButtonData pbDataAddParameterToFamily = new PushButtonData("AddParameterToFamily", "Bind Parameters \nTo Family", 
                AddInPath, "opnRvt.Parameters.AddParameterToFamily");
            PushButton pbAddParameterToFamily = ribbonPanel.AddItem(pbDataAddParameterToFamily) as PushButton;
            pbAddParameterToFamily.ToolTip = "Bind Shared Parameters From External File to Family";
            pbAddParameterToFamily.LargeImage = new BitmapImage(new Uri(Path.Combine(ButtonIconsFolder, "Resources\\table-import-icon.png"), UriKind.Absolute));
            pbAddParameterToFamily.Image = new BitmapImage(new Uri(Path.Combine(ButtonIconsFolder, "Resources\\table-import-icon-s.png"), UriKind.Absolute));

            
            return ribbonPanel;
        }
    }
}
