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
    public class ExtCmd1 : IExternalCommand
    {
        #region Cached Variables

        private static ExternalCommandData _cachedCmdData;
        
        public static UIApplication CachedUiApp
        {
            get
            {
                return _cachedCmdData.Application;
            }
        }

        public static RvtApplication CachedApp
        {
            get
            {
                return CachedUiApp.Application;
            }
        }

        public static RvtDocument CachedDoc
        {
            get
            {
                return CachedUiApp.ActiveUIDocument.Document;
            }
        }
        
        #endregion

        #region IExternalCommand Members         

        public Result Execute(ExternalCommandData cmdData, ref string msg, ElementSet elemSet)
        {
            _cachedCmdData = cmdData;

            try
            {
                //TODO: add your code below.


                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                msg = ex.ToString();
                return Result.Failed;
            }
        }

        #endregion
        
        public static ICollection<ElementId> FilterElementParameter(RvtDocument doc)
        {
            ParameterValueProvider provider = new ParameterValueProvider(new ElementId((int)BuiltInParameter.RBS_PIPE_DIAMETER_PARAM));
        
            FilterIntegerRule rule1 = new FilterIntegerRule(provider, new FilterNumericEquals(), 2);
            ElementParameterFilter filter1 = new ElementParameterFilter(rule1);
               
            return (new FilteredElementCollector(doc))
                    .OfClass(typeof(Autodesk.Revit.DB.Plumbing.Pipe))
                    .WherePasses(filter1)
                    .ToElementIds();
        }
        
        
        public static void RawSetParameterValue(Parameter p, object value)
        {
            try
            {
                if (value.GetType().Equals(typeof(string)))
                {
                    if (p.SetValueString(value as string))
                        return;
                }
        
                switch (p.StorageType)
                {
                    case StorageType.None:
                        break;
                    case StorageType.Double:
                        if (value.GetType().Equals(typeof(string)))
                        {
                            p.Set(double.Parse(value as string));
                        }
                        else
                        {
                            p.Set(Convert.ToDouble(value));
                        }
                        break;
                    case StorageType.Integer:
                        if (value.GetType().Equals(typeof(string)))
                        {
                            p.Set(int.Parse(value as string));
                        }
                        else
                        {
                            p.Set(Convert.ToInt32(value));
                        }
                        break;
                    case StorageType.ElementId:
                        if (value.GetType().Equals(typeof(ElementId)))
                        {
                            p.Set(value as ElementId);
                        }
                        else if (value.GetType().Equals(typeof(string)))
                        {
                            p.Set(new ElementId(int.Parse(value as string)));
                        }
                        else
                        {
                            p.Set(new ElementId(Convert.ToInt32(value)));
                        }
                        break;
                    case StorageType.String:
                        p.Set(value.ToString());
                        break;
                }
            }
            catch
            {
                throw new Exception("Invalid Value Input!");
            }
        }
    }
}
