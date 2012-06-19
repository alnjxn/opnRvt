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

namespace opnRvt.Parameters
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ParamRecover : IExternalCommand
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
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.InitialDirectory = Convert.ToString(Environment.SpecialFolder.MyDocuments);
                saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";
                saveFileDialog.FilterIndex = 1;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (CachedDoc.IsFamilyDocument)
                    {
                        List<SharedParameterInfo> sharedParamsInfo = GetSharedParametersInfo(CachedDoc);
                        using (StreamWriter sw = new StreamWriter(saveFileDialog.FileName))
                        {
                            string title = string.Empty;
                            string rows = ParamToCSV.RawParametersInfoToCSVString(sharedParamsInfo, ref title);
                            sw.WriteLine(title);
                            sw.Write(rows);
                        }

                    }
                }

                MessageBox.Show("Wrote file to " + saveFileDialog.FileName, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                msg = ex.ToString();
                return Result.Failed;
            }
        }

        #endregion


        public class RawSharedParameterInfo
        {
            public static string FileName { get; set; }
            public string Name { get; set; }
            public string GUID { get; set; }
            public string Owner { get; set; }
            public BuiltInParameterGroup Group { get; set; }
            public ParameterType Type { get; set; }
            public bool ReadOnly { get; set; }
            public bool Visible { get; set; }
        }

        public class SharedParameterInfo
        {
            public static string FileName { get; set; }
            public string Name { get; set; }
            public string GUID { get; set; }
            public string Owner { get; set; }
            public BuiltInParameterGroup Group { get; set; }
            public ParameterType Type { get; set; }
            public bool ReadOnly { get; set; }
            public bool Visible { get; set; }
        }



        public static List<SharedParameterInfo> GetSharedParametersInfo(Document doc)
        {
            List<SharedParameterInfo> paramList =
                (from FamilyParameter p in doc.FamilyManager.Parameters
                 select new SharedParameterInfo
                 {
                     Name = p.Definition.Name
                 }).ToList();

            return paramList;
        }

        public static List<RawSharedParameterInfo> RawGetSharedParametersInfo(DefinitionFile defFile)
        {
            RawSharedParameterInfo.FileName = defFile.Filename;

            List<RawSharedParameterInfo> paramList =
                (from DefinitionGroup dg in defFile.Groups
                 from ExternalDefinition d in dg.Definitions
                 select new RawSharedParameterInfo
                 {
                     Name = d.Name,
                     GUID = d.GUID.ToString(),
                     Owner = d.OwnerGroup.Name,
                     Group = d.ParameterGroup,
                     Type = d.ParameterType,
                     ReadOnly = d.IsReadOnly,
                     Visible = d.Visible,

                 }).ToList();
            return paramList;
        }




    }
}

