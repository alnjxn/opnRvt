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
    public class ParamToCSV : IExternalCommand
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
                        List<FamilyParameterInfo> famParamsInfo = GetFamilyParametersInfo(CachedDoc);
                        using (StreamWriter sw = new StreamWriter(saveFileDialog.FileName))
                        {
                            string title = string.Empty;
                            string rows = RawParametersInfoToCSVString(famParamsInfo, ref title);
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

        public class FamilyParameterInfo
        {
            public string Name { get; set; }
            //public string Value { get; set; } // Parameter Only!
            public string ElementID { get; set; }
            public string GUID { get; set; }
            public string BuiltinID { get; set; }
            public BuiltInParameterGroup Group { get; set; }
            public ParameterType Type { get; set; }
            public StorageType Storage { get; set; }
            public string Unit { get; set; } //DisplayUnitType doesn't work!
            public bool Shared { get; set; }    // Both but different handling!
            public bool Instance { get; set; }  // FamilyParameter Only!
            public bool Reporting { get; set; }  // FamilyParameter Only!
            public bool FormulaDetermined { get; set; } // FamilyParameter Only!
            public bool CanAssignFormula { get; set; }
            public bool ReadOnly { get; set; }
        }

        public class RawProjectParameterInfo
        {
            public static string FileName { get; set; }
            public string Name { get; set; }
            public BuiltInParameterGroup Group { get; set; }
            public ParameterType Type { get; set; }
            public bool ReadOnly { get; set; }
            public bool BoundToInstance { get; set; }
            public string[] BoundCategories { get; set; }
            public bool FromShared { get; set; }
            public string GUID { get; set; }
            public string Owner { get; set; }
            public bool Visible { get; set; }
        }

        public static string GetDUTString(FamilyParameter p)
        {
            string unitType = string.Empty;
            try { unitType = p.DisplayUnitType.ToString(); }
            catch { }

            return unitType;
        }
        public static List<FamilyParameterInfo> GetFamilyParametersInfo(Document doc)
        {
            List<FamilyParameterInfo> paramList =
                (from FamilyParameter p in doc.FamilyManager.Parameters
                 select new FamilyParameterInfo
                 {
                     Name = p.Definition.Name,
                     //Value = p.AsValueString(),   //N/A
                     ElementID = p.Id.IntegerValue.ToString(),
                     GUID = (p.Definition as ExternalDefinition) != null ?
                        (p.Definition as ExternalDefinition).GUID.ToString() : string.Empty,
                     BuiltinID = (p.Definition as InternalDefinition) != null ?
                        (p.Definition as InternalDefinition).BuiltInParameter.ToString() : string.Empty,
                     Group = p.Definition.ParameterGroup,
                     Type = p.Definition.ParameterType,
                     Storage = p.StorageType,
                     Unit = GetDUTString(p),
                     Shared = (p.Definition as ExternalDefinition) != null,
                     Instance = p.IsInstance,
                     Reporting = p.IsReporting,
                     FormulaDetermined = p.IsDeterminedByFormula,
                     CanAssignFormula = p.CanAssignFormula,
                     ReadOnly = p.IsReadOnly,
                 }).ToList();

            return paramList;
        }
        public static List<RawProjectParameterInfo> RawGetProjectParametersInfo(Document doc)
        {
            RawProjectParameterInfo.FileName = doc.Title;
            List<RawProjectParameterInfo> paramList = new List<RawProjectParameterInfo>();

            BindingMap map = doc.ParameterBindings;
            DefinitionBindingMapIterator it = map.ForwardIterator();
            it.Reset();
            while (it.MoveNext())
            {
                ElementBinding eleBinding = it.Current as ElementBinding;
                InstanceBinding insBinding = eleBinding as InstanceBinding;
                Definition def = it.Key;
                if (def != null)
                {
                    ExternalDefinition extDef = def as ExternalDefinition;
                    bool shared = extDef != null;
                    RawProjectParameterInfo param = new RawProjectParameterInfo
                    {
                        Name = def.Name,
                        Group = def.ParameterGroup,
                        Type = def.ParameterType,
                        ReadOnly = def.IsReadOnly,
                        BoundToInstance = insBinding != null,
                        BoundCategories = RawConvertSetToList<Category>(eleBinding.Categories).Select(c => c.Name).ToArray(),
                        FromShared = shared,
                        GUID = shared ? extDef.GUID.ToString() : string.Empty,
                        Owner = shared ? extDef.OwnerGroup.Name : string.Empty,
                        Visible = shared ? extDef.Visible : true,

                    };

                    paramList.Add(param);
                }
            }

            return paramList;
        }

        public static List<T> RawConvertSetToList<T>(IEnumerable set)
        {
            List<T> list = (from T p in set select p).ToList<T>();
            return list;
        }


        public static string RawParametersInfoToCSVString<T>(List<T> infoList, ref string title)
        {
            StringBuilder sb = new StringBuilder();
            PropertyInfo[] propInfoArrary = typeof(T).GetProperties();
            foreach (PropertyInfo pi in propInfoArrary)
            {
                title += pi.Name + ",";
            }
            title = title.Remove(title.Length - 1);

            foreach (T info in infoList)
            {
                foreach (PropertyInfo pi in propInfoArrary)
                {
                    object obj = info.GetType().InvokeMember(pi.Name, BindingFlags.GetProperty, null, info, null);
                    IList list = obj as IList;
                    if (list != null)
                    {
                        string str = string.Empty;
                        foreach (object e in list)
                        {
                            str += e.ToString() + ";";
                        }
                        str = str.Remove(str.Length - 1);

                        sb.Append(str + ",");
                    }
                    else
                    {
                        sb.Append((obj == null ? string.Empty : obj.ToString()) + ",");
                    }
                }
                sb.Remove(sb.Length - 1, 1).Append(Environment.NewLine);
            }

            return sb.ToString();
        }

    }
}