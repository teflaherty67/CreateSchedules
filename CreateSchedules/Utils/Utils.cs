using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace CreateSchedules
{
    internal static class Utils
    {
        public static List<View> GetAllElevationViews(Document doc)
        {
            List<View> returnList = new List<View>();

            FilteredElementCollector colViews = new FilteredElementCollector(doc);
            colViews.OfClass(typeof(View));

            // loop through views and check for elevation views
            foreach (View x in colViews)
            {
                if (x.GetType() == typeof(ViewSection))
                {
                    if (x.IsTemplate == false)
                    {
                        if (x.ViewType == ViewType.Elevation)
                        {
                            // add view to list
                            returnList.Add(x);
                        }
                    }
                }
            }

            return returnList;
        }

        internal static string GetParameterValueByName(Element element, string paramName)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);

            if (paramList != null)
                try
                {
                    Parameter param = paramList[0];
                    string paramValue = param.AsValueString();
                    return paramValue;
                }
                catch (System.ArgumentOutOfRangeException)
                {
                    return null;
                }

            return "";
        }

        internal static void SetParameterByName(Element element, string paramName, string value)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);

            if (paramList != null)
            {
                Parameter param = paramList[0];

                param.Set(value);
            }
        }

        internal static void SetParameterByName(Element element, string paramName, int value)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);

            if (paramList != null)
            {
                Parameter param = paramList[0];

                param.Set(value);
            }
        }

        internal static BitmapImage BitmapToImageSource(Bitmap bm)
        {
            using (MemoryStream mem = new MemoryStream())
            {
                bm.Save(mem, System.Drawing.Imaging.ImageFormat.Png);
                mem.Position = 0;
                BitmapImage bmi = new BitmapImage();
                bmi.BeginInit();
                bmi.StreamSource = mem;
                bmi.CacheOption = BitmapCacheOption.OnLoad;
                bmi.EndInit();

                return bmi;
            }
        }

        internal static RibbonPanel CreateRibbonPanel(UIControlledApplication app, string tabName, string panelName)
        {
            RibbonPanel curPanel = GetRibbonPanelByName(app, tabName, panelName);

            if (curPanel == null)
                curPanel = app.CreateRibbonPanel(tabName, panelName);

            return curPanel;
        }

        private static RibbonPanel GetRibbonPanelByName(UIControlledApplication app, string tabName, string panelName)
        {
            foreach (RibbonPanel tempPanel in app.GetRibbonPanels(tabName))
            {
                if (tempPanel.Name == panelName)
                    return tempPanel;
            }

            return null;
        }

        internal static string GetStringBetweenCharacters(string input, string charFrom, string charTo)
        {
            //string cleanInput = CleanSheetNumber(input);

            int posFrom = input.IndexOf(charFrom);
            if (posFrom != -1) //if found char
            {
                int posTo = input.IndexOf(charTo, posFrom + 1);
                if (posTo != -1) //if found char
                {
                    return input.Substring(posFrom + 1, posTo - posFrom - 1);
                }
            }

            return string.Empty;
        }

        internal static string CleanSheetNumber(string sheetNumber)
        {
            Regex regex = new Regex(@"[^a-zA-Z0-9\s]", (RegexOptions)0);
            string replaceText = regex.Replace(sheetNumber, "");

            return replaceText;
        }

        public static List<ViewSheet> GetSheetsByNumber(Document curDoc, string sheetNumber)
        {
            List<ViewSheet> returnSheets = new List<ViewSheet>();

            //get all sheets
            List<ViewSheet> curSheets = GetAllSheets(curDoc);

            //loop through sheets and check sheet name
            foreach (ViewSheet curSheet in curSheets)
            {
                if (curSheet.SheetNumber.Contains(sheetNumber))
                {
                    returnSheets.Add(curSheet);
                }
            }

            return returnSheets;
        }

        public static List<ViewSheet> GetAllSheets(Document curDoc)
        {
            //get all sheets
            FilteredElementCollector m_colViews = new FilteredElementCollector(curDoc);
            m_colViews.OfCategory(BuiltInCategory.OST_Sheets);

            List<ViewSheet> m_sheets = new List<ViewSheet>();
            foreach (ViewSheet x in m_colViews.ToElements())
            {
                m_sheets.Add(x);
            }

            return m_sheets;
        }

        internal static AreaScheme GetAreaSchemeByName(Document doc, string schemeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(AreaScheme));

            foreach (AreaScheme areaScheme in collector)
            {
                if (areaScheme.Name == schemeName)
                {
                    return areaScheme;
                }
            }

            return null;
        }

        internal static List<Parameter> GetParametersByName(Document doc, List<string> paramNames)
        {
            List<Parameter> returnList = new List<Parameter>();

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Areas);

            foreach (string curName in paramNames)
            {
                Parameter curParam = collector.FirstElement().LookupParameter(curName);

                if (curParam != null)
                    returnList.Add(curParam);
            }

            return returnList;
        }

        internal static ViewSchedule CreateAreaSchedule(Document doc, string schedName, AreaScheme curAreaScheme)
        {
            ElementId catId = new ElementId(BuiltInCategory.OST_Areas);
            ViewSchedule newSchedule = ViewSchedule.CreateSchedule(doc, catId, curAreaScheme.Id);
            newSchedule.Name = schedName;

            return newSchedule;
        }
        internal static ViewSchedule CreateSchedule(Document doc, BuiltInCategory curCat, string name)
        {
            ElementId catId = new ElementId(curCat);
            ViewSchedule newSchedule = ViewSchedule.CreateSchedule(doc, catId);
            newSchedule.Name = name;

            return newSchedule;
        }
        internal static void AddFieldsToSchedule(Document doc, ViewSchedule newSched, List<Parameter> paramList)
        {
            foreach (Parameter curParam in paramList)
            {
                SchedulableField newField = new SchedulableField(ScheduleFieldType.Instance, curParam.Id);
                newSched.Definition.AddField(newField);
            }
        }

        internal static List<ViewPlan> GetAllAreaPlans(Document curDoc)
        {
            List<ViewPlan> returnList = new List<ViewPlan>();
            List<ViewPlan> viewList = GetAllViewPlans(curDoc);

            foreach (View x in viewList)
            {
                if (x.ViewType == ViewType.AreaPlan)
                {
                    returnList.Add((ViewPlan)x);
                }
            }

            return returnList;
        }

        public static List<ViewPlan> GetAllViewPlans(Document curDoc)
        {
            List<ViewPlan> returnList = new List<ViewPlan>();

            FilteredElementCollector viewCollector = new FilteredElementCollector(curDoc);
            viewCollector.OfCategory(BuiltInCategory.OST_Views);
            viewCollector.OfClass(typeof(ViewPlan)).ToElements();

            foreach (ViewPlan vp in viewCollector)
            {
                if (vp.IsTemplate == false)
                    returnList.Add(vp);
            }

            return returnList;
        }

        internal static ViewPlan GetAreaPlanByViewFamilyName(Document doc, string vftName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewPlan));

            foreach (ViewPlan curViewPlan in collector)
            {
                if (curViewPlan.ViewType == ViewType.AreaPlan)
                {
                    ViewFamilyType curVFT = doc.GetElement(curViewPlan.GetTypeId()) as ViewFamilyType;

                    if (curVFT.Name == vftName)
                        return curViewPlan;
                }
            }

            return null;
        }

        internal static List<ElementId> GetAllLevelIds(Document doc)
        {
            FilteredElementCollector colLevelId = new FilteredElementCollector(doc);
            colLevelId.OfClass(typeof(Level));

            List<Level> sortedList = colLevelId.Cast<Level>().ToList().OrderBy(x => x.Elevation).ToList();

            List<ElementId> returnList = new List<ElementId>();

            foreach (Level curLevel in sortedList)
            {
                if (curLevel.Name.Contains("Floor") || curLevel.Name.Contains("Level"))
                    returnList.Add(curLevel.Id);
            }

            return returnList;
        }

        internal static List<Level> GetLevelByNameContains(Document doc, string levelWord)
        {
            List<Level> levels = GetAllLevels(doc);

            List<Level> returnList = new List<Level>();

            foreach (Level curLevel in levels)
            {
                if (curLevel.Name.Contains(levelWord))
                    returnList.Add(curLevel);
            }

            return returnList;
        }

        public static List<Level> GetAllLevels(Document doc)
        {
            FilteredElementCollector colLevels = new FilteredElementCollector(doc);
            colLevels.OfCategory(BuiltInCategory.OST_Levels);

            List<Level> levels = new List<Level>();
            foreach (Element x in colLevels.ToElements())
            {
                if (x.GetType() == typeof(Level))
                {
                    levels.Add((Level)x);
                }
            }

            return levels;
            //order list by elevation
            //m_levels = (From l In m_levels Order By l.Elevation).tolist()
        }

        public static View GetViewTemplateByName(Document curDoc, string viewTemplateName)
        {
            List<View> viewTemplateList = GetAllViewTemplates(curDoc);

            foreach (View v in viewTemplateList)
            {
                if (v.Name == viewTemplateName)
                {
                    return v;
                }
            }

            return null;
        }

        public static List<View> GetAllViewTemplates(Document curDoc)
        {
            List<View> returnList = new List<View>();
            List<View> viewList = GetAllViews(curDoc);

            //loop through views and check if is view template
            foreach (View v in viewList)
            {
                if (v.IsTemplate == true)
                {
                    //add view template to list
                    returnList.Add(v);
                }
            }

            return returnList;
        }

        public static List<View> GetAllViews(Document curDoc)
        {
            FilteredElementCollector m_colviews = new FilteredElementCollector(curDoc);
            m_colviews.OfCategory(BuiltInCategory.OST_Views);

            List<View> m_views = new List<View>();
            foreach (View x in m_colviews.ToElements())
            {
                m_views.Add(x);
            }

            return m_views;
        }

        internal static ViewSchedule GetScheduleByNameContains(Document doc, string scheduleString)
        {
            List<ViewSchedule> m_scheduleList = GetAllSchedules(doc);           

            foreach (ViewSchedule curSchedule in m_scheduleList)
            {
                if (curSchedule.Name.Contains(scheduleString))
                    return curSchedule;
            }

            return null;
        }

        private static List<ViewSchedule> GetAllSchedules(Document doc)
        {
            {
                List<ViewSchedule> m_schedList = new List<ViewSchedule>();

                FilteredElementCollector curCollector = new FilteredElementCollector(doc);
                curCollector.OfClass(typeof(ViewSchedule));

                //loop through views and check if schedule - if so then put into schedule list
                foreach (ViewSchedule curView in curCollector)
                {
                    if (curView.ViewType == ViewType.Schedule)
                    {
                        m_schedList.Add((ViewSchedule)curView);
                    }
                }

                return m_schedList;
            }
        }

        internal static List<ViewSchedule> GetAllScheduleByNameContains(Document doc, string schedName)
        {
            List<ViewSchedule> m_scheduleList = GetAllSchedules(doc);

            List<ViewSchedule> m_returnList = new List<ViewSchedule>();

            foreach (ViewSchedule curSchedule in m_scheduleList)
            {
                if (curSchedule.Name.Contains(schedName))
                    m_returnList.Add(curSchedule);
            }

            return m_returnList;
        }

        internal static ScheduleField FindScheduleField(ViewSchedule newFloorSched, Parameter paramName)
        {
            ScheduleDefinition definition = newFloorSched.Definition;
            ScheduleField foundField = null;
            ElementId paramId = new ElementId(paramName);

            foreach (ScheduleFieldId fieldId in definition.GetFieldOrder())
            {
                foundField = definition.GetField(fieldId);
                if (foundField.ParameterId == paramId)
                {
                    return foundField;
                }
            }

            return null;
        }

        #region Design options
        internal static List<DesignOption> getAllDesignOptions(Document curDoc)
        {
            FilteredElementCollector curCol = new FilteredElementCollector(curDoc);
            curCol.OfCategory(BuiltInCategory.OST_DesignOptions);

            List<DesignOption> doList = new List<DesignOption>();
            foreach (DesignOption curOpt in curCol)
            {
                doList.Add(curOpt);
            }

            return doList;
        }

        internal static DesignOption getDesignOptionByName(Document curDoc, string designOpt)
        {
            //get all design options
            List<DesignOption> doList = getAllDesignOptions(curDoc);

            foreach (DesignOption curOpt in doList)
            {
                if (curOpt.Name == designOpt)
                {
                    return curOpt;
                }
            }

            return null;
        }
        #endregion
    }
}