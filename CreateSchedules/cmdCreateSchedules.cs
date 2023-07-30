#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.ExceptionServices;
using Forms = System.Windows;
using System.Reflection;
using System.Windows.Controls;
#endregion

namespace CreateSchedules
{
    [Transaction(TransactionMode.Manual)]
    public class cmdCreateSchedules : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document curDoc = uidoc.Document;

            frmCreateSchedules curForm = new frmCreateSchedules()
            {
                Width = 420,
                Height = 400,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
            };

            curForm.ShowDialog();

            // get data from the form

            Globals.ElevDesignation = curForm.GetComboBoxElevationSelectedItem();
            int floorNum = curForm.GetComboBoxFloorsSelectedItem();
            string typeFoundation = curForm.GetGroup1();
            string typeAttic = curForm.GetGroup2();


            bool chbIndexResult = curForm.GetCheckboxIndex();
            bool chbVeneerResult = curForm.GetCheckboxVeneer();
            bool chbFloorResult = curForm.GetCheckboxFloor();
            bool chbFrameResult = curForm.GetCheckboxFrame();
            bool chbAtticResult = curForm.GetCheckboxAttic();

            // start the transaction

            using (Transaction t = new Transaction(curDoc))
            {
                t.Start("Create Schedules");

                #region Sheet Index

                // check to see if the sheet index exists

                ViewSchedule schedIndex = Utils.GetScheduleByNameContains(curDoc, "Sheet Index - Elevation " + Globals.ElevDesignation);

                if (chbIndexResult == true)
                {
                    if (schedIndex == null)
                    {
                        // duplicate the first schedule found with Sheet Index in the name
                        List<ViewSchedule> listSched = Utils.GetAllScheduleByNameContains(curDoc, "Sheet Index");

                        ViewSchedule dupSched = listSched.FirstOrDefault();

                        Element viewSched = curDoc.GetElement(dupSched.Duplicate(ViewDuplicateOption.Duplicate));

                        // rename the duplicated schedule to the new elevation

                        string originalName = viewSched.Name;
                        string[] schedTitle = originalName.Split('C');

                        string curTitle = schedTitle[0];

                        string lastChar = curTitle.Substring(curTitle.Length - 2);
                        string newLast = Globals.ElevDesignation.ToString();

                        viewSched.Name = curTitle.Replace(lastChar, newLast);

                        // set the design option to the specified elevation designation

                        DesignOption curOption = Utils.getDesignOptionByName(curDoc, "Elevation : " + lastChar); // !!! this code throws an error
                    }
                }

                #endregion

                #region Exterior Veneer Calculations

                ViewSchedule veneerIndex = Utils.GetScheduleByNameContains(curDoc, "Exterior Veneer Calculations - Elevation " + Globals.ElevDesignation);

                if (chbVeneerResult == true)
                {
                    if (veneerIndex == null)
                    {
                        // duplicate the first schedule with "Exterior Venner Calculations" in the name
                        List<ViewSchedule> listSched = Utils.GetAllScheduleByNameContains(curDoc, "Exterior Veneer Calculations");

                        ViewSchedule dupSched = listSched.FirstOrDefault();

                        Element viewSched = curDoc.GetElement(dupSched.Duplicate(ViewDuplicateOption.Duplicate));

                        // rename the duplicated schedule to the new elevation

                        string originalName = viewSched.Name;
                        string[] schedTitle = originalName.Split('-');                        

                        viewSched.Name = schedTitle[0] + "- Elevation " + Globals.ElevDesignation;

                        // set the design option to the specified elevation designation

                        DesignOption curOption = Utils.getDesignOptionByName(curDoc, "Elevation : " + Globals.ElevDesignation); // !!! this code throws an error
                    }
                }

                #endregion

                #region Floor Area Schedules

                // check to see if the floor area scheme exists

                if (chbFloorResult == true)
                {
                    // set the variable for the floor Area Scheme name

                    AreaScheme schemeFloor = Utils.GetAreaSchemeByName(curDoc, Globals.ElevDesignation + " Floor");

                    if (schemeFloor == null)
                    {
                        // if null, warn the user & exit the command

                        Forms.MessageBox.Show("The Area Scheme does not exist or is named incorrectly. Resolve the issue & try again");
                        return Result.Failed;
                    }

                    // if the floor area scheme exists, check to see if the floor area plans exist

                    if (schemeFloor != null)
                    {
                        // set some variables

                        ViewPlan areaFloorView = Utils.GetAreaPlanByViewFamilyName(curDoc, Globals.ElevDesignation + " Floor");

                        // if not, create the area plans

                        if (areaFloorView == null)
                        {
                            string levelWord = "";

                            if (typeFoundation == "Slab")
                            {
                                levelWord = "Floor";
                            }
                            else if (typeFoundation == "Basement" || typeFoundation == "Crawlspace")
                            {
                                levelWord = "Level";
                            }

                            List<Level> levelList = Utils.GetLevelByNameContains(curDoc, levelWord);

                            ElementId schemeFloorId = schemeFloor.Id;

                            List<View> areaViews = new List<View>();

                            int countFloor = 1;

                            foreach (Level curlevel in levelList)
                            {
                                ElementId curLevelId = curlevel.Id;

                                View vtFloorAreas = Utils.GetViewTemplateByName(curDoc, "10-Floor Area");

                                ViewPlan areaFloor = ViewPlan.CreateAreaPlan(curDoc, schemeFloorId, curLevelId);
                                areaFloor.Name = "Floor_" + countFloor.ToString();
                                areaFloor.ViewTemplateId = vtFloorAreas.Id;

                                areaViews.Add(areaFloor);

                                countFloor++;

                                // loop through each newly created area plan

                                foreach (View curView in areaViews)
                                {
                                    uidoc.ActiveView = curView;

                                    //set the color scheme

                                    Parameter level = curView.LookupParameter("Associated Level");

                                    if (level.ToString() == "Lower Level")
                                    {
                                        // add these areas
                                    }
                                    else if (level.ToString() == "Main Level" || level.ToString() == "First Floor")
                                    {
                                        // add these areas, can this be done as a data class?

                                        // need to set a range and put this in a loop that increments the insPoint

                                        XYZ insStart = new XYZ(0, 0, 0);

                                        double calcOffset = 1.0 * curDoc.ActiveView.Scale;

                                        XYZ offset = new XYZ(0, calcOffset, 0);

                                        UV insPoint = new UV(insStart.X, insStart.Y);

                                        Area areaLiving1 = curDoc.Create.NewArea(areaFloor, insPoint);
                                        areaLiving1.Number = "1";
                                        areaLiving1.Name = "Living";                                    

                                        Area areaGarage = curDoc.Create.NewArea(areaFloor, insPoint);
                                        areaGarage.Number = "2";
                                        areaGarage.Name = "Garage";

                                        Area areaCoveredPatio = curDoc.Create.NewArea(areaFloor, insPoint);
                                        areaCoveredPatio.Number = "3";
                                        areaCoveredPatio.Name = "Covered Patio";

                                        Area areaCoveredPorch = curDoc.Create.NewArea(areaFloor, insPoint);
                                        areaCoveredPorch.Number = "4";
                                        areaCoveredPorch.Name = "Covered Porch";

                                        Area areaPorteCochere = curDoc.Create.NewArea(areaFloor, insPoint);
                                        areaPorteCochere.Number = "5";
                                        areaPorteCochere.Name = "Porte Cochere";

                                        Area areaPatio = curDoc.Create.NewArea(areaFloor, insPoint);
                                        areaPatio.Number = "6";
                                        areaPatio.Name = "Patio";

                                        Area areaPorch = curDoc.Create.NewArea(areaFloor, insPoint);
                                        areaPorch.Number = "7";
                                        areaPorch.Name = "Porch";

                                        Area areaOption1 = curDoc.Create.NewArea(areaFloor, insPoint);
                                        areaOption1.Number = "8";
                                        areaOption1.Name = "Option";

                                        // !!! still need to set values for Area Category & Comments
                                    }
                                    else if (level.ToString() == "Upper Level" || level.ToString() == "Second Floor")
                                    {
                                        // add these areas
                                    }
                                }
                            }
                        }

                        // if the floor area plans exist, create the schedule

                        // get the Area category Id
                        ElementId areaCatId = new ElementId(BuiltInCategory.OST_Areas); // ??? is this needed

                        // get the area scheme for the schedule
                        AreaScheme curAreaScheme = Utils.GetAreaSchemeByName(curDoc, Globals.ElevDesignation + " Floor");

                        // create the new schedule
                        ViewSchedule newFloorSched = Utils.CreateAreaSchedule(curDoc, "Floor Areas - Elevation " + Globals.ElevDesignation, curAreaScheme);

                        if (areaFloorView != null)
                        {
                            if (floorNum == 1)
                            {
                                // create a list of the fields for the schedule
                                List<string> paramNames = new List<string>() { "Area Category", "Comments", "Name", "Area", "Number" };

                                // get the associated parameters & add them to the schedule
                                List<Parameter> paramsFloorSingle = Utils.GetParametersByName(curDoc, paramNames);                                
                                Utils.AddFieldsToSchedule(curDoc, newFloorSched, paramsFloorSingle);

                                // create the fields to use for filter and formatting

                                //// get element Id of the parameters
                                ElementId catFieldId = Utils.GetElementIdFromSharedParameter(curDoc, "Area Category");
                                ElementId comFieldId = Utils.GetElementIdFromSharedParameter(curDoc, "Comments");
                                ElementId nameFieldId = Utils.GetElementIdFromSharedParameter(curDoc, "Name");
                                ElementId areaFieldId = Utils.GetElementIdFromSharedParameter(curDoc, "Area");
                                ElementId numFieldId = Utils.GetElementIdFromSharedParameter(curDoc, "Number");

                                ScheduleField catField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, catFieldId);
                                catField.IsHidden = true;

                                ScheduleField comField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, comFieldId);
                                comField.IsHidden = true;

                                ScheduleField nameField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, nameFieldId);
                                nameField.IsHidden = false;
                                nameField.ColumnHeading = "Name";
                                nameField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                                nameField.HorizontalAlignment = ScheduleHorizontalAlignment.Left;

                                ScheduleField areaField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, areaFieldId);
                                areaField.IsHidden = false;
                                areaField.ColumnHeading = "Area";
                                areaField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                                areaField.HorizontalAlignment = ScheduleHorizontalAlignment.Right;
                                // areaField.IsCalculatedField = true;  // ??? can this be set to "Calculate totals"

                                ScheduleField numField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, numFieldId);
                                numField.IsHidden = true;

                                // create the filters
                                ScheduleFilter catFilter = new ScheduleFilter(catField.FieldId, ScheduleFilterType.Contains, "Options");
                                newFloorSched.Definition.AddFilter(catFilter);

                                ScheduleFilter areaFilter = new ScheduleFilter(areaField.FieldId, ScheduleFilterType.GreaterThan, "0 SF");
                                newFloorSched.Definition.AddFilter(areaFilter);

                                // set the sorting
                                ScheduleSortGroupField catSort = new ScheduleSortGroupField(catField.FieldId, ScheduleSortOrder.Ascending);
                                catSort.ShowFooter = true;
                                catSort.ShowFooterCount = true; //??? how to set the footer to be "Title and Totals"
                                catSort.ShowBlankLine = true;

                                ScheduleSortGroupField comSort = new ScheduleSortGroupField(comField.FieldId, ScheduleSortOrder.Ascending);
                            }

                            else if (floorNum == 2 || floorNum == 3)
                            {
                                List<string> paramNames = new List<string>() { "Area Category", "Comments", "Level", "Name", "Area", "Number" };
                                List<Parameter> paramsFloorMulti = Utils.GetParametersByName(curDoc, paramNames);

                                Utils.AddFieldsToSchedule(curDoc, newFloorSched, paramsFloorMulti);

                                ScheduleField catField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, paramsFloorMulti[0].Id);
                                ScheduleField commentField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, paramsFloorMulti[1].Id);
                                ScheduleField levelField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, paramsFloorMulti[1].Id);
                                ScheduleField nameField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, paramsFloorMulti[3].Id);
                                ScheduleField areaField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, paramsFloorMulti[4].Id);
                                ScheduleField numField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, paramsFloorMulti[5].Id);

                                ScheduleFilter catFilter = new ScheduleFilter(catField.FieldId, ScheduleFilterType.Contains, "Options");
                                ScheduleFilter areaFilter = new ScheduleFilter(areaField.FieldId, ScheduleFilterType.GreaterThan, "0 SF");

                                ScheduleSortGroupField catSort = new ScheduleSortGroupField(catField.FieldId, ScheduleSortOrder.Ascending);
                                ScheduleSortGroupField commentSort = new ScheduleSortGroupField(commentField.FieldId, ScheduleSortOrder.Ascending);
                                ScheduleSortGroupField nameSort = new ScheduleSortGroupField(nameField.FieldId, ScheduleSortOrder.Ascending);
                                ScheduleSortGroupField levelSort = new ScheduleSortGroupField(levelField.FieldId, ScheduleSortOrder.Ascending);
                            }
                        }
                    }
                }

                #endregion

                #region Frame Area Schedules

                // check to see if the frame area scheme exists

                if (chbFrameResult == true)
                {
                    // set the variable for the frame Area Scheme name

                    AreaScheme schemeFrame = Utils.GetAreaSchemeByName(curDoc, Globals.ElevDesignation + " Frame");

                    if (schemeFrame == null)
                    {
                        // if null, warn the user & exit the command

                        Forms.MessageBox.Show("The Area Scheme does not exist or is named incorrectly. Resolve the issue & try again.");
                        return Result.Failed;
                    }

                    if (schemeFrame != null)
                    {
                        // set some variables

                        ViewPlan areaFrameView = Utils.GetAreaPlanByViewFamilyName(curDoc, Globals.ElevDesignation + " Frame");

                        // if not, create the area plans

                        if (areaFrameView == null)
                        {
                            string levelWord = "";

                            if (typeFoundation == "Slab")
                            {
                                levelWord = "Floor";
                            }
                            else if (typeFoundation == "Basement" || typeFoundation == "Crawlspace")
                            {
                                levelWord = "Level";
                            }

                            List<Level> levelList = Utils.GetLevelByNameContains(curDoc, levelWord);

                            ElementId schemeFrameId = schemeFrame.Id;

                            List<View> frameViews = new List<View>();

                            int countFloor = 1;

                            foreach (Level curlevel in levelList)
                            {
                                ElementId curLevelId = curlevel.Id;

                                View vtFrameAreas = Utils.GetViewTemplateByName(curDoc, "11-Frame Area");

                                ViewPlan areaFrame = ViewPlan.CreateAreaPlan(curDoc, schemeFrameId, curLevelId);
                                areaFrame.Name = "Frame_" + countFloor.ToString();
                                areaFrame.ViewTemplateId = vtFrameAreas.Id;

                                frameViews.Add(areaFrame);

                                countFloor++;

                                // loop through each newly created area plan

                                foreach (View curView in frameViews)
                                {
                                    uidoc.ActiveView = curView;

                                    //set the color scheme

                                    Parameter level = curView.LookupParameter("Associated Level");

                                    if (level.ToString() == "Lower Level")
                                    {
                                        // add these areas, can this be done as a data class?

                                        // need to set a range and put this in a loop that increments the insPoint

                                        XYZ insStart = new XYZ(0, 0, 0);

                                        double calcOffset = 1.0 * curDoc.ActiveView.Scale;

                                        XYZ offset = new XYZ(0, calcOffset, 0);

                                        UV insPoint = new UV(insStart.X, insStart.Y);

                                        Area areaStandard1 = curDoc.Create.NewArea(areaFrame, insPoint);
                                        areaStandard1.Number = "1";
                                        areaStandard1.Name = "Standard";

                                        Area areaOption1 = curDoc.Create.NewArea(areaFrame, insPoint);
                                        areaOption1.Number = "2";
                                        areaOption1.Name = "Option";

                                        // !!! still need to set values for Area Category & Comments
                                    }
                                    else if (level.ToString() == "Main Level" || level.ToString() == "First Floor")
                                    {
                                        // add these areas, can this be done as a data class?

                                        // need to set a range and put this in a loop that increments the insPoint

                                        XYZ insStart = new XYZ(0, 0, 0);

                                        double calcOffset = 1.0 * curDoc.ActiveView.Scale;

                                        XYZ offset = new XYZ(0, calcOffset, 0);

                                        UV insPoint = new UV(insStart.X, insStart.Y);

                                        Area areaStandard2 = curDoc.Create.NewArea(areaFrame, insPoint);
                                        areaStandard2.Number = "3";
                                        areaStandard2.Name = "Standard";                                        

                                        Area areaOption2 = curDoc.Create.NewArea(areaFrame, insPoint);
                                        areaOption2.Number = "4";
                                        areaOption2.Name = "Option";

                                        // !!! still need to set values for Area Category & Comments
                                    }
                                    else if (level.ToString() == "Upper Level" || level.ToString() == "Second Floor")
                                    {
                                        // add these areas, can this be done as a data class?

                                        // need to set a range and put this in a loop that increments the insPoint

                                        XYZ insStart = new XYZ(0, 0, 0);

                                        double calcOffset = 1.0 * curDoc.ActiveView.Scale;

                                        XYZ offset = new XYZ(0, calcOffset, 0);

                                        UV insPoint = new UV(insStart.X, insStart.Y);

                                        Area areaStandard3 = curDoc.Create.NewArea(areaFrame, insPoint);
                                        areaStandard3.Number = "5";
                                        areaStandard3.Name = "Standard";

                                        Area areaOption3 = curDoc.Create.NewArea(areaFrame, insPoint);
                                        areaOption3.Number = "6";
                                        areaOption3.Name = "Option";
                                    }
                                }
                            }
                        }

                        // if the frame area plans exist, create the schedule

                        // get the Area category Id
                        ElementId areaCatId = new ElementId(BuiltInCategory.OST_Areas); // ??? is this needed

                        // get the area scheme for the schedule
                        AreaScheme curAreaScheme = Utils.GetAreaSchemeByName(curDoc, Globals.ElevDesignation + " Frame");

                        // create the new schedule
                        ViewSchedule newFrameSched = Utils.CreateAreaSchedule(curDoc, "Frame Areas - Elevation " + Globals.ElevDesignation, curAreaScheme);

                        if (areaFrameView != null)
                        {
                            if (floorNum == 1)
                            {
                                // create a list of the fields for the schedule
                                List<string> paramNames = new List<string>() { "Area Category", "Name", "Area" };

                                // get the associated parameters & add them to the schedule
                                List<Parameter> paramsFloorSingle = Utils.GetParametersByName(curDoc, paramNames);
                                Utils.AddFieldsToSchedule(curDoc, newFrameSched, paramsFloorSingle);

                                // create the fields to use for filter and formatting

                                // get element Id of the parameters
                                ElementId catFieldId = Utils.GetElementIdFromSharedParameter(curDoc, "Area Category");
                                ElementId nameFieldId = Utils.GetElementIdFromSharedParameter(curDoc, "Name");                                
                                ElementId areaFieldId = Utils.GetElementIdFromSharedParameter(curDoc, "Area");

                                ScheduleField catField = newFrameSched.Definition.AddField(ScheduleFieldType.Instance, catFieldId);
                                catField.IsHidden = true;

                                ScheduleField nameField = newFrameSched.Definition.AddField(ScheduleFieldType.Instance, nameFieldId);
                                nameField.IsHidden = false;                               
                                nameField.ColumnHeading = "Name";
                                nameField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                                nameField.HorizontalAlignment = ScheduleHorizontalAlignment.Left;

                                ScheduleField areaField = newFrameSched.Definition.AddField(ScheduleFieldType.Instance, areaFieldId);
                                areaField.IsHidden = false;
                                areaField.ColumnHeading = "Area";
                                areaField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                                areaField.HorizontalAlignment = ScheduleHorizontalAlignment.Right;
                                // areaField.IsCalculatedField = true;  // ??? can this be set to "Calculate totals"

                                // create the filters

                                ScheduleFilter catFilter = new ScheduleFilter(catField.FieldId, ScheduleFilterType.Contains, "Options");
                                newFrameSched.Definition.AddFilter(catFilter);

                                ScheduleFilter areaFilter = new ScheduleFilter(areaField.FieldId, ScheduleFilterType.GreaterThan, "0 SF");

                                // set the sorting

                                ScheduleSortGroupField catSort = new ScheduleSortGroupField(catField.FieldId, ScheduleSortOrder.Ascending);
                                catSort.ShowHeader = true;                                 
                                catSort.ShowBlankLine = true;
                            }

                            else if (floorNum == 2 || floorNum == 3)
                            {
                                // create a list of the fields for the schedule
                                List<string> paramNames = new List<string>() { "Area Category", "Name", "Level", "Area" };

                                // get the associated parameters & add them to the schedule
                                List<Parameter> paramsFloorSingle = Utils.GetParametersByName(curDoc, paramNames);
                                Utils.AddFieldsToSchedule(curDoc, newFrameSched, paramsFloorSingle);

                                // create the fields to use for filter and formatting

                                // get element Id of the parameters
                                ElementId catFieldId = Utils.GetElementIdFromSharedParameter(curDoc, "Area Category");
                                ElementId nameFieldId = Utils.GetElementIdFromSharedParameter(curDoc, "Name");
                                ElementId levelFieldId = Utils.GetElementIdFromSharedParameter(curDoc, "Level");
                                ElementId areaFieldId = Utils.GetElementIdFromSharedParameter(curDoc, "Area");


                                ScheduleField catField = newFrameSched.Definition.AddField(ScheduleFieldType.Instance, catFieldId);
                                catField.IsHidden = true;

                                ScheduleField nameField = newFrameSched.Definition.AddField(ScheduleFieldType.Instance, nameFieldId);
                                nameField.IsHidden = true;

                                ScheduleField levelField = newFrameSched.Definition.AddField(ScheduleFieldType.Instance, levelFieldId);
                                levelField.IsHidden = false;
                                levelField.ColumnHeading = "Level";
                                levelField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                                levelField.HorizontalAlignment = ScheduleHorizontalAlignment.Left;

                                ScheduleField areaField = newFrameSched.Definition.AddField(ScheduleFieldType.Instance, areaFieldId);
                                areaField.IsHidden = false;
                                areaField.ColumnHeading = "Area";
                                areaField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                                areaField.HorizontalAlignment = ScheduleHorizontalAlignment.Right;
                                // areaField.IsCalculatedField = true;  // ??? can this be set to "Calculate totals"

                                // create the filters

                                ScheduleFilter catFilter = new ScheduleFilter(catField.FieldId, ScheduleFilterType.Contains, "Options");
                                newFrameSched.Definition.AddFilter(catFilter);

                                ScheduleFilter areaFilter = new ScheduleFilter(areaField.FieldId, ScheduleFilterType.GreaterThan, "0 SF");

                                // set the sorting

                                ScheduleSortGroupField catSort = new ScheduleSortGroupField(catField.FieldId, ScheduleSortOrder.Ascending);
                                catSort.ShowHeader = true;
                                catSort.ShowBlankLine = true;

                                ScheduleSortGroupField nameSort = new ScheduleSortGroupField(nameField.FieldId, ScheduleSortOrder.Ascending);
                                nameSort.ShowHeader = true;
                                nameSort.ShowFooter = true;

                                ScheduleSortGroupField levelSort = new ScheduleSortGroupField(levelField.FieldId, ScheduleSortOrder.Ascending);
                            }
                        }
                    }
                }

                #endregion

                #region Roof Ventilation Schedules

                // chck to see if the attic area scheme exists

                if (chbAtticResult == true)
                {
                    // set the variable for the frame Area Scheme name

                    AreaScheme schemeAttic = Utils.GetAreaSchemeByName(curDoc, Globals.ElevDesignation + " Roof Ventilation");

                    if (schemeAttic == null)
                    {
                        // if null, warn the user & exit the command

                        Forms.MessageBox.Show("The Area Scheme does not exist or is named incorrectly. Resolve the issue & try again.");
                        return Result.Failed;
                    }

                    if (schemeAttic != null)
                    {
                        // check to see if the attic area plans exist
                    }
                }

                #endregion

                t.Commit();
            }

            return Result.Succeeded;
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}