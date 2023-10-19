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

            // set some variables for paramter values
            string newFilter = "";

            if (Globals.ElevDesignation == "A")
                newFilter = "1";
            else if (Globals.ElevDesignation == "B")
                newFilter = "2";
            else if (Globals.ElevDesignation == "C")
                newFilter = "3";
            else if (Globals.ElevDesignation == "D")
                newFilter = "4";
            else if (Globals.ElevDesignation == "S")
                newFilter = "5";
            else if (Globals.ElevDesignation == "T")
                newFilter = "6";

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

                if (chbIndexResult == true && schedIndex == null)
                {
                    Utils.DuplicateAndRenameSheetIndex(curDoc, newFilter);
                }                

                #endregion

                #region Exterior Veneer Calculations

                ViewSchedule veneerIndex = Utils.GetScheduleByNameContains(curDoc, "Exterior Veneer Calculations - Elevation " + Globals.ElevDesignation);

                if (chbVeneerResult == true && veneerIndex == null)
                {
                    Utils.DuplicateAndConfigureVeneerSchedule(curDoc);
                }

                #endregion

                #region Floor Areas                

                // check to see if the floor area scheme exists

                if (chbFloorResult == true)
                {
                    // set the variable for the floor Area Scheme name
                    AreaScheme schemeFloor = Utils.GetAreaSchemeByName(curDoc, Globals.ElevDesignation + " Floor");

                    // set the variable for the color fill scheme
                    ColorFillScheme schemeColorFill = Utils.GetColorFillSchemeByName(curDoc, "Floor", schemeFloor);

                    if (schemeFloor == null || schemeColorFill == null)
                    {
                        // if null, warn the user & exit the command
                        TaskDialog tdSchemeError = new TaskDialog("Error");
                        tdSchemeError.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
                        tdSchemeError.Title = "Create Schedules";
                        tdSchemeError.TitleAutoPrefix = false;
                        tdSchemeError.MainContent = "Either the Area Scheme, or the Color Scheme, does not exist " +
                            "or is named incorrectly. Resolve the issue & try again.";
                        tdSchemeError.CommonButtons = TaskDialogCommonButtons.Close;

                        TaskDialogResult tdSchemeErrorRes = tdSchemeError.Show();

                        return Result.Failed;
                    }

                    #region Floor Area Plans

                    // if the floor area scheme exists, check to see if the floor area plans exist

                    if (schemeFloor != null)
                    {
                        // check if area plans exist
                        ViewPlan areaFloorView = Utils.GetAreaPlanByViewFamilyName(curDoc, Globals.ElevDesignation + " Floor");                        

                        // if not, create the area plans

                        if (areaFloorView == null)
                        {
                            string levelWord = "";

                           if (typeFoundation == "Basement" || typeFoundation == "Crawlspace")
                            {
                                levelWord = "Level";
                            }
                           else
                            {
                                levelWord = "Floor";
                            }

                           List<Level> levelList = Utils.GetLevelByNameContains(curDoc, levelWord);

                           List<View> areaViews = new List<View>();

                            foreach (Level curlevel in levelList)
                            {
                                // get the category & set category Id
                                Category areaCat = curDoc.Settings.Categories.get_Item(BuiltInCategory.OST_Areas);

                                // get the element Id of the current level
                                ElementId curLevelId = curlevel.Id;

                                // create & set variable for the view template
                                View vtFloorAreas = Utils.GetViewTemplateByName(curDoc, "10-Floor Area");

                                // create the area plan
                                ViewPlan areaFloor = ViewPlan.CreateAreaPlan(curDoc, schemeFloor.Id, curLevelId);
                                areaFloor.ViewTemplateId = vtFloorAreas.Id;
                                areaFloor.SetColorFillSchemeId(areaCat.Id, schemeColorFill.Id);

                                areaViews.Add(areaFloor);
                            }

                            foreach (ViewPlan curView in areaViews)
                            {
                                //set the color scheme                                   

                                // create insertion points
                                XYZ insStart = new XYZ(50, 0, 0);

                                UV insPoint = new UV(insStart.X, insStart.Y);
                                UV offset = new UV(0, 8);

                                XYZ tagInsert = new XYZ(50, 0, 0);
                                XYZ tagOffset = new XYZ(0, 8, 0);

                                if (curView.Name == "Lower Level")
                                {
                                    // add these areas
                                    List<clsAreaData> areasLower = new List<clsAreaData>()
                                        {
                                            new clsAreaData("13", "Living", "Total Covered", "A"),
                                            new clsAreaData("14", "Mechanical", "Total Covered", "K"),
                                            new clsAreaData("15", "Unfinished Basement", "Total Covered", "L"),
                                            new clsAreaData("16", "Option", "Options", "H")
                                        };
                                    foreach (var areaInfo in areasLower)
                                    {
                                        Utils.CreateFloorAreaWithTag(curDoc, curView, ref insPoint, ref tagInsert, areaInfo);
                                    }
                                }
                                else if (curView.Name == "Main Level" || curView.Name == "First Floor")
                                {
                                    // add these areas
                                    List<clsAreaData> areasMain = new List<clsAreaData>
                                        {
                                            new clsAreaData("1", "Living", "Total Covered", "A"),
                                            new clsAreaData("2", "Garage", "Total Covered", "B"),
                                            new clsAreaData("3", "Covered Patio", "Total Covered", "C"),
                                            new clsAreaData("4", "Covered Porch", "Total Covered", "D"),
                                            new clsAreaData("5", "Porte Cochere", "Total Covered", "E"),
                                            new clsAreaData("6", "Patio", "Total Uncovered", "F"),
                                            new clsAreaData("7", "Porch", "Total Uncovered", "G"),
                                            new clsAreaData("8", "Option", "Options", "H")
                                        };

                                    foreach (var areaInfo in areasMain)
                                    {
                                        Utils.CreateFloorAreaWithTag(curDoc, curView, ref insPoint, ref tagInsert, areaInfo);
                                    }
                                }
                                else
                                {
                                    // add these areas
                                    List<clsAreaData> areasUpper = new List<clsAreaData>
                                        {
                                            new clsAreaData("9", "Living", "Total Covered", "A"),
                                            new clsAreaData("10", "Covered Balcony", "Total Covered", "I"),
                                            new clsAreaData("11", "Balcony", "Total Uncovered", "J"),
                                            new clsAreaData("12", "Option", "Options", "H")
                                        };

                                    foreach (var areaInfo in areasUpper)
                                    {
                                        Utils.CreateFloorAreaWithTag(curDoc, curView, ref insPoint, ref tagInsert, areaInfo);
                                    }
                                }
                            }
                        }

                        #endregion

                        #region Floor Area Schedule

                        // if the floor area plans exist, create the schedule                        

                        // get the area scheme for the schedule
                        AreaScheme curAreaScheme = Utils.GetAreaSchemeByName(curDoc, Globals.ElevDesignation + " Floor");

                        // create the new schedule
                        ViewSchedule newFloorSched = Utils.CreateAreaSchedule(curDoc, "Floor Areas - Elevation " + Globals.ElevDesignation, curAreaScheme);

                        if (areaFloorView != null)
                        {
                            if (floorNum == 2 || floorNum == 3)
                            {
                                // get element Id of the fields to be used in the schedule
                                ElementId catFieldId = Utils.GetProjectParameterId(curDoc, "Area Category");
                                ElementId comFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Rooms, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                                ElementId levelFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Rooms, BuiltInParameter.LEVEL_NAME);
                                ElementId nameFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Rooms, BuiltInParameter.ROOM_NAME);
                                ElementId areaFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Rooms, BuiltInParameter.ROOM_AREA);
                                ElementId numFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Rooms, BuiltInParameter.ROOM_NUMBER);

                                // create the fields and assign formatting properties
                                ScheduleField catField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, catFieldId);
                                catField.IsHidden = true;

                                ScheduleField comField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, comFieldId);
                                comField.IsHidden = true;

                                ScheduleField levelField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, levelFieldId);
                                levelField.IsHidden = false;
                                levelField.ColumnHeading = "Level";
                                levelField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                                levelField.HorizontalAlignment = ScheduleHorizontalAlignment.Left;

                                ScheduleField nameField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, nameFieldId);
                                nameField.IsHidden = true;

                                ScheduleField areaField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, areaFieldId);
                                areaField.IsHidden = false;
                                areaField.ColumnHeading = "Area";
                                areaField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                                areaField.HorizontalAlignment = ScheduleHorizontalAlignment.Right;
                                areaField.DisplayType = ScheduleFieldDisplayType.Totals;

                                ScheduleField numField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, numFieldId);
                                numField.IsHidden = true;

                                // create the filters
                                ScheduleFilter catFilter = new ScheduleFilter(catField.FieldId, ScheduleFilterType.NotContains, "Options");
                                newFloorSched.Definition.AddFilter(catFilter);

                                ScheduleFilter areaFilter = new ScheduleFilter(areaField.FieldId, ScheduleFilterType.HasValue);
                                newFloorSched.Definition.AddFilter(areaFilter);

                                // set the sorting
                                ScheduleSortGroupField catSort = new ScheduleSortGroupField(catField.FieldId, ScheduleSortOrder.Ascending);
                                catSort.ShowFooter = true;
                                catSort.ShowFooterTitle = true;
                                catSort.ShowFooterCount = true;
                                catSort.ShowBlankLine = true;                           
                                newFloorSched.Definition.AddSortGroupField(catSort);

                                ScheduleSortGroupField comSort = new ScheduleSortGroupField(comField.FieldId, ScheduleSortOrder.Ascending);
                                newFloorSched.Definition.AddSortGroupField(comSort);

                                ScheduleSortGroupField nameSort = new ScheduleSortGroupField(nameField.FieldId, ScheduleSortOrder.Ascending);
                                nameSort.ShowHeader = true;
                                nameSort.ShowFooter = true;
                                newFloorSched.Definition.AddSortGroupField(nameSort);

                                ScheduleSortGroupField levelSort = new ScheduleSortGroupField(levelField.FieldId, ScheduleSortOrder.Ascending);
                                newFloorSched.Definition.AddSortGroupField(levelSort);
                            }
                            else
                            {
                                // get element Id of the fields to be used in the schedule
                                ElementId catFieldId = Utils.GetProjectParameterId(curDoc, "Area Category");
                                ElementId comFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Rooms, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                                ElementId nameFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Rooms, BuiltInParameter.ROOM_NAME);
                                ElementId areaFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Rooms, BuiltInParameter.ROOM_AREA);
                                ElementId numFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Rooms, BuiltInParameter.ROOM_NUMBER);

                                // create the fields and assign formatting properties
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
                                areaField.DisplayType = ScheduleFieldDisplayType.Totals;

                                ScheduleField numField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, numFieldId);
                                numField.IsHidden = true;

                                // create the filters
                                ScheduleFilter catFilter = new ScheduleFilter(catField.FieldId, ScheduleFilterType.NotContains, "Options");
                                newFloorSched.Definition.AddFilter(catFilter);

                                ScheduleFilter areaFilter = new ScheduleFilter(areaField.FieldId, ScheduleFilterType.HasValue);
                                newFloorSched.Definition.AddFilter(areaFilter);

                                // set the sorting
                                ScheduleSortGroupField catSort = new ScheduleSortGroupField(catField.FieldId, ScheduleSortOrder.Ascending);
                                catSort.ShowFooter = true;
                                catSort.ShowFooterTitle = true;
                                catSort.ShowFooterCount = true;
                                catSort.ShowBlankLine = true;
                                newFloorSched.Definition.AddSortGroupField(catSort);

                                ScheduleSortGroupField comSort = new ScheduleSortGroupField(comField.FieldId, ScheduleSortOrder.Ascending);
                                newFloorSched.Definition.AddSortGroupField(comSort);

                                newFloorSched.Definition.IsItemized = false;
                            }
                        }

                        #endregion
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
                                List<Parameter> paramsFloorSingle = Utils.GetParametersByName(curDoc, paramNames, BuiltInCategory.OST_Areas);
                                Utils.AddFieldsToSchedule(curDoc, newFrameSched, paramsFloorSingle);

                                // create the fields to use for filter and formatting

                                // get element Id of the parameters
                                ElementId catFieldId = Utils.GetProjectParameterId(curDoc, "Area Category");
                                ElementId nameFieldId = Utils.GetProjectParameterId(curDoc, "Name");                                
                                ElementId areaFieldId = Utils.GetProjectParameterId(curDoc, "Area");

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
                                List<Parameter> paramsFloorSingle = Utils.GetParametersByName(curDoc, paramNames, BuiltInCategory.OST_Areas);
                                Utils.AddFieldsToSchedule(curDoc, newFrameSched, paramsFloorSingle);

                                // create the fields to use for filter and formatting

                                // get element Id of the parameters
                                ElementId catFieldId = Utils.GetProjectParameterId(curDoc, "Area Category");
                                ElementId nameFieldId = Utils.GetProjectParameterId(curDoc, "Name");
                                ElementId levelFieldId = Utils.GetProjectParameterId(curDoc, "Level");
                                ElementId areaFieldId = Utils.GetProjectParameterId(curDoc, "Area");


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

                // set a variable for the equipment schedule

                ViewSchedule schedEquipment = Utils.GetScheduleByNameContains(curDoc, "Roof Ventilation Equipment - Elevation " + Globals.ElevDesignation);

                // check to see if the attic area scheme exists

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
                        // set some variables

                        ViewPlan areaAtticView = Utils.GetAreaPlanByViewFamilyName(curDoc, Globals.ElevDesignation + " Roof Ventilation");

                        // if not, create the area plans

                        if (areaAtticView == null)
                        {
                            string levelWord = "";

                            if (floorNum == 1 && typeFoundation == "Slab")
                            {
                                levelWord = "First Floor";
                            }
                            else if (floorNum == 1 && typeFoundation == "Basement" || typeFoundation == "Crawlspace")
                            {
                                levelWord = "Main Level";
                            }
                            else if (floorNum == 2 && typeFoundation == "Slab")
                            {
                                levelWord = "Second Floor";
                            }
                            else if (floorNum == 2 && typeFoundation == "Basement" || typeFoundation == "Crawlspace")
                            {
                                levelWord = "Upper Level";
                            }

                            Level levelAttic = Utils.GetLevelByName(curDoc, levelWord);

                            ElementId schemeAtticId = schemeAttic.Id;

                            ElementId levelAtticId = levelAttic.Id;

                            View vtAtticAreas = Utils.GetViewTemplateByName(curDoc, "12-Attic Area");

                            ViewPlan areaAttic = ViewPlan.CreateAreaPlan(curDoc, schemeAtticId, levelAtticId);
                            areaAttic.Name = "Roof";
                            areaAttic.ViewTemplateId = vtAtticAreas.Id;

                            uidoc.ActiveView = areaAttic;

                            XYZ insStart = new XYZ(0, 0, 0);

                            double calcOffset = 1.0 * curDoc.ActiveView.Scale;

                            XYZ offset = new XYZ(0, calcOffset, 0);

                            UV insPoint = new UV(insStart.X, insStart.Y);

                            Area areaAttic1 = curDoc.Create.NewArea(areaAttic, insPoint);
                            areaAttic1.Number = "1";
                            areaAttic1.Name = "Attic 1";

                            Area areaAttic2 = curDoc.Create.NewArea(areaAttic, insPoint);
                            areaAttic2.Number = "2";
                            areaAttic2.Name = "Attic 2";}
                        }

                        // if the attic area plans exist, create the schedule

                        // get the area scheme for the schedule
                        AreaScheme curAreaScheme = Utils.GetAreaSchemeByName(curDoc, Globals.ElevDesignation + " Roof Ventilation");

                        // create the new schedule
                        ViewSchedule newAtticSched = Utils.CreateAreaSchedule(curDoc, 
                            "Roof Ventilation Calculations - Elevation " + Globals.ElevDesignation, curAreaScheme);

                    if (typeAttic == "Single")
                    {
                        // create a list of the fields for the schedule
                        List<string> paramNames = new List<string>() { "Area", "1/150 Ratio" }; // ??? add calculated parameter

                        // get the associated parameters & add them to the schedule
                        List<Parameter> paramsAtticSingle = Utils.GetParametersByName(curDoc, paramNames, BuiltInCategory.OST_Areas);
                        Utils.AddFieldsToSchedule(curDoc, newAtticSched, paramsAtticSingle);
                    }

                    else if (typeAttic == "Multiple")
                    {
                        // create a list of the fields for the schedule
                        List<string> paramNames = new List<string>() { "Name", "Area", "1/150 Ratio" }; // ??? add calculated parameter

                        // get the associated parameters & add them to the schedule
                        List<Parameter> paramsAtticMulti = Utils.GetParametersByName(curDoc, paramNames, BuiltInCategory.OST_Areas);
                        Utils.AddFieldsToSchedule(curDoc, newAtticSched, paramsAtticMulti);
                    }

                    if (schedEquipment == null)
                    {
                        // duplicate the first schedule found with Roof Ventilation Equipment in the name
                        List<ViewSchedule> listSched = Utils.GetAllScheduleByNameContains(curDoc, "Roof Ventilation Equipment");

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