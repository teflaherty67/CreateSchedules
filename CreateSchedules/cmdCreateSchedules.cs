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

                        DesignOption curOption = Utils.getDesignOptionByName(curDoc, "Elevation " + lastChar);
                    }
                }

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

                                        // still need to set values for Area Category & Comments
                                    }
                                    else if (level.ToString() == "Upper Level" || level.ToString() == "Second Floor")
                                    {
                                        // add these areas
                                    }
                                }
                            }
                        }

                        // if the floor area plans exist, create the schedule

                        ElementId areaCategoryId = new ElementId(BuiltInCategory.OST_Areas);

                        AreaScheme curAreaScheme = Utils.GetAreaSchemeByName(curDoc, Globals.ElevDesignation + " Floor");
                        ViewSchedule newFloorSched = Utils.CreateAreaSchedule(curDoc, "Floor Areas - Elevation " + Globals.ElevDesignation, curAreaScheme);

                        if (areaFloorView != null)
                        {
                            if (floorNum == 1)
                            {
                                List<string> paramNames = new List<string>() { "Area Category", "Comments", "Name", "Area", "Number" };
                                List<Parameter> paramsFloorSingle = Utils.GetParametersByName(curDoc, paramNames);

                                Utils.AddFieldsToSchedule(curDoc, newFloorSched, paramsFloorSingle);

                                // find the fields

                                //ScheduleField catFields = Utils.FindScheduleField(newFloorSched, paramsFloorSingle[0]);

                                ScheduleField catField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, paramsFloorSingle[0].Id);
                                ScheduleField commentField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, paramsFloorSingle[1].Id);
                                ScheduleField nameField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, paramsFloorSingle[2].Id);
                                ScheduleField areaField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, paramsFloorSingle[3].Id);
                                ScheduleField numField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, paramsFloorSingle[4].Id);


                                // set the filters

                                ScheduleFilter catFilter = new ScheduleFilter(catField.FieldId, ScheduleFilterType.Contains, "Options");
                                ScheduleFilter areaFilter = new ScheduleFilter(areaField.FieldId, ScheduleFilterType.GreaterThan, "0 SF");

                                // set the sorting

                                ScheduleSortGroupField catSort = new ScheduleSortGroupField(catField.FieldId, ScheduleSortOrder.Ascending);
                                ScheduleSortGroupField commentSort = new ScheduleSortGroupField(commentField.FieldId, ScheduleSortOrder.Ascending);
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
                        // check to see if the frame area plans exist                       
                    }
                }

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