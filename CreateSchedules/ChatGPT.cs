using Autodesk.Revit.DB;
using CreateSchedules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace CreateSchedules
if (chbVeneerResult)
{
    DuplicateAndConfigureVeneerSchedule(curDoc);
}

void DuplicateAndConfigureVeneerSchedule(Document doc)
{
    // Find the first schedule with "Exterior Veneer Calculations" in the name
    List<ViewSchedule> listSched = Utils.GetAllScheduleByNameContains(doc, "Exterior Veneer Calculations");
    ViewSchedule dupSched = listSched.FirstOrDefault();

    if (dupSched == null)
    {
        return; // No schedule to duplicate
    }

    // Duplicate the schedule
    ViewSchedule veneerSched = doc.GetElement(dupSched.Duplicate(ViewDuplicateOption.Duplicate)) as ViewSchedule;

    if (veneerSched == null)
    {
        return; // Duplicate operation failed
    }

    // Rename the duplicated schedule
    string originalName = veneerSched.Name;
    string[] schedTitle = originalName.Split('-');

    if (schedTitle.Length >= 1)
    {
        veneerSched.Name = schedTitle[0] + "- Elevation " + Globals.ElevDesignation;
    }

    // Set the design option to the specified elevation designation
    DesignOption curOption = Utils.getDesignOptionByName(doc, "Elevation : " + Globals.ElevDesignation);

    if (curOption != null)
    {
        Parameter doParam = veneerSched.get_Parameter(BuiltInParameter.VIEWER_OPTION_VISIBILITY);

        if (doParam != null)
        {
            doParam.Set(curOption.Id);
        }
    }
}
