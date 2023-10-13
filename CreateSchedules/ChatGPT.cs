using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace CreateSchedules
{
    internal class ChatGPT
    {
        ViewSchedule schedIndex = Utils.GetScheduleByNameContains(curDoc, "Sheet Index - Elevation " + Globals.ElevDesignation);

if (chbIndexResult && schedIndex == null)
{
    DuplicateAndRenameSheetIndexSchedule(curDoc);
    }

    void DuplicateAndRenameSheetIndexSchedule(Document doc)
    {
        // Duplicate the first schedule found with "Sheet Index" in the name
        List<ViewSchedule> listSched = Utils.GetAllScheduleByNameContains(doc, "Sheet Index");
        ViewSchedule dupSched = listSched.FirstOrDefault();

        if (dupSched == null)
        {
            return; // No schedule to duplicate
        }

        ViewSchedule indexSched = doc.GetElement(dupSched.Duplicate(ViewDuplicateOption.Duplicate)) as ViewSchedule;

        if (indexSched == null)
        {
            return; // Duplicate operation failed
        }

        // Rename the duplicated schedule to the new elevation
        string originalName = indexSched.Name;
        string[] schedTitle = originalName.Split('C');

        if (schedTitle.Length < 2)
        {
            return; // Invalid schedule name format
        }

        string curTitle = schedTitle[0];
        string lastChar = curTitle.Substring(curTitle.Length - 2);
        string newLast = Globals.ElevDesignation.ToString();

        indexSched.Name = curTitle.Replace(lastChar, newLast);

        // Update the filter value to the new elevation code filter
        ScheduleFilter codeFilter = indexSched.Definition.GetFilter(0);

        if (codeFilter.IsStringValue)
        {
            codeFilter.SetValue(newFilter);
            indexSched.Definition.SetFilter(0, codeFilter);
        }
    }

}
}
