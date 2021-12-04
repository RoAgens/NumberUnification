using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace V2Architects.NumberSheets
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]
    class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            try
            {
                var sheets = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .ToList();

                if (sheets.Count == 0)
                {
                    var taskDialog = new TaskDialog("Ошибка")
                    {
                        TitleAutoPrefix = false,
                        MainIcon = TaskDialogIcon.TaskDialogIconError,
                        MainInstruction = "В проекте нет листов!"
                    };
                    taskDialog.Show();
                    return Result.Failed;
                }


                //bool isLetterOn = App.OpenedRevitProjects[doc];

                using (var t = new Transaction(doc, "Переименование номеров листов"))
                {
                    t.Start();

                    var numberManager = new NumberService(sheets);
                    numberManager.ReplacePrefixWithUnicode();
                    numberManager.ReplaceUnicodeWithPrefix();

                    //if (isLetterOn)
                    //{
                    //    numberManager.ReplaceSubgroupWithUnicode(out List<string> wrongNubmers);
                    //}
                    //else
                    //{
                    //    numberManager.ReplaceUnicodeWithSubgroup(out List<string> wrongNubmers);
                    //}

                    UpdateReviUI();

                    t.Commit();
                }

                //App.OpenedRevitProjects[doc] = !isLetterOn;
                //App.ButtonLetterTurnOn.Visible = isLetterOn;
                //App.ButtonLetterTurnOff.Visible = !isLetterOn;

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                TaskDialog.Show("Отмена", "Операция отменена пользователем.");
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Ошибка", message);
                return Result.Failed;
            }
        }

        private void UpdateReviUI()
        {
            DockablePaneId dockablePaneId = DockablePanes.BuiltInDockablePanes.ProjectBrowser;
            var dockablePane = new DockablePane(dockablePaneId);
            dockablePane.Show();
            dockablePane.Hide();
        }
    }
}
