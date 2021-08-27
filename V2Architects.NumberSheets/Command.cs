using Autodesk.Revit.ApplicationServices;
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
        Dictionary<string, string> unicodeLetterMap;
        Dictionary<string, string> letterUnicodeMap = new Dictionary<string, string>()
        {
            { "А", "\u202A" },
            { "Б", "\u202A\u202A" },
            { "В", "\u202A\u202A\u202A" },
            { "Г", "\u202A\u202A\u202A\u202A" },
            { "Д", "\u202A\u202A\u202A\u202A\u202A" },
            { "Е", "\u202A\u202A\u202A\u202A\u202A\u202A" },
            { "Ж", "\u202A\u202A\u202A\u202A\u202A\u202A\u202A" },
            { "З", "\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A" },
            { "И", "\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A" },
            { "Й", "\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A" },
            { "К", "\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A" },
            { "Л", "\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A" },
            { "М", "\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A" },
            { "Н", "\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A" },
            { "О", "\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A" },
            { "П", "\u202C" },
            { "Р", "\u202C\u202C" },
            { "С", "\u202C\u202C\u202C" },
            { "Т", "\u202C\u202C\u202C\u202C" },
            { "У", "\u202C\u202C\u202C\u202C\u202C" },
            { "Ф", "\u202C\u202C\u202C\u202C\u202C\u202C" },
            { "Х", "\u202C\u202C\u202C\u202C\u202C\u202C\u202C" },
            { "Ц", "\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C" },
            { "Ч", "\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C" },
            { "Ш", "\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C" },
            { "Щ", "\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C" },
            { "Э", "\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C" },
            { "Ю", "\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C" },
            { "Я", "\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C" }
        };

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

                bool isLetterOff = App.OpenedRevitProjects[doc];

                using (var t = new Transaction(doc, "Переименование номеров листов"))
                {
                    t.Start();

                    if (isLetterOff)
                    {
                        unicodeLetterMap = InvertDictionary(letterUnicodeMap);
                        ReplaceUnicodeWithLetter(sheets);
                    }
                    else
                    {
                        ReplaceLetterWithUnicode(sheets);
                    }

                    UpdateReviUI();

                    t.Commit();
                }

                App.OpenedRevitProjects[doc] = !isLetterOff;
                App.ButtonLetterTurnOn.Visible = !isLetterOff;
                App.ButtonLetterTurnOff.Visible = isLetterOff;

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

        private void ReplaceUnicodeWithLetter(ICollection<ViewSheet> sheets)
        {
            foreach (ViewSheet sheet in sheets)
            {
                foreach (var pair in unicodeLetterMap)
                {
                    string unicode = pair.Key;
                    string letter = pair.Value;

                    if (sheet.SheetNumber.StartsWith(unicode))
                    {
                        sheet.SheetNumber = sheet.SheetNumber.Replace(unicode, letter);
                        break;
                    }
                }
            }
        }

        private void ReplaceLetterWithUnicode(ICollection<ViewSheet> sheets)
        {
            foreach (ViewSheet sheet in sheets)
            {
                foreach (var pair in letterUnicodeMap)
                {
                    string letter = pair.Key;
                    string unicode = pair.Value;

                    if (sheet.SheetNumber.StartsWith(letter))
                    {
                        sheet.SheetNumber = sheet.SheetNumber.Replace(letter, unicode);
                        break;
                    }
                }
            }
        }

        private void UpdateReviUI()
        {
            DockablePaneId dockablePaneId = DockablePanes.BuiltInDockablePanes.ProjectBrowser;
            var dockablePane = new DockablePane(dockablePaneId);
            dockablePane.Show();
            dockablePane.Hide();
        }

        private Dictionary<string, string> InvertDictionary(Dictionary<string, string> origin)
        {
            var invert = new Dictionary<string, string>();

            foreach (string letter in origin.Keys)
            {
                string unicode = origin[letter];
                invert.Add(unicode, letter);
            }

            return invert;
        }
    }
}
