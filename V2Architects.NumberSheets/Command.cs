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
        List<Tuple<string, string>> unicodeLetterMap = new List<Tuple<string, string>>()
        {
            new Tuple<string, string>("\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C", "Я"),
            new Tuple<string, string>("\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C", "Ю"),
            new Tuple<string, string>("\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C", "Э"),
            new Tuple<string, string>("\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C", "Щ"),
            new Tuple<string, string>("\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C", "Ш"),
            new Tuple<string, string>("\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C", "Ч"),
            new Tuple<string, string>("\u202C\u202C\u202C\u202C\u202C\u202C\u202C\u202C", "Ц"),
            new Tuple<string, string>("\u202C\u202C\u202C\u202C\u202C\u202C\u202C", "Х"),
            new Tuple<string, string>("\u202C\u202C\u202C\u202C\u202C\u202C", "Ф"),
            new Tuple<string, string>("\u202C\u202C\u202C\u202C\u202C", "У"),
            new Tuple<string, string>("\u202C\u202C\u202C\u202C", "Т"),
            new Tuple<string, string>("\u202C\u202C\u202C", "С"),
            new Tuple<string, string>("\u202C\u202C", "Р"),
            new Tuple<string, string>("\u202C", "П"),
            new Tuple<string, string>("\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A", "О"),
            new Tuple<string, string>("\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A", "Н"),
            new Tuple<string, string>("\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A", "М"),
            new Tuple<string, string>("\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A", "Л"),
            new Tuple<string, string>("\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A", "К"),
            new Tuple<string, string>("\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A", "Й"),
            new Tuple<string, string>("\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A", "И"),
            new Tuple<string, string>("\u202A\u202A\u202A\u202A\u202A\u202A\u202A\u202A", "З"),
            new Tuple<string, string>("\u202A\u202A\u202A\u202A\u202A\u202A\u202A", "Ж"),
            new Tuple<string, string>("\u202A\u202A\u202A\u202A\u202A\u202A", "Е"),
            new Tuple<string, string>("\u202A\u202A\u202A\u202A\u202A", "Д"),
            new Tuple<string, string>("\u202A\u202A\u202A\u202A", "Г"),
            new Tuple<string, string>("\u202A\u202A\u202A", "В"),
            new Tuple<string, string>("\u202A\u202A", "Б"),
            new Tuple<string, string>("\u202A", "А"),
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

                bool isLetterOn = App.OpenedRevitProjects[doc];

                using (var t = new Transaction(doc, "Переименование номеров листов"))
                {
                    t.Start();

                    if (isLetterOn)
                    {
                        ReplaceLetterWithUnicode(sheets);
                    }
                    else
                    {
                        ReplaceUnicodeWithLetter(sheets);
                    }

                    UpdateReviUI();

                    t.Commit();
                }

                App.OpenedRevitProjects[doc] = !isLetterOn;
                App.ButtonLetterTurnOn.Visible = isLetterOn;
                App.ButtonLetterTurnOff.Visible = !isLetterOn;

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
                    string unicode = pair.Item1;
                    string letter = pair.Item2;

                    // StartsWith по какой-то причине не работает нормально с символами unicode.
                    if (sheet.SheetNumber.Contains(unicode))
                    {
                        // есть некоторая проблема: данное решение может привести к тому что все
                        // символы юникода изменяться на букву.
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
                        sheet.SheetNumber = $"{unicode}{sheet.SheetNumber.Substring(1)}";
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
    }
}
