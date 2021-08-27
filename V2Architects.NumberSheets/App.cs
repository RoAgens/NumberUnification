using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace V2Architects.NumberSheets
{
    public class App : IExternalApplication
    {
        private string tabName = "V2 Tools";
        private string panelName = "Листы";

        public static PushButton ButtonLetterTurnOn { get; set; }
        public static PushButton ButtonLetterTurnOff { get; set; }
        public static Dictionary<Document, bool> OpenedRevitProjects { get; set; } = new Dictionary<Document, bool>();

        public Result OnStartup(UIControlledApplication revit)
        {
            if (StartupWrongRevitVersion(revit.ControlledApplication.VersionNumber))
            {
                return Result.Cancelled;
            }

            CreateRibbonTab(revit);
            CreateButtons(CreateRibbonPanel(revit));
            revit.ViewActivated += OnViewActivated;

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }


        private static bool StartupWrongRevitVersion(string currentRevitVersion)
        {
            var requiredRevitVersions = new List<string> { "2019", "2020" };
            return !requiredRevitVersions.Contains(currentRevitVersion);
        }

        private void CreateRibbonTab(UIControlledApplication revit)
        {
            try
            {
                revit.CreateRibbonTab(tabName);
            }
            catch { }
        }

        private RibbonPanel CreateRibbonPanel(UIControlledApplication revit)
        {
            foreach (RibbonPanel panel in revit.GetRibbonPanels(tabName))
            {
                if (panel.Name == panelName)
                {
                    return panel;
                }
            }

            return revit.CreateRibbonPanel(tabName, panelName);
        }

        private void CreateButtons(RibbonPanel panel)
        {
            var buttonOnData = new PushButtonData(
                "NumberSheetsOn",
                "Номер листа\n1\u2192А1",
                typeof(Command).Assembly.Location,
                typeof(Command).FullName
            );

            ButtonLetterTurnOn = panel.AddItem(buttonOnData) as PushButton;
            ButtonLetterTurnOn.LargeImage = GetImageSourceByBitMapFromResource(Properties.Resources.LargeImageOn);
            ButtonLetterTurnOn.Image = GetImageSourceByBitMapFromResource(Properties.Resources.ImageOn);
            ButtonLetterTurnOn.ToolTip = "Отобразить заглавную букву в номере листа.\n" +
                                        $"v{typeof(App).Assembly.GetName().Version}";


            var buttonOffData = new PushButtonData(
                "NumberSheetsOff",
                "Номер листа\nA1\u21921",
                typeof(Command).Assembly.Location,
                typeof(Command).FullName
            );

            ButtonLetterTurnOff = panel.AddItem(buttonOffData) as PushButton;
            ButtonLetterTurnOff.LargeImage = GetImageSourceByBitMapFromResource(Properties.Resources.LargeImageOff);
            ButtonLetterTurnOff.Image = GetImageSourceByBitMapFromResource(Properties.Resources.ImageOff);
            ButtonLetterTurnOff.ToolTip = "Скрыть заглавную букву в номере листа.\n" +
                                         $"v{typeof(App).Assembly.GetName().Version}";
            ButtonLetterTurnOff.Visible = false;
        }

        private ImageSource GetImageSourceByBitMapFromResource(Bitmap source)
        {
            return Imaging.CreateBitmapSourceFromHBitmap(
                source.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions()
            );
        }

        private static void OnViewActivated(object sender, ViewActivatedEventArgs e)
        {
            Document doc = e.Document;
            bool isLetterOff;

            if (!OpenedRevitProjects.ContainsKey(doc))
            {
                isLetterOff = true;
                OpenedRevitProjects.Add(doc, isLetterOff);
            }

            isLetterOff = OpenedRevitProjects[doc];

            if (isLetterOff)
            {
                ButtonLetterTurnOn.Visible = true;
                ButtonLetterTurnOff.Visible = false;
            }
            else
            {
                ButtonLetterTurnOn.Visible = false;
                ButtonLetterTurnOff.Visible = true;
            }
        }
    }
}
