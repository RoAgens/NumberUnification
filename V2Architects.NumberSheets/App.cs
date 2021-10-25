﻿using Autodesk.Revit.DB;
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
            var buttonOffData = new PushButtonData(
                "NumberSheetsOff",
                $"Нумерация\nбез префикса",
                typeof(Command).Assembly.Location,
                typeof(Command).FullName
            );

            ButtonLetterTurnOff = panel.AddItem(buttonOffData) as PushButton;
            ButtonLetterTurnOff.LargeImage = GetImageSourceByBitMapFromResource(Properties.Resources.LargeImageOff);
            ButtonLetterTurnOff.Image = GetImageSourceByBitMapFromResource(Properties.Resources.ImageOff);
            ButtonLetterTurnOff.ToolTip = "Скрыть заглавную букву в номере листа.\n" +
                                         $"v{typeof(App).Assembly.GetName().Version}";


            var buttonOnData = new PushButtonData(
                "NumberSheetsOn",
                $"Нумерация\nc префиксом",
                typeof(Command).Assembly.Location,
                typeof(Command).FullName
            );

            ButtonLetterTurnOn = panel.AddItem(buttonOnData) as PushButton;
            ButtonLetterTurnOn.LargeImage = GetImageSourceByBitMapFromResource(Properties.Resources.LargeImageOn);
            ButtonLetterTurnOn.Image = GetImageSourceByBitMapFromResource(Properties.Resources.ImageOn);
            ButtonLetterTurnOn.ToolTip = "Отобразить заглавную букву в номере листа.\n" +
                                        $"v{typeof(App).Assembly.GetName().Version}";
            ButtonLetterTurnOn.Visible = false;
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
            bool isLetterOn;  // Ужастное решение, очень просто запутаться в этих флагах,
                              // если будет возможность, то надо переписать без true / fasle
                              // Решение отвратительно потому, что для корректной работы нужно поменять
                              // код в 3х местах

            if (!OpenedRevitProjects.ContainsKey(doc))
            {
                isLetterOn = true;
                OpenedRevitProjects.Add(doc, isLetterOn);
            }

            isLetterOn = OpenedRevitProjects[doc];

            if (isLetterOn)
            {
                ButtonLetterTurnOn.Visible = false;
                ButtonLetterTurnOff.Visible = true;
            }
            else
            {
                ButtonLetterTurnOn.Visible = true;
                ButtonLetterTurnOff.Visible = false;
            }
        }
    }
}
