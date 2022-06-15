using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using AcAp = Autodesk.AutoCAD.ApplicationServices;
using AcIntCom = Autodesk.AutoCAD.Interop.Common;
using AXDBLib;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Geometry;

namespace AutoCADPrintFromModel
{
    public class ACadPrint : IExtensionApplication
    {
        public static void PrintOneListFromModel(Layout acLayout, PlotSettings acPlSet, PlotInfo acPlInfo, Document acDoc, ObjectId[] ids, int i, int lastNumber)
        {
            // Update the layout
            acLayout.UpgradeOpen();
            acLayout.CopyFrom(acPlSet);

            // Set the plot info as an override since it will not be saved back to the layout
            acPlInfo.OverrideSettings = acPlSet;
            // Validate the plot info
            PlotInfoValidator acPlInfoVdr = new PlotInfoValidator();
            acPlInfoVdr.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
            acPlInfoVdr.Validate(acPlInfo);

            // Выводим в консоль номер печатаемого листа
            acDoc.Editor.WriteMessage("\nНапечатано: " + (i + 1) + " из " + ids.Length);

            // Check to see if a plot is already in progress
            if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
            {
                using (PlotEngine acPlEng = PlotFactory.CreatePublishEngine())
                {
                    // Track the plot progress with a Progress dialog
                    PlotProgressDialog acPlProgDlg = new PlotProgressDialog(false, 1, true);
                    using (acPlProgDlg)
                    {
                        // Start to plot the layout
                        acPlEng.BeginPlot(acPlProgDlg, null);
                        // Define the plot output
                        acPlEng.BeginDocument(acPlInfo, acDoc.Name, null, 1, true, $"d:\\Лист {(lastNumber - i):000}");
                        // Plot the first sheet/layout
                        PlotPageInfo acPlPageInfo = new PlotPageInfo();
                        acPlEng.BeginPage(acPlPageInfo, acPlInfo, true, null);
                        acPlEng.BeginGenerateGraphics(null);
                        acPlEng.EndGenerateGraphics(null);
                        // Finish plotting the sheet/layout
                        acPlEng.EndPage(null);
                        // Finish plotting the document
                        acPlEng.EndDocument(null);
                        // Finish the plot
                        acPlEng.EndPlot(null);
                    }
                }
            }
            if (i < ids.Length - 1)
            {
                System.Threading.Thread.Sleep(35000);
            }
        }

        public void Initialize()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("Плагин пакетной печати чертежей из модели загружен" + Environment.NewLine);
        }

        public void Terminate()
        {

        }
        [CommandMethod("PrintFromModel")]
        public static void PrintFromModel()
        {
            DocumentCollection acDocMgr = AcAp.Application.DocumentManager;
            Document acDoc = acDocMgr.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            ObjectId[] ids;

            // Получаем текущий редактор документов
            Editor activeDocumentEditor = Application.DocumentManager.MdiActiveDocument.Editor;

            // Reference the Layout Manager
            LayoutManager acLayoutMgr;
            acLayoutMgr = LayoutManager.Current;

            // Создаем и назначаем критерии для фильртации выбора
            TypedValue[] activeTypedValueArray = new TypedValue[4];
            activeTypedValueArray.SetValue(new TypedValue((int)DxfCode.Operator, "<and"), 0);
            activeTypedValueArray.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 1);
            activeTypedValueArray.SetValue(new TypedValue((int)DxfCode.LayerName, "PRINT"), 2);
            activeTypedValueArray.SetValue(new TypedValue((int)DxfCode.Operator, "and>"), 3);
            SelectionFilter activeSelectionFilter = new SelectionFilter(activeTypedValueArray);

            // Запрос на выделение объектов по критериям
            PromptSelectionOptions options = new PromptSelectionOptions();
            options.MessageForAdding = "Выделите область модели, где находятся чертежи";
            PromptSelectionResult activeSelectionPrompt = activeDocumentEditor.GetSelection(options, activeSelectionFilter);
            
            PromptIntegerOptions intOptions = new PromptIntegerOptions("Введите номер последнего листа:");
            intOptions.Message = "Введите номер последней страницы:";
            PromptIntegerResult lastNumber = activeDocumentEditor.GetInteger(intOptions);

            using (Transaction tr = acCurDb.TransactionManager.StartTransaction())
            {
                // Если статус запроса в порядке
                if (activeSelectionPrompt.Status == PromptStatus.OK)
                {
                    SelectionSet activeSelectionSet = activeSelectionPrompt.Value;

                    ids = activeSelectionSet.GetObjectIds();
                    // Выводим имя листа и текущего принтера в строке Автокада
                    Layout acLayout = tr.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout), OpenMode.ForRead) as Layout;
                    acDoc.Editor.WriteMessage("\nCurrent layout: " + acLayout.LayoutName);
                    acDoc.Editor.WriteMessage("\nCurrent device name: " + acLayout.PlotConfigurationName);
                    // Output the name of the new device assigned to the layout
                    acDoc.Editor.WriteMessage("\nNew device name: " + acLayout.PlotConfigurationName);
                    acDoc.Editor.WriteMessage($"\nПечать началась в d:\\. Нумерация начинается с листа {lastNumber.Value-ids.Length+1} и заканчивается листом {lastNumber.Value}");

                    for (int i = 0; i < ids.Length; i++)
                    {
                        Polyline pl = tr.GetObject(ids[i], OpenMode.ForRead) as Polyline;
                        //Application.ShowAlertDialog("StartPoint: " + pl.StartPoint.ToString() + "\nEndPoint: " + pl.EndPoint.ToString());

                        // Get the PlotInfo from the layout
                        PlotInfo acPlInfo = new PlotInfo();
                        acPlInfo.Layout = acLayout.ObjectId;

                        // Get a copy of the PlotSettings from the layout
                        PlotSettings acPlSet = new PlotSettings(acLayout.ModelType);
                        acPlSet.CopyFrom(acLayout);

                        // Update the PlotSettings object
                        PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;

                        // Устанавливаем область видимости
                        Point2d plotPointStart = new Point2d(Math.Min(pl.StartPoint.X, pl.EndPoint.X), Math.Min(pl.StartPoint.Y, pl.EndPoint.Y));
                        Point2d plotPointEnd = new Point2d(Math.Max(pl.StartPoint.X, pl.EndPoint.X), Math.Max(pl.StartPoint.Y, pl.EndPoint.Y));
                        Extents2d plotArea = new Extents2d(plotPointStart, plotPointEnd);
                        acPlSetVdr.SetPlotWindowArea(acPlSet, plotArea);
                        acPlSetVdr.SetPlotType(acPlSet, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);

                        // Устанавливаем масштаб печати
                        acPlSetVdr.SetUseStandardScale(acPlSet, true);
                        acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.ScaleToFit);
                        // Центрируем печать
                        acPlSetVdr.SetPlotCentered(acPlSet, true);

                        // Устанавливаем принтер
                        double sheetWidth = Math.Round((Math.Max(pl.StartPoint.X, pl.EndPoint.X) - Math.Min(pl.StartPoint.X, pl.EndPoint.X)), 0);
                        double sheetHeight = Math.Round((Math.Max(pl.StartPoint.Y, pl.EndPoint.Y) - Math.Min(pl.StartPoint.Y, pl.EndPoint.Y)), 0);
                        //A4 вертикальный
                        if (sheetWidth == 2100 && sheetHeight == 2970)
                        {
                            acPlSetVdr.SetPlotConfigurationName(acPlSet, "DWG To PDF.pc3", "ISO_full_bleed_A4_(210.00_x_297.00_mm)");
                            acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees180);
                            PrintOneListFromModel(acLayout, acPlSet, acPlInfo, acDoc, ids, i, lastNumber.Value);
                        }
                        //A4 горизонтальный
                        if (sheetWidth == 2970 && sheetHeight == 2100)
                        {
                            acPlSetVdr.SetPlotConfigurationName(acPlSet, "DWG To PDF.pc3", "ISO_full_bleed_A4_(210.00_x_297.00_mm)");
                            acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees090);
                            PrintOneListFromModel(acLayout, acPlSet, acPlInfo, acDoc, ids, i, lastNumber.Value);
                        }
                        //A3 вертикальный
                        if (sheetWidth == 2970 && sheetHeight == 4200)
                        {
                            acPlSetVdr.SetPlotConfigurationName(acPlSet, "DWG To PDF.pc3", "ISO_full_bleed_A3_(297.00_x_420.00_MM)");
                            acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees180);
                            PrintOneListFromModel(acLayout, acPlSet, acPlInfo, acDoc, ids, i, lastNumber.Value);
                        }
                        //A3 горизонтальный
                        if (sheetWidth == 4200 && sheetHeight == 2970)
                        {
                            acPlSetVdr.SetPlotConfigurationName(acPlSet, "DWG To PDF.pc3", "ISO_full_bleed_A3_(297.00_x_420.00_MM)");
                            acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees090);
                            PrintOneListFromModel(acLayout, acPlSet, acPlInfo, acDoc, ids, i, lastNumber.Value);
                        }
                    }
                }
                else
                {
                    Application.ShowAlertDialog("Number of objects selected: 0");
                }
                //Save the new objects to the database
                tr.Commit();
            }
            acDoc.Editor.WriteMessage("\nПечать завершена");
        }
    }
}