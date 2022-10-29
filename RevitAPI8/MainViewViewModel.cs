using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Prism.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI8
{
    public class MainViewViewModel
    {
        private ExternalCommandData _commandData;
        private Document _doc;

        public DelegateCommand ExportIFC { get; }
        public DelegateCommand ExportNWC { get; }
        public DelegateCommand ExportImage { get; }

        public MainViewViewModel(ExternalCommandData commandData)
        {
            _commandData = commandData;
            _doc = _commandData.Application.ActiveUIDocument.Document;

            ExportIFC = new DelegateCommand(OnExportIFC);
            ExportNWC = new DelegateCommand(OnExportNWC);
            ExportImage = new DelegateCommand(OnExportImage);
        }

        private void OnExportIFC()
        {
            UIApplication uiapp = _commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            using (var ts = new Transaction(doc, "Export IFC"))
            {
                ts.Start();
                var ifcOption = new IFCExportOptions();
                doc.Export(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "export.ifc", ifcOption);
                ts.Commit();
            }

            RaiseCloseRequest();
        }

        private void OnExportNWC()
        {
            UIApplication uiapp = _commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var nwcOption = new NavisworksExportOptions();
            doc.Export(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "export.nwc", nwcOption);

            RaiseCloseRequest();
        }

        private void OnExportImage()
        {
            UIApplication uiapp = _commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            using (var ts = new Transaction(doc, "Export Image"))
            {
                ts.Start();
                ViewPlan viewPlan = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewPlan))
                    .Cast<ViewPlan>()
                    .FirstOrDefault(v => v.ViewType == ViewType.FloorPlan && v.Name.Equals("Level 1"));

                //uidoc.ActiveView = viewPlan;
                
                var imageOption = new ImageExportOptions
                {
                    ZoomType = ZoomFitType.FitToPage,
                    PixelSize = 2024,
                    FilePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    FitDirection = FitDirectionType.Horizontal,
                    HLRandWFViewsFileType = ImageFileType.PNG,
                    ShadowViewsFileType = ImageFileType.PNG,
                    ImageResolution = ImageResolution.DPI_600,
                    ExportRange = ExportRange.CurrentView,
                };
                doc.ExportImage(imageOption);
                ts.Commit();
            }

            RaiseCloseRequest();
        }


        public event EventHandler CloseRequest;

        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }

    }
}
