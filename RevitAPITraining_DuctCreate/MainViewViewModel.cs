using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Prism.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPITraining_DuctCreate
{
    internal class MainViewViewModel
    {
        private ExternalCommandData _commandData;

        public object DuctTypes { get; } = new List<DuctType>();
        public DuctType SelectedDuctType { get; set; }
        public object Levels { get; } = new List<Level>();
        public Level SelectedLevel { get; set; }
        public int DuctOffset { get; set; }
        public DelegateCommand SaveCommand { get; }
        public List<XYZ> Points { get; } = new List<XYZ>();

        public MainViewViewModel(ExternalCommandData commandData)
        {
            _commandData = commandData;
            DuctTypes = Utils.GetDuctTypes(commandData);
            Levels = Utils.GetLevels(commandData);
            DuctOffset = 100;
            SaveCommand = new DelegateCommand(OnSaveCommand);
            Points = Utils.GetPoints(_commandData, "Выберите точки", ObjectSnapTypes.Endpoints);
        }

        private void OnSaveCommand()
        {
            var doc = _commandData.Application.ActiveUIDocument.Document;

            if (Points.Count != 2 ||
                SelectedDuctType == null ||
                SelectedLevel == null)
            {
                return;
            }
            var firstPoint = Points[0];
            var secondPoint = Points[1];

            using (var ts = new Transaction(doc, "Create duct"))
            {
                ts.Start();
                var duct = Duct.Create(doc, (ElementId)Utils.GetDuctSystemType(_commandData), SelectedDuctType.Id, SelectedLevel.Id, firstPoint, secondPoint);
                duct.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).Set(UnitUtils.ConvertToInternalUnits(DuctOffset, UnitTypeId.Millimeters));
                ts.Commit();
            }
            RaiseCloseRequest();
        }

        public event EventHandler CloseRequest;
        public void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }

    }
    public class Utils
    {
        public static object GetDuctTypes(ExternalCommandData commandData)
        {
            var doc = commandData.Application.ActiveUIDocument.Document;

            var ductTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(DuctType))
                .Cast<DuctType>()
                .ToList();
            return ductTypes;

        }
        public static object GetDuctSystemType(ExternalCommandData commandData)
        {
            var doc = commandData.Application.ActiveUIDocument.Document;

            var systemType = new FilteredElementCollector(doc)
                .OfClass(typeof(MEPSystemType))
                .Cast<MEPSystemType>()
                .FirstOrDefault(m => m.SystemClassification == MEPSystemClassification.SupplyAir);

            return systemType.Id;
        }
        public static object GetLevels(ExternalCommandData commandData)
        {
            var doc = commandData.Application.ActiveUIDocument.Document;

            List<Level> levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .ToList();
            return levels;

        }
        public static List<XYZ> GetPoints(ExternalCommandData commandData, string promptMessage, ObjectSnapTypes objectSnapTypes)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;

            List<XYZ> points = new List<XYZ>();

            while (points.Count < 2)
            {
                XYZ pickedPoint = null;
                try
                {
                    pickedPoint = uidoc.Selection.PickPoint(objectSnapTypes, promptMessage);
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException ex)
                {
                    break;
                }
                points.Add(pickedPoint);
            }
            return points;
        }

    }
}
