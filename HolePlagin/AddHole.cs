using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HolePlagin
{
    [Transaction(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            Document ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault();
            if (ovDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден ОВ файл");
                return Result.Cancelled;
            }

            Document vkDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ВК")).FirstOrDefault();
            if (vkDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден ОВ файл");
                return Result.Cancelled;
            }

            FamilySymbol familySymbol = GetFamilySymbol(arDoc, "Отверстие");
            if (familySymbol == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство \"Отверстия\"");
                return Result.Cancelled;
            }

            List<Duct> ducts = GetListOfDucts(ovDoc);

            List<Pipe> pipes = GetListOfPipes(vkDoc);

            View3D view3D = Get3DView(arDoc);
            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вид");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);
            
            Transaction transaction = new Transaction(arDoc);
            transaction.Start("Расстановка отверстий");
            SetOpeningsByDucts(arDoc, ducts, referenceIntersector, familySymbol);
            SetOpeningsByPipes(arDoc, pipes, referenceIntersector, familySymbol);
            transaction.Commit();

            return Result.Succeeded;
        }

        private FamilySymbol GetFamilySymbol(Document arDoc, string str)
        {
            FamilySymbol familySymbol = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals(str))
                .FirstOrDefault();
            return familySymbol;
        }

        private List<Duct> GetListOfDucts(Document ovDoc)
        {
            List<Duct> ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();
            return ducts;
        }

        private List<Pipe> GetListOfPipes(Document vkDoc)
        {
            List<Pipe> pipes = new FilteredElementCollector(vkDoc)
                .OfClass(typeof(Pipe))
                .OfType<Pipe>()
                .ToList();
            return pipes;
        }

        private View3D Get3DView(Document arDoc)
        {
            View3D view3D = new FilteredElementCollector(arDoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate)
                .FirstOrDefault();
            return view3D;
        }

        private void SetOpeningsByPipes(Document arDoc, List<Pipe> pipes, ReferenceIntersector referenceIntersector, FamilySymbol familySymbol)
        {
            if (!familySymbol.IsActive)
                familySymbol.Activate();
            
            foreach (Pipe pipe in pipes)
            {
                Line curve = (pipe.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                     .Where(x => x.Proximity < curve.Length)
                     .Distinct(new ReferenceWithContextElementEqualityComparer())
                     .ToList();

                foreach (ReferenceWithContext refer in intersections)
                {
                    double proximity = refer.Proximity;
                    Reference reference = refer.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");

                    width.Set(pipe.Diameter + 0.1);
                    height.Set(pipe.Diameter + 0.1);
                }
            }
        }

        private void SetOpeningsByDucts(Document arDoc, List<Duct> ducts, ReferenceIntersector referenceIntersector, FamilySymbol familySymbol)
        {
            if (!familySymbol.IsActive)
                familySymbol.Activate();

            foreach (Duct duct in ducts)
            {
                Line curve = (duct.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                     .Where(x => x.Proximity < curve.Length)
                     .Distinct(new ReferenceWithContextElementEqualityComparer())
                     .ToList();

                foreach (ReferenceWithContext refer in intersections)
                {
                    double proximity = refer.Proximity;
                    Reference reference = refer.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");

                    if (duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM) != null)
                    {
                        width.Set(duct.Diameter + 0.1);
                        height.Set(duct.Diameter + 0.1);
                    }
                    else
                    {
                        width.Set(duct.Width + 0.1);
                        height.Set(duct.Height + 0.1);
                    }
                }
            }
        }

        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        {
            public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();

                var yReference = y.GetReference();

                return xReference.LinkedElementId == yReference.LinkedElementId
                           && xReference.ElementId == yReference.ElementId;
            }

            public int GetHashCode(ReferenceWithContext obj)
            {
                var reference = obj.GetReference();

                unchecked
                {
                    return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
                }
            }
        }
    }
}
