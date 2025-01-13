#region Namespaces
using System;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Attributes;

#endregion

namespace AlignTag
{
    /// <summary>
    /// Filter to select only tag elements
    /// </summary>
    public class TagFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is IndependentTag;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }

    /// <summary>
    /// Base class for distribute commands
    /// </summary>
    public abstract class DistributeBase : IExternalCommand
    {
        protected abstract string GetCommandName();
        protected abstract double GetCoordinate(XYZ point);
        protected abstract XYZ CreateNewPoint(XYZ currentPoint, double newCoordinate);

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Get the handle of current document
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // Get selected elements
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

                // If nothing is selected, ask user to select elements
                if (selectedIds.Count == 0)
                {
                    try
                    {
                        IList<Reference> selectedReferences = uidoc.Selection.PickObjects(ObjectType.Element, new TagFilter(), "Select tags to distribute");
                        selectedIds = Tools.RevitReferencesToElementIds(doc, selectedReferences);
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        return Result.Cancelled;
                    }
                }

                // Need at least 3 elements to distribute
                if (selectedIds.Count < 3)
                {
                    TaskDialog.Show("Error", "Please select at least 3 tags to distribute.");
                    return Result.Failed;
                }

                using (Transaction trans = new Transaction(doc, $"Distribute Tags {GetCommandName()}"))
                {
                    trans.Start();

                    try
                    {
                        // Get all selected elements
                        List<IndependentTag> tags = selectedIds
                            .Select(id => doc.GetElement(id))
                            .OfType<IndependentTag>()
                            .ToList();

                        if (tags.Count < 3)
                        {
                            TaskDialog.Show("Error", "Please select at least 3 tags to distribute.");
                            return Result.Failed;
                        }

                        // Sort elements by coordinate
                        var sortedTags = tags
                            .OrderBy(tag => GetCoordinate(tag.TagHeadPosition))
                            .ToList();

                        // Get first and last elements (they will stay in place)
                        IndependentTag firstTag = sortedTags.First();
                        IndependentTag lastTag = sortedTags.Last();

                        // Calculate total distance and step size
                        double totalDistance = GetCoordinate(lastTag.TagHeadPosition) - GetCoordinate(firstTag.TagHeadPosition);
                        double step = totalDistance / (sortedTags.Count - 1);

                        // Distribute middle elements
                        for (int i = 1; i < sortedTags.Count - 1; i++)
                        {
                            IndependentTag tag = sortedTags[i];
                            
                            // Calculate new coordinate position
                            double newCoordinate = GetCoordinate(firstTag.TagHeadPosition) + (step * i);
                            
                            // Create new point keeping other coordinates unchanged
                            XYZ newPoint = CreateNewPoint(tag.TagHeadPosition, newCoordinate);
                            
                            // Move element to new position
                            tag.TagHeadPosition = newPoint;
                        }

                        trans.Commit();
                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        message = ex.Message;
                        return Result.Failed;
                    }
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    /// <summary>
    /// Command to distribute tags horizontally with equal spacing
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class DistributeHorizontally : DistributeBase
    {
        protected override string GetCommandName() => "Horizontally";
        
        protected override double GetCoordinate(XYZ point) => point.X;
        
        protected override XYZ CreateNewPoint(XYZ currentPoint, double newX)
            => new XYZ(newX, currentPoint.Y, currentPoint.Z);
    }

    /// <summary>
    /// Command to distribute tags vertically with equal spacing
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class DistributeVertically : DistributeBase
    {
        protected override string GetCommandName() => "Vertically";
        
        protected override double GetCoordinate(XYZ point) => point.Y;
        
        protected override XYZ CreateNewPoint(XYZ currentPoint, double newY)
            => new XYZ(currentPoint.X, newY, currentPoint.Z);
    }
}
