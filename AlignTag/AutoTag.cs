#region Namespaces
using System;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Attributes;
using System.Windows.Forms;
using System.Diagnostics;
#endregion

namespace AlignTag
{
    [Transaction(TransactionMode.Manual)]
    public class AutoTagCommand : IExternalCommand
    {
        private UIDocument UIdoc;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get the handle of current document
            UIdoc = commandData.Application.ActiveUIDocument;
            Document document = UIdoc.Document;

            using (TransactionGroup txg = new TransactionGroup(document))
            {
                try
                {
                    txg.Start("Auto Tag Elements");

                    // 1. Element seçimi
                    IList<Reference> selectedReferences = UIdoc.Selection.PickObjects(
                        ObjectType.Element,
                        "Please select elements to tag (ESC to finish)"
                    );
                    
                    if (selectedReferences == null || selectedReferences.Count == 0)
                    {
                        return Result.Cancelled;
                    }

                    // Seçilen elementleri al
                    List<Element> selectedElements = selectedReferences
                        .Select(r => document.GetElement(r))
                        .ToList();

                    // 2. Tag seçimi
                    Reference tagReference = UIdoc.Selection.PickObject(
                        ObjectType.Element,
                        new TagSelectionFilter(),
                        "Please select a tag to use as template"
                    );

                    if (tagReference == null)
                    {
                        return Result.Cancelled;
                    }

                    IndependentTag templateTag = document.GetElement(tagReference) as IndependentTag;
                    if (templateTag == null)
                    {
                        message = "Selected element is not a valid tag.";
                        return Result.Failed;
                    }

                    // 3. Tag yerleştirme pozisyonu
                    XYZ point = UIdoc.Selection.PickPoint("Click to specify vertical position for tags");

                    using (Transaction tx = new Transaction(document))
                    {
                        tx.Start("Create Tags");

                        foreach (Element elem in selectedElements)
                        {
                            try
                            {
                                // Get element's bounding box to find center point
                                BoundingBoxXYZ bbox = elem.get_BoundingBox(document.ActiveView);
                                if (bbox != null)
                                {
                                    // Calculate center X coordinate
                                    double centerX = (bbox.Min.X + bbox.Max.X) / 2;
                                    
                                    // Create new point using element's center X and picked point's Y
                                    XYZ tagPoint = new XYZ(centerX, point.Y, point.Z);

                                    // Get template tag's type ID
                                    ElementId tagTypeId = templateTag.GetTypeId();

                                    // Create new tag
                                    IndependentTag newTag = IndependentTag.Create(
                                        document,
                                        tagTypeId,
                                        document.ActiveView.Id,
                                        new Reference(elem),
                                        true,
                                        TagOrientation.Horizontal,
                                        tagPoint
                                    );

                                    if (newTag != null)
                                    {
                                        // Copy settings from template tag
                                        newTag.TagHeadPosition = tagPoint;
                                        newTag.LeaderEndCondition = templateTag.LeaderEndCondition;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                // Log error but continue with next element
                                Debug.WriteLine($"Error creating tag for element {elem.Id}: {ex.Message}");
                            }
                        }

                        tx.Commit();
                    }

                    txg.Assimilate();
                    return Result.Succeeded;
                }
                catch (OperationCanceledException)
                {
                    if (txg.HasStarted())
                    {
                        txg.RollBack();
                    }
                    return Result.Cancelled;
                }
                catch (Exception ex)
                {
                    message = ex.Message;
                    if (txg.HasStarted())
                    {
                        txg.RollBack();
                    }
                    return Result.Failed;
                }
            }
        }
    }

    public class TagSelectionFilter : ISelectionFilter
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
}
