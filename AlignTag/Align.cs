#region Namespaces
using System;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using System.Diagnostics;
#endregion

namespace AlignTag
{
    public class Align
    {
        private UIDocument UIdoc;
        public Result AlignElements(ExternalCommandData commandData, ref string message, AlignType alignType)
        {
            // Get the handle of current document.
            UIdoc = commandData.Application.ActiveUIDocument;
            Document document = UIdoc.Document;

            using (TransactionGroup txg = new TransactionGroup(document))
            {
                try
                {
                    ICollection<ElementId> selectedIds = UIdoc.Selection.GetElementIds();

                    bool empty = false;

                    if (selectedIds.Count == 0)
                    {
                        empty = true;

                        IList<Reference> selectedReferences = UIdoc.Selection.PickObjects(ObjectType.Element, "Pick elements to be aligned");
                        selectedIds = Tools.RevitReferencesToElementIds(document, selectedReferences);
                        UIdoc.Selection.SetElementIds(selectedIds);
                    }

                    AlignTag(alignType, txg, selectedIds, document);

                    // Disselect if the selection was empty to begin with
                    if (empty) selectedIds = new List<ElementId> { ElementId.InvalidElementId };

                    UIdoc.Selection.SetElementIds(selectedIds);

                    // Return Success
                    return Result.Succeeded;
                }

                catch (Autodesk.Revit.Exceptions.OperationCanceledException exceptionCanceled)
                {
                    Console.WriteLine(exceptionCanceled.Message);
                    //message = exceptionCanceled.Message;
                    if (txg.HasStarted())
                    {
                        txg.RollBack();
                    }
                    return Result.Cancelled;
                }
                catch (ErrorMessageException errorEx)
                {
                    // checked exception need to show in error messagebox
                    message = errorEx.Message;
                    if (txg.HasStarted())
                    {
                        txg.RollBack();
                    }
                    return Result.Failed;
                }
                catch (Exception ex)
                {
                    // unchecked exception cause command failed
                    message = ex.Message;
                    //Trace.WriteLine(ex.ToString());
                    if (txg.HasStarted())
                    {
                        txg.RollBack();
                    }
                    return Result.Failed;
                }
            }
        }

        public void AlignTag(AlignType alignType, TransactionGroup txg, ICollection<ElementId> selectedIds, Document document)
        {
            using (Transaction tx = new Transaction(document))
            {
                txg.Start(AlignTypeToText(alignType));

                tx.Start("Prepare tags");
                Debug.WriteLine(DateTime.Now.ToString() + " - Start Prepare tags");

                List<AnnotationElement> annotationElements = RetriveAnnotationElementsFromSelection(document, tx, selectedIds);

                txg.RollBack();
                Debug.WriteLine(DateTime.Now.ToString() + " - Rollback Prepare tags");

                txg.Start(AlignTypeToText(alignType));

                tx.Start(AlignTypeToText(alignType));
                Debug.WriteLine(DateTime.Now.ToString() + " - Start align tags");

                if (annotationElements.Count > 1)
                {
                    AlignAnnotationElements(annotationElements, alignType, document);
                }

                Debug.WriteLine(DateTime.Now.ToString() + " - Commit align tags");

                tx.Commit();

                txg.Assimilate();
            }
        }

        private List<AnnotationElement> RetriveAnnotationElementsFromSelection(Document document, Transaction tx, ICollection<ElementId> ids)
        {
            List<PreparationElement> preparationElements = new List<PreparationElement>();

            List<AnnotationElement> annotationElements = new List<AnnotationElement>();

            //Remove all leader to find the correct tag height and width
            foreach (ElementId id in ids)
            {
                Element e = document.GetElement(id);

                if (e is IndependentTag tag)
                {
                    XYZ elementPosition;
#if Version2022 || Version2023 || Version2024
                    Reference referencedElement = tag.GetTaggedReferences().FirstOrDefault();
                    Element element = document.GetElement(referencedElement);
                    if (element.Location is LocationPoint locationPoint)
                    {
                        elementPosition = locationPoint.Point;
                    }
                    else
                    {
                        BoundingBoxXYZ bbox = element.get_BoundingBox(document.ActiveView);
                        elementPosition = (bbox.Min + bbox.Max) * 0.5;
                    }
#elif Version2019 || Version2020 || Version2021
                    Element element = document.GetElement(tag.TaggedElementId);
                    if (element.Location is LocationPoint locationPoint)
                    {
                        elementPosition = locationPoint.Point;
                    }
                    else
                    {
                        BoundingBoxXYZ bbox = element.get_BoundingBox(document.ActiveView);
                        elementPosition = (bbox.Min + bbox.Max) * 0.5;
                    }
#endif

                    // Yatay mesafe ve dik çıkış mesafesi hesaplama
                    double horizontalDistance = Math.Abs(tag.TagHeadPosition.X - elementPosition.X);
                    double verticalExtension = Math.Min(Math.Max(horizontalDistance * 0.05, 5), 10);

                    // Leader ayarları
                    tag.LeaderEndCondition = LeaderEndCondition.Free;

                    // Önce leader başlangıç noktasını ayarla (elementten dik çıkış için)
                    XYZ leaderEnd = new XYZ(elementPosition.X, elementPosition.Y, 0);
#if Version2022 || Version2023 || Version2024
                    tag.SetLeaderEnd(referencedElement, leaderEnd);
#elif Version2019 || Version2020 || Version2021
                    tag.LeaderEnd = leaderEnd;
#endif

                    // Sonra kırılma noktasını ayarla
                    XYZ elbowPosition = new XYZ(elementPosition.X, elementPosition.Y + verticalExtension, 0);
#if Version2022 || Version2023 || Version2024
                    tag.SetLeaderElbow(referencedElement, elbowPosition);
#elif Version2019 || Version2020 || Version2021
                    tag.LeaderElbow = elbowPosition;
#endif

                    preparationElements.Add(new PreparationElement(e, null));
                }
                else if (e is TextNote note)
                {
                    note.RemoveLeaders();
                    preparationElements.Add(new PreparationElement(e, null));
                }
                else if (e.GetType().IsSubclassOf(typeof(SpatialElementTag)))
                {
                    SpatialElementTag spatialTag = e as SpatialElementTag;

                    XYZ displacementVector = null;

                    if (spatialTag.HasLeader)
                    {
                        displacementVector = spatialTag.LeaderEnd - spatialTag.TagHeadPosition;
                        spatialTag.HasLeader = false;
                    }
                    
                    preparationElements.Add(new PreparationElement(e, displacementVector));
                }
                else
                {
                    preparationElements.Add(new PreparationElement(e, null));
                }
            }

            FailureHandlingOptions options = tx.GetFailureHandlingOptions();

            options.SetFailuresPreprocessor(new TemporaryCommitPreprocessor());
            // Now, showing of any eventual mini-warnings will be postponed until the following transaction.
            tx.Commit(options);

            foreach (PreparationElement e in preparationElements)
            {
                annotationElements.Add(new AnnotationElement(e));
            }

            return annotationElements;
        }

        private void AlignAnnotationElements(List<AnnotationElement> annotationElements, AlignType alignType, Document document)
        {
            View currentView = document.ActiveView;

            switch (alignType)
            {
                case AlignType.Left:
                    {
                        // En soldaki tag'i bul
                        AnnotationElement farthestAnnotation =
                            annotationElements.OrderBy(x => x.UpRight.X).FirstOrDefault();

                        double alignLineX = farthestAnnotation.UpLeft.X;

                        // Elementleri merkeze olan uzaklığa göre sırala 
                        // (ve Align Right'taki gibi davranması için .Reverse() ekle)
                        var sortedElements = annotationElements
                            .OrderByDescending(ae => 
                            {
                                if (ae.Element is IndependentTag tag)
                                {
                                    XYZ elementPosition = GetElementPosition(tag, document, currentView);
                                    // Sol çizgiye olan mutlak uzaklık
                                    return Math.Abs(elementPosition.X - alignLineX);
                                }
                                return 0;
                            })
                            .ToList();

                        // En üstteki tag'den başla (Align Right'ta olduğu gibi)
                        double currentY = sortedElements.First().UpRight.Y;
                        double verticalSpacing = 2; // her bir sonraki etiketi 2 birim alta yerleştireceğiz

                        foreach (var annotationElement in sortedElements)
                        {
                            if (annotationElement.Element is IndependentTag tag)
                            {
                                XYZ elementPosition = GetElementPosition(tag, document, currentView);
                                
                                // Tag'in baş pozisyonu (head) soldaki hizalama çizgisinde
                                XYZ newHeadPosition = new XYZ(alignLineX, currentY, 0);

                                // Leader hesaplamaları
                                double horizontalDistance = Math.Abs(newHeadPosition.X - elementPosition.X);
                                double verticalExtension = Math.Min(Math.Max(horizontalDistance * 0.05, 2), 10);

                                tag.LeaderEndCondition = LeaderEndCondition.Free;
                                XYZ leaderEnd = new XYZ(elementPosition.X, elementPosition.Y, 0);
#if Version2022 || Version2023 || Version2024
                                Reference referencedElement = tag.GetTaggedReferences().FirstOrDefault();
                                tag.SetLeaderEnd(referencedElement, leaderEnd);
#elif Version2019 || Version2020 || Version2021
                                tag.LeaderEnd = leaderEnd;
#endif

                                XYZ elbowPosition = new XYZ(elementPosition.X, 
                                                            elementPosition.Y + verticalExtension,
                                                            0);
#if Version2022 || Version2023 || Version2024
                                tag.SetLeaderElbow(referencedElement, elbowPosition);
#elif Version2019 || Version2020 || Version2021
                                tag.LeaderElbow = elbowPosition;
#endif

                                // Tag head'i yeni konuma yerleştir
                                tag.TagHeadPosition = newHeadPosition;
                            }
                            else
                            {
                                // TextNote vb. ise
                                XYZ resultingPoint = new XYZ(alignLineX, currentY, 0);
                                annotationElement.MoveTo(resultingPoint, AlignType.Left);
                            }
                            currentY -= verticalSpacing;
                        }
                    }
                    break;

                case AlignType.Right:
                    {
                        // En sağdaki tag'i bul
                        AnnotationElement farthestAnnotation =
                            annotationElements.OrderByDescending(x => x.UpRight.X).FirstOrDefault();

                        // Elementleri merkeze olan uzaklığa göre sırala (en yakın element en üstte)
                        var sortedElements = annotationElements
                            .OrderBy(ae => { // OrderBy kullanarak en yakından uzağa sıralıyoruz
                                if (ae.Element is IndependentTag tag)
                                {
                                    XYZ elementPosition = GetElementPosition(tag, document, currentView);
                                    return Math.Abs(elementPosition.X);
                                }
                                return 0;
                            })
                            .Reverse() // Listeyi ters çeviriyoruz, böylece en uzak element en üstte olacak
                            .ToList();

                        // En üstteki tag'den başla
                        double currentY = sortedElements.First().UpRight.Y;
                        double verticalSpacing = 2; // Tag'ler arası mesafe

                        foreach (var annotationElement in sortedElements)
                        {
                            if (annotationElement.Element is IndependentTag tag)
                            {
                                XYZ elementPosition = GetElementPosition(tag, document, currentView);

                                // Yeni head pozisyonu
                                XYZ newHeadPosition = new XYZ(farthestAnnotation.UpRight.X, currentY, 0);
                                
                                // Yatay mesafe ve dik çıkış mesafesi hesaplama
                                double horizontalDistance = Math.Abs(newHeadPosition.X - elementPosition.X);
                                double verticalExtension = Math.Min(Math.Max(horizontalDistance * 0.05, 2), 10);

                                // Leader ayarları
                                tag.LeaderEndCondition = LeaderEndCondition.Free;

                                // Önce leader başlangıç noktasını ayarla (elementten dik çıkış için)
                                XYZ leaderEnd = new XYZ(elementPosition.X, elementPosition.Y, 0);
#if Version2022 || Version2023 || Version2024
                                Reference referencedElement = tag.GetTaggedReferences().FirstOrDefault();
                                tag.SetLeaderEnd(referencedElement, leaderEnd);
#elif Version2019 || Version2020 || Version2021
                                tag.LeaderEnd = leaderEnd;
#endif

                                // Sonra kırılma noktasını ayarla
                                XYZ elbowPosition = new XYZ(elementPosition.X, elementPosition.Y + verticalExtension, 0);
#if Version2022 || Version2023 || Version2024
                                tag.SetLeaderElbow(referencedElement, elbowPosition);
#elif Version2019 || Version2020 || Version2021
                                tag.LeaderElbow = elbowPosition;
#endif

                                // Tag'i yeni pozisyona taşı
                                tag.TagHeadPosition = newHeadPosition;
                            }
                            else
                            {
                                XYZ resultingPoint = new XYZ(farthestAnnotation.UpRight.X, currentY, 0);
                                annotationElement.MoveTo(resultingPoint, AlignType.Right);
                            }

                            // Y pozisyonunu aşağı doğru güncelle
                            currentY -= verticalSpacing;
                        }
                    }
                    break;

                case AlignType.Center:
                    {
                        List<AnnotationElement> sortedAnnotationElements = annotationElements.OrderBy(x => x.UpRight.X).ToList();
                        AnnotationElement rightAnnotation = sortedAnnotationElements.LastOrDefault();
                        AnnotationElement leftAnnotation = sortedAnnotationElements.FirstOrDefault();
                        double XCoord = (rightAnnotation.Center.X + leftAnnotation.Center.X) / 2;

                        // Elementleri merkeze olan uzaklığa göre sırala (en yakın element en üstte)
                        var sortedElements = annotationElements
                            .OrderBy(ae => {
                                if (ae.Element is IndependentTag tag)
                                {
                                    XYZ elementPosition = GetElementPosition(tag, document, currentView);
                                    return Math.Abs(elementPosition.X);
                                }
                                return 0;
                            })
                            .Reverse()
                            .ToList();

                        // Right ve Left ile aynı seviyede başla
                        double currentY = annotationElements.Max(x => x.UpRight.Y);

                        foreach (var annotationElement in sortedElements)
                        {
                            if (annotationElement.Element is IndependentTag tag)
                            {
                                XYZ elementPosition = GetElementPosition(tag, document, currentView);
                                XYZ newHeadPosition = new XYZ(XCoord, currentY, 0);
                                
                                double horizontalDistance = Math.Abs(newHeadPosition.X - elementPosition.X);
                                double verticalExtension = Math.Min(Math.Max(horizontalDistance * 0.05, 5), 10);

                                tag.LeaderEndCondition = LeaderEndCondition.Free;
                                XYZ leaderEnd = new XYZ(elementPosition.X, elementPosition.Y, 0);
#if Version2022 || Version2023 || Version2024
                                Reference referencedElement = tag.GetTaggedReferences().FirstOrDefault();
                                tag.SetLeaderEnd(referencedElement, leaderEnd);
#elif Version2019 || Version2020 || Version2021
                                tag.LeaderEnd = leaderEnd;
#endif
                                XYZ elbowPosition = new XYZ(elementPosition.X, elementPosition.Y + verticalExtension, 0);
#if Version2022 || Version2023 || Version2024
                                tag.SetLeaderElbow(referencedElement, elbowPosition);
#elif Version2019 || Version2020 || Version2021
                                tag.LeaderElbow = elbowPosition;
#endif
                                tag.TagHeadPosition = newHeadPosition;
                            }
                            else
                            {
                                XYZ resultingPoint = new XYZ(XCoord, currentY, 0);
                                annotationElement.MoveTo(resultingPoint, AlignType.Center);
                            }
                            currentY -= 2;
                        }
                    }
                    break;

                case AlignType.Up:
                    {
                        AnnotationElement farthestAnnotation =
                            annotationElements.OrderByDescending(x => x.UpRight.Y).FirstOrDefault();
                        foreach (AnnotationElement annotationElement in annotationElements)
                        {
                            XYZ resultingPoint = new XYZ(annotationElement.UpRight.X, farthestAnnotation.UpRight.Y, 0);
                            annotationElement.MoveTo(resultingPoint, AlignType.Up);
                        }
                    }
                    break;
                case AlignType.Down:
                    {
                        AnnotationElement farthestAnnotation = annotationElements.OrderBy(x => x.UpRight.Y).FirstOrDefault();
                        foreach (AnnotationElement annotationElement in annotationElements)
                        {
                            XYZ resultingPoint = new XYZ(annotationElement.Center.X, farthestAnnotation.UpRight.Y, 0);
                            annotationElement.MoveTo(resultingPoint, AlignType.Down);
                        }
                    }
                    break;
                case AlignType.Middle:
                    {
                        // En üst ve en alt tag'i bul
                        var topElement = annotationElements.OrderByDescending(x => x.UpRight.Y).FirstOrDefault();
                        var bottomElement = annotationElements.OrderBy(x => x.UpRight.Y).FirstOrDefault();
                        // Orta noktayı hesapla
                        double middleY = (topElement.UpRight.Y + bottomElement.UpRight.Y) / 2.0;
                        foreach (AnnotationElement annotationElement in annotationElements)
                        {
                            XYZ resultingPoint = new XYZ(annotationElement.Center.X, middleY, 0);
                            annotationElement.MoveTo(resultingPoint, AlignType.Middle);
                        }
                    }
                    break;
                case AlignType.Vertically:
                    {
                        List<AnnotationElement> sortedAnnotationElements = annotationElements.OrderBy(x => x.UpRight.Y).ToList();
                        AnnotationElement upperAnnotation = sortedAnnotationElements.LastOrDefault();
                        AnnotationElement lowerAnnotation = sortedAnnotationElements.FirstOrDefault();
                        double spacing = (upperAnnotation.Center.Y - lowerAnnotation.Center.Y) / (annotationElements.Count - 1);
                        int i = 0;
                        foreach (AnnotationElement annotationElement in sortedAnnotationElements)
                        {
                            XYZ resultingPoint = new XYZ(annotationElement.Center.X, lowerAnnotation.Center.Y + i * spacing, 0);
                            annotationElement.MoveTo(resultingPoint, AlignType.Vertically);
                            i++;
                        }
                    }
                    break;
                case AlignType.Horizontally:
                    {
                        List<AnnotationElement> sortedAnnotationElements = annotationElements.OrderBy(x => x.UpRight.X).ToList();
                        AnnotationElement rightAnnotation = sortedAnnotationElements.LastOrDefault();
                        AnnotationElement leftAnnotation = sortedAnnotationElements.FirstOrDefault();
                        double spacing = (rightAnnotation.Center.X - leftAnnotation.Center.X) / (annotationElements.Count - 1);
                        int i = 0;
                        foreach (AnnotationElement annotationElement in sortedAnnotationElements)
                        {
                            XYZ resultingPoint = new XYZ(leftAnnotation.Center.X + i * spacing, annotationElement.Center.Y, 0);
                            annotationElement.MoveTo(resultingPoint, AlignType.Horizontally);
                            i++;
                        }
                    }
                    break;
            }
        }

        private string AlignTypeToText(AlignType alignType)
        {
            switch (alignType)
            {
                case AlignType.Left:
                    return "Align Left";
                case AlignType.Right:
                    return "Align Right";
                case AlignType.Up:
                    return "Align Up";
                case AlignType.Down:
                    return "Align Down";
                case AlignType.Center:
                    return "Align Center";
                case AlignType.Middle:
                    return "Align Middle";
                case AlignType.Vertically:
                    return "Distribute Vertically";
                case AlignType.Horizontally:
                    return "Distribute Horizontally";
                default:
                    return "Align";
            }
        }

        private XYZ GetElementPosition(IndependentTag tag, Document document, View view)
        {
            XYZ elementPosition;
#if Version2022 || Version2023 || Version2024
            Reference referencedElement = tag.GetTaggedReferences().FirstOrDefault();
            Element element = document.GetElement(referencedElement);
#elif Version2019 || Version2020 || Version2021
            Element element = document.GetElement(tag.TaggedElementId);
#endif
            if (element.Location is LocationPoint locationPoint)
            {
                elementPosition = locationPoint.Point;
            }
            else
            {
                BoundingBoxXYZ bbox = element.get_BoundingBox(view);
                elementPosition = (bbox.Min + bbox.Max) * 0.5;
            }

            return elementPosition;
        }
    }
}
