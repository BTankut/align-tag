using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;

namespace AlignTag
{
    [Transaction(TransactionMode.Manual)]
    class Arrange : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View view = doc.ActiveView;

            try
            {
                // Kullanıcıdan tag seçimini al
                IList<Reference> pickedRefs = uidoc.Selection.PickObjects(ObjectType.Element, new TagFilter(), "Tag'leri seçin");
                if (pickedRefs == null || pickedRefs.Count == 0)
                {
                    return Result.Cancelled;
                }

                // Seçilen elementleri IndependentTag'e dönüştür
                List<IndependentTag> selectedTags = new List<IndependentTag>();
                foreach (Reference reference in pickedRefs)
                {
                    Element elem = doc.GetElement(reference);
                    if (elem is IndependentTag tag)
                    {
                        selectedTags.Add(tag);
                    }
                }

                if (selectedTags.Count == 0)
                {
                    TaskDialog.Show("Uyarı", "Lütfen en az bir tag seçin.");
                    return Result.Cancelled;
                }

                using (Transaction trans = new Transaction(doc))
                {
                    trans.Start("Arrange Tags");
                    ArrangeTags(selectedTags, view, doc);
                    trans.Commit();
                }

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private void ArrangeTags(List<IndependentTag> tags, View view, Document doc)
        {
            //Check the current view
            if (!view.CropBoxActive)
            {
                throw new ErrorMessageException("Please set a crop box to the view");
            }

            //Create two lists of TagLeader
            List<TagLeader> leftTagLeaders = new List<TagLeader>();
            List<TagLeader> rightTagLeaders = new List<TagLeader>();

            foreach (IndependentTag tag in tags)
            {
                TagLeader currentTag = new TagLeader(tag, doc);
                if (currentTag.Side == ViewSides.Left)
                {
                    leftTagLeaders.Add(currentTag);
                }
                else
                {
                    rightTagLeaders.Add(currentTag);
                }
            }

            //Create a list of potential location points for tag headers
            List<XYZ> leftTagHeadPoints = CreateTagPositionPoints(view, leftTagLeaders, ViewSides.Left);
            List<XYZ> rightTagHeadPoints = CreateTagPositionPoints(view, rightTagLeaders, ViewSides.Right);

            //Sort tag by Y position
            leftTagLeaders = leftTagLeaders.OrderBy(x => x.LeaderEnd.X).ToList();
            leftTagLeaders = leftTagLeaders.OrderBy(x => x.LeaderEnd.Y).ToList();

            //place and sort
            PlaceAndSort(leftTagHeadPoints, leftTagLeaders);

            //Sort tag by Y position
            rightTagLeaders = rightTagLeaders.OrderByDescending(x => x.LeaderEnd.X).ToList();
            rightTagLeaders = rightTagLeaders.OrderBy(x => x.LeaderEnd.Y).ToList();

            //place and sort
            PlaceAndSort(rightTagHeadPoints, rightTagLeaders);
        }

        private void PlaceAndSort(List<XYZ> positionPoints, List<TagLeader> tags)
        {
            //place TagLeader
            foreach (TagLeader tag in tags)
            {
                XYZ nearestPoint = FindNearestPoint(positionPoints, tag.TagCenter);
                tag.TagCenter = nearestPoint;

                //remove this point from the list
                positionPoints.Remove(nearestPoint);
            }

            //unCross leaders (2 times)
            UnCross(tags);
            UnCross(tags);

            //update their position
            foreach (TagLeader tag in tags)
            {
                tag.UpdateTagPosition();
            }
        }

        private void UnCross(List<TagLeader> tags)
        {
            foreach (TagLeader tag in tags)
            {
                foreach (TagLeader otherTag in tags)
                {
                    if (tag != otherTag)
                    {
                        if (tag.BaseLine.Intersect(otherTag.BaseLine) == SetComparisonResult.Overlap
                            || tag.BaseLine.Intersect(otherTag.EndLine) == SetComparisonResult.Overlap
                            || tag.EndLine.Intersect(otherTag.BaseLine) == SetComparisonResult.Overlap
                            || tag.EndLine.Intersect(otherTag.EndLine) == SetComparisonResult.Overlap)
                        {
                            XYZ newPosition = tag.TagCenter;
                            tag.TagCenter = otherTag.TagCenter;
                            otherTag.TagCenter = newPosition;
                        }
                    }
                }
            }
        }

        private XYZ FindNearestPoint(List<XYZ> points, XYZ basePoint)
        {
            if (points == null || points.Count == 0)
            {
                // Eğer nokta listesi boşsa, basePoint'i döndür
                return basePoint;
            }

            XYZ nearestPoint = points.FirstOrDefault();
            double nearestDistance = basePoint.DistanceTo(nearestPoint);
            double currentDistance;

            foreach (XYZ point in points)
            {
                currentDistance = basePoint.DistanceTo(point);
                if (currentDistance < nearestDistance)
                {
                    nearestPoint = point;
                    nearestDistance = currentDistance;
                }
            }
            return nearestPoint;
        }

        private List<XYZ> CreateTagPositionPoints(View view, List<TagLeader> tagLeaders, ViewSides side)
        {
            if (tagLeaders.Count == 0)
                return new List<XYZ>();

            // View sınırlarını al
            BoundingBoxXYZ cropBox = view.CropBox;
            XYZ min = cropBox.Min;
            XYZ max = cropBox.Max;

            // Tag'ler arasındaki dikey mesafe
            double verticalSpacing = 0.3; // 30 cm

            // Tag'lerin yerleştirileceği X koordinatı
            double xPosition;
            if (side == ViewSides.Left)
            {
                // Sol taraf için: View'ın sol kenarından tag genişliğinin 2 katı kadar içeride
                xPosition = min.X + Math.Abs(tagLeaders[0].TagWidth) * 2;
            }
            else
            {
                // Sağ taraf için: View'ın sağ kenarından tag genişliğinin 2 katı kadar içeride
                xPosition = max.X - Math.Abs(tagLeaders[0].TagWidth) * 2;
            }

            // Tag'lerin yerleştirileceği Y koordinatlarını hesapla
            List<XYZ> points = new List<XYZ>();
            double totalHeight = (tagLeaders.Count - 1) * verticalSpacing;
            double startY = (max.Y + min.Y - totalHeight) / 2;

            for (int i = 0; i < tagLeaders.Count; i++)
            {
                double yPosition = startY + (i * verticalSpacing);
                points.Add(new XYZ(xPosition, yPosition, 0));
            }

            return points;
        }
    }

    // Tag seçimi için filter class
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

    class TagLeader
    {
        private Document _doc;
        private View _currentView;
        private Element _taggedElement;
        private IndependentTag _tag;

        public TagLeader(IndependentTag tag, Document doc)
        {
            _doc = doc;
            _currentView = _doc.GetElement(tag.OwnerViewId) as View;
            _tag = tag;

            _taggedElement = GetTaggedElement(_doc, _tag);
            _tagHeadPosition = _currentView.CropBox.Transform.Inverse.OfPoint(tag.TagHeadPosition);
            _tagHeadPosition = new XYZ(_tagHeadPosition.X, _tagHeadPosition.Y, 0);
            _leaderEnd = GetLeaderEnd(_taggedElement, _currentView);

            // View'ın ortasını bul
            XYZ viewCenter = (_currentView.CropBox.Max + _currentView.CropBox.Min) / 2;
            
            // Eğer tag'in bağlı olduğu element view'ın ortasından soldaysa, tag sağa
            // Eğer tag'in bağlı olduğu element view'ın ortasından sağdaysa, tag sola
            if (_leaderEnd.X < viewCenter.X)
            {
                _side = ViewSides.Right;  // Element solda, tag sağa
            }
            else
            {
                _side = ViewSides.Left;   // Element sağda, tag sola
            }

            GetTagDimension();
        }

        private XYZ _tagHeadPosition;
        private XYZ _headOffset;

        private XYZ _tagCenter;
        public XYZ TagCenter
        {
            get { return _tagCenter; }
            set
            {
                _tagCenter = value;
                UpdateLeaderPosition();
            }
        }

        private Line _endLine;
        public Line EndLine
        {
            get { return _endLine; }
        }

        private Line _baseLine;
        public Line BaseLine
        {
            get { return _baseLine; }
        }

        private ViewSides _side;
        public ViewSides Side
        {
            get { return _side; }
        }

        private XYZ _elbowPosition;
        public XYZ ElbowPosition
        {
            get { return _elbowPosition; }
        }

        private void UpdateLeaderPosition()
        {
            //Update elbow position
            XYZ AB = _leaderEnd - _tagCenter;
            double mult = AB.X * AB.Y;
            mult = mult / Math.Abs(mult);
            XYZ delta = new XYZ(AB.X - AB.Y * Math.Tan(mult * Math.PI / 4), 0, 0);
            _elbowPosition = _tagCenter + delta;

            //Update lines
            if (_leaderEnd.DistanceTo(_elbowPosition) > _doc.Application.ShortCurveTolerance)
            {
                _endLine = Line.CreateBound(_leaderEnd, _elbowPosition);
            }
            else
            {
                _endLine = Line.CreateBound(new XYZ(0, 0, 0), new XYZ(0, 0, 1));
            }
            if (_elbowPosition.DistanceTo(_tagCenter) > _doc.Application.ShortCurveTolerance)
            {
                _baseLine = Line.CreateBound(_elbowPosition, _tagCenter);
            }
            else
            {
                _baseLine = Line.CreateBound(new XYZ(0, 0, 0), new XYZ(0, 0, 1));
            }
        }

        private XYZ _leaderEnd;
        public XYZ LeaderEnd
        {
            get { return _leaderEnd; }
        }

        private double _tagHeight;
        public double TagHeight
        {
            get { return _tagHeight; }
        }

        private double _tagWidth;
        public double TagWidth
        {
            get { return _tagWidth; }
        }

        private void GetTagDimension()
        {
            BoundingBoxXYZ bbox = _tag.get_BoundingBox(_currentView);
            BoundingBoxXYZ viewBox = _currentView.CropBox;

            _tagHeight = viewBox.Transform.Inverse.OfPoint(bbox.Max).Y - viewBox.Transform.Inverse.OfPoint(bbox.Min).Y;
            _tagWidth = viewBox.Transform.Inverse.OfPoint(bbox.Max).X - viewBox.Transform.Inverse.OfPoint(bbox.Min).X;
            _tagCenter = (viewBox.Transform.Inverse.OfPoint(bbox.Max) + viewBox.Transform.Inverse.OfPoint(bbox.Min)) / 2;
            _tagCenter = new XYZ(_tagCenter.X, _tagCenter.Y, 0);
            _headOffset = _tagHeadPosition - _tagCenter;
        }

        public static Element GetTaggedElement(Document doc, IndependentTag tag)
        {
#if Version2019 || Version2020 || Version2021
            LinkElementId linkElementId = tag.TaggedElementId;
#elif Version2022 || Version2023 || Version2024
            LinkElementId linkElementId = tag.GetTaggedElementIds().FirstOrDefault();
#endif
            Element taggedElement;
            if (linkElementId.HostElementId == ElementId.InvalidElementId)
            {
                RevitLinkInstance linkInstance = doc.GetElement(linkElementId.LinkInstanceId) as RevitLinkInstance;
                Document linkedDocument = linkInstance.GetLinkDocument();

                taggedElement = linkedDocument.GetElement(linkElementId.LinkedElementId);
            }
            else
            {
                taggedElement = doc.GetElement(linkElementId.HostElementId);
            }

            return taggedElement;
        }

        public static XYZ GetLeaderEnd(Element taggedElement, View currentView)
        {
            BoundingBoxXYZ bbox = taggedElement.get_BoundingBox(currentView);
            BoundingBoxXYZ viewBox = currentView.CropBox;

            //Retrive leader end
            XYZ leaderEnd = new XYZ();
            if (bbox != null)
            {
                leaderEnd = (bbox.Max + bbox.Min) / 2;
            }
            else
            {
                leaderEnd = (viewBox.Max + viewBox.Min) / 2 + new XYZ(0.001, 0, 0);
            }

            //Get leader end in view reference
            leaderEnd = viewBox.Transform.Inverse.OfPoint(leaderEnd);
            leaderEnd = new XYZ(Math.Round(leaderEnd.X, 4), Math.Round(leaderEnd.Y, 4), 0);

            return leaderEnd;
        }

        public void UpdateTagPosition()
        {
            _tag.LeaderEndCondition = LeaderEndCondition.Attached;

            // Tag'i kenardan uzaklığını ayarla
            XYZ offsetFromView = new XYZ();
            if (_side == ViewSides.Left)
            {
                // Sol tarafta: View'ın sol kenarından tag genişliğinin 1.5 katı kadar içeride
                offsetFromView = new XYZ(-Math.Abs(_tagWidth) * 1.5, 0, 0);
            }
            else
            {
                // Sağ tarafta: View'ın sağ kenarından tag genişliğinin 1.5 katı kadar içeride
                offsetFromView = new XYZ(Math.Abs(_tagWidth) * 1.5, 0, 0);
            }

            // Tag'in yeni pozisyonunu ayarla
            _tag.TagHeadPosition = _currentView.CropBox.Transform.OfPoint(_headOffset + _tagCenter + offsetFromView);

#if Version2022 || Version2023 || Version2024
            Reference referencedElement = _tag.GetTaggedReferences().FirstOrDefault();
            _tag.SetLeaderElbow(referencedElement, _currentView.CropBox.Transform.OfPoint(_elbowPosition));
#elif Version2019 || Version2020 || Version2021
            _tag.LeaderElbow = _currentView.CropBox.Transform.OfPoint(_elbowPosition);
#endif
        }
    }

    enum ViewSides { Left, Right };
}
