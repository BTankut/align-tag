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
    public enum ViewSides { Left, Right };

    public enum TagAlignment
    {
        AlignToBothSides,
        AlignToLeftSide,
        AlignToRightSide
    }

    public class TagLeader
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
            if (_taggedElement != null)
            {
                // Element'in konumunu al
                LocationPoint location = _taggedElement.Location as LocationPoint;
                if (location != null)
                {
                    XYZ elementLocation = location.Point;
                    if (elementLocation.X < viewCenter.X)
                    {
                        _side = ViewSides.Right;  // Element solda, tag sağa
                    }
                    else
                    {
                        _side = ViewSides.Left;   // Element sağda, tag sola
                    }
                }
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

                // Kullanıcıya hizalama seçeneğini sor
                TaskDialog mainDialog = new TaskDialog("Arrange Tags");
                mainDialog.MainInstruction = "Please select alignment option:";
                mainDialog.MainContent = "How would you like to align the selected tags?";
                mainDialog.CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel;
                
                mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Align to Both Sides", "Distribute tags evenly on both sides");
                mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Align to Left Side", "Place all tags on the left side");
                mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Align to Right Side", "Place all tags on the right side");
                
                TaskDialogResult result = mainDialog.Show();

                if (result == TaskDialogResult.Cancel || result == TaskDialogResult.No)
                {
                    return Result.Cancelled;
                }

                TagAlignment alignment = TagAlignment.AlignToBothSides;
                switch (result.ToString())
                {
                    case "CommandLink1":
                        alignment = TagAlignment.AlignToBothSides;
                        break;
                    case "CommandLink2":
                        alignment = TagAlignment.AlignToLeftSide;
                        break;
                    case "CommandLink3":
                        alignment = TagAlignment.AlignToRightSide;
                        break;
                }

                using (Transaction trans = new Transaction(doc, "Arrange Tags"))
                {
                    trans.Start();
                    ArrangeTags(view, selectedTags, alignment);
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

        private void ArrangeTags(View view, List<IndependentTag> tags, TagAlignment alignment)
        {
            if (tags.Count == 0)
                return;

            // Document nesnesini al
            Document doc = view.Document;

            // View sınırlarını al
            BoundingBoxXYZ cropBox = view.CropBox;
            XYZ min = cropBox.Min;
            XYZ max = cropBox.Max;

            // View boyutlarını hesapla
            double viewWidth = max.X - min.X;
            double viewHeight = max.Y - min.Y;
            double centerX = min.X + (viewWidth / 2);

            // X koordinatları için mesafeler (kenarlardan %5 içeride)
            double margin = viewWidth * 0.05;
            double leftX = min.X + margin;
            double rightX = max.X - margin;

            // Tag'leri grupla ve sırala
            var allTags = new List<IndependentTag>();
            var leftSideTags = new List<IndependentTag>();
            var rightSideTags = new List<IndependentTag>();

            switch (alignment)
            {
                case TagAlignment.AlignToBothSides:
                    // İki yana yaslama için mevcut mantık
                    foreach (var tag in tags)
                    {
                        XYZ currentPos = tag.TagHeadPosition;
                        if (currentPos.X < centerX)
                        {
                            leftSideTags.Add(tag);
                        }
                        else
                        {
                            rightSideTags.Add(tag);
                        }
                    }
                    break;

                case TagAlignment.AlignToLeftSide:
                    // Tüm tag'leri sol tarafa ekle
                    leftSideTags.AddRange(tags);
                    break;

                case TagAlignment.AlignToRightSide:
                    // Tüm tag'leri sağ tarafa ekle
                    rightSideTags.AddRange(tags);
                    break;
            }

            // Tag'leri sırala
            if (leftSideTags.Any())
            {
                leftSideTags = leftSideTags.OrderByDescending(t => Math.Abs(t.TagHeadPosition.X - min.X)).ToList();
            }
            if (rightSideTags.Any())
            {
                rightSideTags = rightSideTags.OrderByDescending(t => Math.Abs(t.TagHeadPosition.X - max.X)).ToList();
            }

            // Dikey aralık ve sınırlar
            double verticalSpacing = viewHeight * 0.05;
            double startY = max.Y - (viewHeight * 0.10);  // Üstten %10 aşağıda başla
            double minYPosition = min.Y + (viewHeight * 0.10);  // Alt sınır

            // Sol taraftaki tag'leri yerleştir
            double currentY = startY;
            foreach (var tag in leftSideTags)
            {
                // Y pozisyonunu sınırlar içinde tut
                if (currentY < minYPosition)
                    currentY = startY;

                // Yeni head pozisyonu
                XYZ newHeadPosition = new XYZ(leftX, currentY, 0);

                // Element pozisyonunu al
                XYZ elementPosition;
#if Version2022 || Version2023 || Version2024
                Reference referencedElement = tag.GetTaggedReferences().FirstOrDefault();
                Element element = doc.GetElement(referencedElement);
                if (element.Location is LocationPoint locationPoint)
                {
                    elementPosition = locationPoint.Point;
                }
                else
                {
                    BoundingBoxXYZ bbox = element.get_BoundingBox(view);
                    elementPosition = (bbox.Min + bbox.Max) * 0.5;
                }
#elif Version2019 || Version2020 || Version2021
                Element element = doc.GetElement(tag.TaggedElementId);
                if (element.Location is LocationPoint locationPoint)
                {
                    elementPosition = locationPoint.Point;
                }
                else
                {
                    BoundingBoxXYZ bbox = element.get_BoundingBox(view);
                    elementPosition = (bbox.Min + bbox.Max) * 0.5;
                }
#endif

                // Yatay mesafe ve dik çıkış mesafesi hesaplama
                double horizontalDistance = Math.Abs(newHeadPosition.X - elementPosition.X);
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

                // Tag'i yeni pozisyona taşı
                tag.TagHeadPosition = newHeadPosition;

                currentY -= verticalSpacing;
            }

            // Sağ taraftaki tag'leri yerleştir
            currentY = startY;
            foreach (var tag in rightSideTags)
            {
                // Y pozisyonunu sınırlar içinde tut
                if (currentY < minYPosition)
                    currentY = startY;

                // Yeni head pozisyonu
                XYZ newHeadPosition = new XYZ(rightX, currentY, 0);

                // Element pozisyonunu al
                XYZ elementPosition;
#if Version2022 || Version2023 || Version2024
                Reference referencedElement = tag.GetTaggedReferences().FirstOrDefault();
                Element element = doc.GetElement(referencedElement);
                if (element.Location is LocationPoint locationPoint)
                {
                    elementPosition = locationPoint.Point;
                }
                else
                {
                    BoundingBoxXYZ bbox = element.get_BoundingBox(view);
                    elementPosition = (bbox.Min + bbox.Max) * 0.5;
                }
#elif Version2019 || Version2020 || Version2021
                Element element = doc.GetElement(tag.TaggedElementId);
                if (element.Location is LocationPoint locationPoint)
                {
                    elementPosition = locationPoint.Point;
                }
                else
                {
                    BoundingBoxXYZ bbox = element.get_BoundingBox(view);
                    elementPosition = (bbox.Min + bbox.Max) * 0.5;
                }
#endif

                // Yatay mesafe ve dik çıkış mesafesi hesaplama
                double horizontalDistance = Math.Abs(newHeadPosition.X - elementPosition.X);
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

                // Tag'i yeni pozisyona taşı
                tag.TagHeadPosition = newHeadPosition;

                currentY -= verticalSpacing;
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
    }
}
