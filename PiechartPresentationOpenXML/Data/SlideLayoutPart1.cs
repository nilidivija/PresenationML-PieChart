using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D= DocumentFormat.OpenXml.Drawing;


namespace PiechartPresentationOpenXML.Data
{
    public class SlideLayoutPart1
    {
        public static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1)
        { 
            SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");
            SlideLayout slideLayout = new SlideLayout(
            new CommonSlideData(new ShapeTree(
              new NonVisualGroupShapeProperties(
              new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()),
              new GroupShapeProperties(new D.TransformGroup()),
              new Shape(
              new NonVisualShapeProperties(
                new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
                new NonVisualShapeDrawingProperties(new D.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
              new ShapeProperties(),
              new TextBody(
                new D.BodyProperties(),
                new D.ListStyle(),
                new D.Paragraph(new D.EndParagraphRunProperties()))))),
            new ColorMapOverride(new D.MasterColorMapping())){ Type = SlideLayoutValues.Blank, Preserve = true };
            slideLayoutPart1.SlideLayout = slideLayout;
            return slideLayoutPart1;

        }

    }
}
