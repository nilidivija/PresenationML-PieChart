using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D= DocumentFormat.OpenXml.Drawing;

namespace PiechartPresentationOpenXML.Data
{
    public class SlideMasterPart1
    {
        public static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1)
        {

            SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
                        SlideMaster slideMaster = new SlideMaster(
                        new CommonSlideData(new ShapeTree(
                        new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new D.TransformGroup()),
                        new Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
                            new NonVisualShapeDrawingProperties(new D.ShapeLocks() { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })),
                        new ShapeProperties(),
                        new TextBody(
                            new D.BodyProperties(),
                            new D.ListStyle(),
                            new D.Paragraph())))),
                            new ColorMap() { 
                                Background1 = D.ColorSchemeIndexValues.Light1, 
                                Text1 = D.ColorSchemeIndexValues.Dark1, 
                                Background2 = D.ColorSchemeIndexValues.Light2, 
                                Text2 = D.ColorSchemeIndexValues.Dark2, 
                                Accent1 = D.ColorSchemeIndexValues.Accent1, 
                                Accent2 = D.ColorSchemeIndexValues.Accent2, 
                                Accent3 = D.ColorSchemeIndexValues.Accent3,
                                Accent4 = D.ColorSchemeIndexValues.Accent4, 
                                Accent5 = D.ColorSchemeIndexValues.Accent5, 
                                Accent6 = D.ColorSchemeIndexValues.Accent6, 
                                Hyperlink = D.ColorSchemeIndexValues.Hyperlink, 
                                FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink },
                        new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
                        new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
            slideMasterPart1.SlideMaster = slideMaster;

            return slideMasterPart1;
      }


    }
}
