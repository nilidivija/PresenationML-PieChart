using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D= DocumentFormat.OpenXml.Drawing;
using C= DocumentFormat.OpenXml.Drawing.Charts;
namespace PiechartPresentationOpenXML.Data
{
    public class SlidePart1
    {
        public static SlidePart CreateSlidePart(PresentationPart presentationPart)
        {
            SlidePart slidePart1=presentationPart.AddNewPart<SlidePart>("rId2");
            slidePart1.Slide=new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new NonVisualGroupShapeProperties(
                                new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                new NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(new D.TransformGroup()),
                            new GraphicFrame(
                                new NonVisualGraphicFrameProperties(
                                    new NonVisualDrawingProperties() {Id=(UInt32Value)1U, Name="Chart 1"},
                                    new NonVisualGraphicFrameDrawingProperties(),
                                    new ApplicationNonVisualDrawingProperties()),
                                new Transform(
                                    new D.Offset(){X=2032000L,Y=719666L},
                                    new D.Extents(){Cx=8128000L, Cy=5418667L}),
                                new D.Graphic(
                                    new D.GraphicData(
                                        new C.ChartReference(){Id="rId2"}){ Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }))
                                        )),
                    new ColorMapOverride(new D.MasterColorMapping()));
           
            return slidePart1;
        }


    }
}