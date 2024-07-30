using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;


namespace PiechartPresentationOpenXML.Data
{
    public class Create
    {          
        public void CreatePresentation(string filePath)
        {
            List<ChartData> chartDataList= ChartDataCollection.ChartDataList;

            PresentationDocument presentationDocument= PresentationDocument.Create(filePath,PresentationDocumentType.Presentation);
            PresentationPart presentationPart= presentationDocument.AddPresentationPart();
            presentationPart.Presentation= new Presentation();

            CreatePresentationParts(presentationPart,chartDataList);

            presentationDocument.Save();
            presentationDocument.Dispose();

        }
        private static void CreatePresentationParts(PresentationPart presentationPart,List<ChartData> chartDataList)
        {
            
            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
            SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
            SlideSize slideSize1 = new SlideSize() { Cx = 12192000, Cy = 6858000, Type = SlideSizeValues.Screen16x9 };
            NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();
            presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);
            SlidePart slidePart1;
            SlideLayoutPart slideLayoutPart1;
            SlideMasterPart slideMasterPart1;
            ThemePart themePart1;
            slidePart1 = SlidePart1.CreateSlidePart(presentationPart);
            slideLayoutPart1 = SlideLayoutPart1.CreateSlideLayoutPart(slidePart1);
            slideMasterPart1 = SlideMasterPart1.CreateSlideMasterPart(slideLayoutPart1);
            themePart1 = ThemePart1.CreateThemePart(slideMasterPart1);
            slideMasterPart1.AddPart(slideLayoutPart1, "rId1");
            presentationPart.AddPart(slideMasterPart1, "rId1");
            presentationPart.AddPart(themePart1, "rId5");
            var chartPart1 = slidePart1.AddNewPart<ChartPart>("rId2");
            ChartPart1.CreateChartPart(chartPart1,chartDataList);
           
            
        }
        
    }
}
