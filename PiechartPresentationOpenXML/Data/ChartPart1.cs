using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using D = DocumentFormat.OpenXml.Drawing;

namespace PiechartPresentationOpenXML.Data
{

    public class ChartPart1
    {
        public static void CreateChartPart(ChartPart chartPart1, List<ChartData> chartDataList){
            
            C.ChartSpace chartSpace1 = new C.ChartSpace(
                                                new C.Date1904(), 
                                                new C.EditingLanguage() { Val = "en-US" },
                                                new C.RoundedCorners() { Val = false });

            C.Chart chart = new C.Chart();
            chart.Append(new C.AutoTitleDeleted() { Val=true });
            C.Title title = new C.Title(
                                new C.ChartText(
                                    new C.RichText(
                                        new D.BodyProperties(),
                                        new D.ListStyle(),
                                        new D.Paragraph(
                                            new D.ParagraphProperties(
                                                new D.DefaultRunProperties()),
                                            new D.Run(
                                                new D.RunProperties(){Language="en-US", FontSize=1862},
                                                new D.Text(){Text="Butter Cake"})))),
                                new C.Overlay(){Val=false});
            
            // Define the 3D view
            C.View3D view3D = new C.View3D();
            view3D.Append(new C.RotateX() { Val = 30 });
            view3D.Append(new C.RotateY() { Val = 0 });
            
            // Intiliazes a new instance of the PlotArea class
            C.PlotArea plotArea = new C.PlotArea();
            plotArea.Append(new C.Layout());

            // the type of Chart 
            C.Pie3DChart pie3DChart = new C.Pie3DChart();
            pie3DChart.Append(new C.VaryColors() { Val = true });
            C.PieChartSeries pieChartSers = new C.PieChartSeries(   
                                    new C.Index() { Val = (UInt32Value)0U },
                                    new C.Order() { Val = (UInt32Value)0U }, 
                                    new C.SeriesText(
                                        new C.StringReference(
                                             new C.Formula() { Text = "Sheet1!$B$1" },
                                             new C.StringCache(
                                                    new C.PointCount() { Val = (UInt32Value)1U },
                                                    new C.StringPoint(
                                                        new C.NumericValue() { Text = "Butter Cake" }){ Index = (UInt32Value)0U }))));
         

            uint rowcount = 0;
            uint count = UInt32.Parse(chartDataList.Count.ToString());
            string endCell = (count + 1).ToString();
            C.ChartShapeProperties chartShapePros = new C.ChartShapeProperties();
           
            // Define cell for lable information
            C.CategoryAxisData cateAxisData = new C.CategoryAxisData();
            C.StringReference stringRef = new C.StringReference();
            stringRef.Append(new C.Formula() { Text = "Sheet1!$A$2:$A$" + endCell });
            C.StringCache stringCache = new C.StringCache();
            stringCache.Append(new C.PointCount() { Val = count });

            // Define cells for value information
            C.Values values = new C.Values();
            C.NumberReference numRef = new C.NumberReference();
            numRef.Append(new C.Formula() { Text = "Sheet1!$B$2:$B$" + endCell });

            C.NumberingCache numCache = new C.NumberingCache();
            numCache.Append(new C.FormatCode() { Text = "General" });
            numCache.Append(new C.PointCount() { Val = count });

            // Fill data for chart
            foreach (var item in chartDataList)
            {
                if (count == 0)
                {
                    chartShapePros.Append(new D.SolidFill(new D.SchemeColor() { Val = item.Color }));
                    pieChartSers.Append(chartShapePros);
                }
                else
                {         
                    C.DataPoint dataPoint = new C.DataPoint();
                    dataPoint.Append(new C.Index() { Val = rowcount });
                    chartShapePros = new C.ChartShapeProperties();
                    chartShapePros.Append(new D.SolidFill(new D.SchemeColor() { Val = item.Color }));
                    dataPoint.Append(chartShapePros);
                    pieChartSers.Append(dataPoint);
                }

                C.StringPoint stringPoint = new C.StringPoint() { Index = rowcount };
                stringPoint.Append(new C.NumericValue() { Text = item.Ingredient });
                stringCache.Append(stringPoint);

                C.NumericPoint numericPoint = new C.NumericPoint() { Index = rowcount };
                numericPoint.Append(new C.NumericValue() { Text = item.Quantity });
                numCache.Append(numericPoint);
                rowcount++;
            }

            // Create c:cat and c:val element 
            stringRef.Append(stringCache);
            cateAxisData.Append(stringRef);
            numRef.Append(numCache);
            values.Append(numRef);

            // Append c:cat and c:val to the end of c:ser element
            pieChartSers.Append(cateAxisData);
            pieChartSers.Append(values);
            //Create Data Labels
            C.DataLabels dataLabels = new C.DataLabels(new C.ShowLegendKey() { Val = false },
                                                        new C.ShowValue() { Val = true },
                                                        new C.ShowCategoryName() { Val = false },
                                                        new C.ShowSeriesName() { Val = false },
                                                        new C.ShowPercent() { Val = false },
                                                        new C.ShowBubbleSize() { Val = false },
                                                        new C.ShowLeaderLines() { Val = false });

            // Append c:ser to the end of c:pie3DChart element
            pie3DChart.Append(pieChartSers);
            pie3DChart.Append(dataLabels);

            // Append c:pie3DChart to the end of s:plotArea element
            plotArea.Append(pie3DChart);

            // create child elements of the c:legend element
            C.Legend legend = new C.Legend(new C.LegendPosition() { Val = C.LegendPositionValues.Bottom },
                                            new C.Overlay() { Val = false },
                                            new C.TextProperties(
                                                new D.BodyProperties(),
                                                new D.ListStyle(),
                                                new D.Paragraph(
                                                    new D.ParagraphProperties(
                                                    new D.DefaultRunProperties(
                                                        new D.SolidFill(
                                                        new D.SchemeColor(){ Val = D.SchemeColorValues.Text1 }),
                                                    new D.LatinFont() { Typeface = "+mn-lt" },
                                                    new D.EastAsianFont() { Typeface = "+mn-ea" },
                                                    new D.ComplexScriptFont() { Typeface = "+mn-cs" }){ FontSize = 1197}),
                                                    new D.EndParagraphRunProperties() { Language = "en-US" })));
           

            // Append c:view3D, c:plotArea and c:legend elements to the end of c:chart element
            chart.Append(title);
            chart.Append(view3D);
            chart.Append(plotArea);
            chart.Append(legend);

            // Append the c:chart element to the end of c:chartSpace element     
            chartSpace1.Append(chart);   

            chartPart1.ChartSpace = chartSpace1;
       
        }


    }
}