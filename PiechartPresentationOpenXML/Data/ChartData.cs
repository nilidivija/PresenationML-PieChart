using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing;

namespace PiechartPresentationOpenXML.Data
{
    public class ChartData
    {
        public SchemeColorValues Color {get; set;}
        public String Ingredient {get; set;}
        public string Quantity{get;set;}
    }
    public sealed class ChartDataCollection
    {
        static List<ChartData> _chartData;

        public static List<ChartData> ChartDataList
        {
            private set {}
            get {
                return _chartData;
            }
        }
        static ChartDataCollection(){
            Initialize();
        }

        private static void Initialize()
        {
            _chartData= new List<ChartData>{
                new() {
                    Color=SchemeColorValues.Accent1,
                    Ingredient="Flour",
                    Quantity= "30"
                },
                new() {
                    Color=SchemeColorValues.Accent2,
                    Ingredient="Sugar",
                    Quantity= "20"
                },
                new() {
                    Color=SchemeColorValues.Accent3,
                    Ingredient="Egg",
                    Quantity= "40"
                },
                new() {
                    Color=SchemeColorValues.Accent4,
                    Ingredient="Butter",
                    Quantity= "10"
                },
                };
       }
    }

}
