using DocumentService.Models;
using Spire.Presentation;
using Spire.Presentation.Charts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using System.Web.Http;
using System.Web.Http.Description;

namespace DocumentService.Controllers
{
    [Route("document")]
    public class DocumentController : ApiController
    {
        [HttpGet]
        [Route("document/test")]
        public IEnumerable<string> Get()
        {
            return new string[] { "hi", "working" };
        }

        private byte[] GenerateDailyInspectionReport(List<Inspection> inspections, bool pdf)
        {
            var accepted = inspections.Where(i => i.Status.Equals(InspectionStatus.Part_Accepted)).Count();
            var rejected = inspections.Where(i => i.Status.Equals(InspectionStatus.Part_Rejected)).Count();
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Templates\inspection_report_template.pptx");
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(filePath, FileFormat.Pptx2010);

            IChart chart = ppt.Slides[1].Shapes[0] as IChart;
            chart.ChartData["A2"].Text = "Total";
            chart.ChartData["A3"].Text = "Accepted";
            chart.ChartData["A4"].Text = "Rejected";
            chart.ChartData["C3"].NumberValue = accepted;
            chart.ChartData["C4"].NumberValue = rejected;
            chart.ChartData["C2"].NumberValue = chart.ChartData["C3"].NumberValue + chart.ChartData["C4"].NumberValue;
            chart.ChartData["B3"].NumberValue = chart.ChartData["C2"].NumberValue - chart.ChartData["C3"].NumberValue;
            chart.ChartData["D2"].NumberValue = chart.ChartData["C2"].NumberValue;
            chart.ChartData["D3"].NumberValue = chart.ChartData["C2"].NumberValue;
            chart.ChartData["E3"].NumberValue = chart.ChartData["C2"].NumberValue - chart.ChartData["C3"].NumberValue;
            chart.ChartData["E4"].NumberValue = chart.ChartData["C2"].NumberValue - chart.ChartData["C3"].NumberValue;

            ppt.Slides.Append();

            Double[] heights = new double[inspections.Count + 1];
            for (int rowIndex = 0; rowIndex <= inspections.Count; rowIndex++)
            {
                heights[rowIndex] = 2;
            }
            int columnCount = 6;
            Double[] widths = new double[columnCount];
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                widths[columnIndex] = 100;
                if (columnIndex == 3 || columnIndex == 4)
                {
                    widths[columnIndex] = 200;
                }
            }
            ITable table = ppt.Slides[2].Shapes.AppendTable(ppt.SlideSize.Size.Width / 2 - 400, 40, widths, heights);
            //ITable table = ppt.Slides[2].Shapes.AppendTable(10, 40, widths, heights);
            table.StylePreset = TableStylePreset.DarkStyle1Accent1;
            // Table Header
            table[0, 0].TextFrame.Text = "S No";
            table[1, 0].TextFrame.Text = "Part Inspected";
            table[2, 0].TextFrame.Text = "Supplier";
            table[3, 0].TextFrame.Text = "Inspector";
            table[4, 0].TextFrame.Text = "Inspection Date";
            table[5, 0].TextFrame.Text = "Status";

            //data fill
            for (int row = 0; row < inspections.Count; row++)
            {
                table[0, row + 1].TextFrame.Text = (row + 1).ToString();
                table[1, row + 1].TextFrame.Text = inspections[row].PartNo;
                table[2, row + 1].TextFrame.Text = inspections[row].SupplierName;
                table[3, row + 1].TextFrame.Text = inspections[row].UserCreated;
                table[4, row + 1].TextFrame.Text = inspections[row].DateCreated.ToString();
                table[5, row + 1].TextFrame.Text = inspections[row].Status.ToString();
            }

            // table style
            for (int column = 0; column < 6; column++)
            {
                for (int row = 1; row <= inspections.Count; row++)
                {
                    table[column, row].TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Calibri");
                    table[column, row].TextFrame.Paragraphs[0].TextRanges[0].Format.FontHeight = 10;
                    table[column, 0].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
                }
            }

            chart.ChartTitle.TextProperties.Text = "Daily Inspection Report";
            chart.Series[1].DataLabels.LabelValueVisible = true;

            Dictionary<string, string> TagValues = new Dictionary<string, string>();
            TagValues.Add("column1", "Total");
            TagValues.Add("column2", "Accepted");
            TagValues.Add("column3", "Rejected");
            ReplaceText(TagValues, ref ppt, 1);

            byte[] pptBuffer = null;
            using (var ms = new MemoryStream())
            {
                if (pdf)
                {
                    ppt.SaveToFile(ms, FileFormat.PDF);
                }
                else
                {
                    ppt.SaveToFile(ms, FileFormat.Pptx2010);
                }
                pptBuffer = ms.ToArray();
            }
            return pptBuffer;
        }

        private byte[] GenerateVendorPerformanceReport(List<Inspection> inspections, bool pdf)
        {
            var accepted = inspections.Where(i => i.Status.Equals(InspectionStatus.Part_Accepted)).Count();
            var rejected = inspections.Where(i => i.Status.Equals(InspectionStatus.Part_Rejected)).Count();
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Templates\inspection_report_template.pptx");
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(filePath, FileFormat.Pptx2010);

            IChart chart = ppt.Slides[1].Shapes[0] as IChart;
            chart.ChartData["A2"].Text = "Total";
            chart.ChartData["A3"].Text = "Accepted";
            chart.ChartData["A4"].Text = "Rejected";
            chart.ChartData["C3"].NumberValue = accepted;
            chart.ChartData["C4"].NumberValue = rejected;
            chart.ChartData["C2"].NumberValue = chart.ChartData["C3"].NumberValue + chart.ChartData["C4"].NumberValue;
            chart.ChartData["B3"].NumberValue = chart.ChartData["C2"].NumberValue - chart.ChartData["C3"].NumberValue;
            chart.ChartData["D2"].NumberValue = chart.ChartData["C2"].NumberValue;
            chart.ChartData["D3"].NumberValue = chart.ChartData["C2"].NumberValue;
            chart.ChartData["E3"].NumberValue = chart.ChartData["C2"].NumberValue - chart.ChartData["C3"].NumberValue;
            chart.ChartData["E4"].NumberValue = chart.ChartData["C2"].NumberValue - chart.ChartData["C3"].NumberValue;

            ppt.Slides.Append();

            Double[] heights = new double[inspections.Count + 1];
            for (int rowIndex = 0; rowIndex <= inspections.Count; rowIndex++)
            {
                heights[rowIndex] = 2;
            }
            int columnCount = 6;
            Double[] widths = new double[columnCount];
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                widths[columnIndex] = 100;
                if (columnIndex == 3 || columnIndex == 4)
                {
                    widths[columnIndex] = 200;
                }
            }
            ITable table = ppt.Slides[2].Shapes.AppendTable(ppt.SlideSize.Size.Width / 2 - 400, 40, widths, heights);
            //ITable table = ppt.Slides[2].Shapes.AppendTable(10, 40, widths, heights);
            table.StylePreset = TableStylePreset.DarkStyle1Accent1;
            // Table Header
            table[0, 0].TextFrame.Text = "S No";
            table[1, 0].TextFrame.Text = "Part Inspected";
            table[2, 0].TextFrame.Text = "Supplier";
            table[3, 0].TextFrame.Text = "Inspector";
            table[4, 0].TextFrame.Text = "Inspection Date";
            table[5, 0].TextFrame.Text = "Status";

            //data fill
            for (int row = 0; row < inspections.Count; row++)
            {
                table[0, row + 1].TextFrame.Text = (row + 1).ToString();
                table[1, row + 1].TextFrame.Text = inspections[row].PartNo;
                table[2, row + 1].TextFrame.Text = inspections[row].SupplierName;
                table[3, row + 1].TextFrame.Text = inspections[row].UserCreated;
                table[4, row + 1].TextFrame.Text = inspections[row].DateCreated.ToString();
                table[5, row + 1].TextFrame.Text = inspections[row].Status.ToString();
            }

            // table style
            for (int column = 0; column < 6; column++)
            {
                for (int row = 1; row <= inspections.Count; row++)
                {
                    table[column, row].TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Calibri");
                    table[column, row].TextFrame.Paragraphs[0].TextRanges[0].Format.FontHeight = 10;
                    table[column, 0].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
                }
            }

            var partNo = inspections.First().PartNo;
            var DistinctPartNo = inspections.All(i => i.PartNo == partNo) ? partNo : "";
            StringBuilder chartName = new StringBuilder("Vendor Performance Report - " + inspections[0].SupplierName);
            chart.ChartTitle.TextProperties.Text = string.IsNullOrEmpty(DistinctPartNo) ? chartName.ToString() : chartName.Append("-" + DistinctPartNo).ToString();

            chart.Series[1].DataLabels.LabelValueVisible = true;

            Dictionary<string, string> TagValues = new Dictionary<string, string>();
            TagValues.Add("column1", "Total");
            TagValues.Add("column2", "Accepted");
            TagValues.Add("column3", "Rejected");
            ReplaceText(TagValues, ref ppt, 1);

            byte[] pptBuffer = null;
            using (var ms = new MemoryStream())
            {
                if (pdf)
                {
                    ppt.SaveToFile(ms, FileFormat.PDF);
                }
                else
                {
                    ppt.SaveToFile(ms, FileFormat.Pptx2010);
                }
                pptBuffer = ms.ToArray();
            }
            return pptBuffer;
        }

        private byte[] GenerateInspectorPerformanceReport(List<Inspection> inspections, bool pdf)
        {
            var accepted = inspections.Where(i => i.Status.Equals(InspectionStatus.Part_Accepted)).Count();
            var rejected = inspections.Where(i => i.Status.Equals(InspectionStatus.Part_Rejected)).Count();
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Templates\inspection_report_template.pptx");
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(filePath, FileFormat.Pptx2010);

            IChart chart = ppt.Slides[1].Shapes[0] as IChart;
            chart.ChartData["A2"].Text = "Total";
            chart.ChartData["A3"].Text = "Accepted";
            chart.ChartData["A4"].Text = "Rejected";
            chart.ChartData["C3"].NumberValue = accepted;
            chart.ChartData["C4"].NumberValue = rejected;
            chart.ChartData["C2"].NumberValue = chart.ChartData["C3"].NumberValue + chart.ChartData["C4"].NumberValue;
            chart.ChartData["B3"].NumberValue = chart.ChartData["C2"].NumberValue - chart.ChartData["C3"].NumberValue;
            chart.ChartData["D2"].NumberValue = chart.ChartData["C2"].NumberValue;
            chart.ChartData["D3"].NumberValue = chart.ChartData["C2"].NumberValue;
            chart.ChartData["E3"].NumberValue = chart.ChartData["C2"].NumberValue - chart.ChartData["C3"].NumberValue;
            chart.ChartData["E4"].NumberValue = chart.ChartData["C2"].NumberValue - chart.ChartData["C3"].NumberValue;

            ppt.Slides.Append();

            Double[] heights = new double[inspections.Count + 1];
            for (int rowIndex = 0; rowIndex <= inspections.Count; rowIndex++)
            {
                heights[rowIndex] = 2;
            }
            int columnCount = 6;
            Double[] widths = new double[columnCount];
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                widths[columnIndex] = 100;
                if (columnIndex == 3 || columnIndex == 4)
                {
                    widths[columnIndex] = 200;
                }
            }
            ITable table = ppt.Slides[2].Shapes.AppendTable(ppt.SlideSize.Size.Width / 2 - 400, 40, widths, heights);
            //ITable table = ppt.Slides[2].Shapes.AppendTable(10, 40, widths, heights);
            table.StylePreset = TableStylePreset.DarkStyle1Accent1;
            // Table Header
            table[0, 0].TextFrame.Text = "S No";
            table[1, 0].TextFrame.Text = "Part Inspected";
            table[2, 0].TextFrame.Text = "Supplier";
            table[3, 0].TextFrame.Text = "Inspector";
            table[4, 0].TextFrame.Text = "Inspection Date";
            table[5, 0].TextFrame.Text = "Status";

            //data fill
            for (int row = 0; row < inspections.Count; row++)
            {
                table[0, row + 1].TextFrame.Text = (row + 1).ToString();
                table[1, row + 1].TextFrame.Text = inspections[row].PartNo;
                table[2, row + 1].TextFrame.Text = inspections[row].SupplierName;
                table[3, row + 1].TextFrame.Text = inspections[row].UserCreated;
                table[4, row + 1].TextFrame.Text = inspections[row].DateCreated.ToString();
                table[5, row + 1].TextFrame.Text = inspections[row].Status.ToString();
            }

            // table style
            for (int column = 0; column < 6; column++)
            {
                for (int row = 1; row <= inspections.Count; row++)
                {
                    table[column, row].TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Calibri");
                    table[column, row].TextFrame.Paragraphs[0].TextRanges[0].Format.FontHeight = 10;
                    table[column, 0].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
                }
            }

            StringBuilder chartName = new StringBuilder("Inspector Performance Report - " + inspections[0].UserCreated.Substring(0, inspections[0].UserCreated.LastIndexOf('@')));
            chart.ChartTitle.TextProperties.Text = chartName.ToString();

            chart.Series[1].DataLabels.LabelValueVisible = true;

            Dictionary<string, string> TagValues = new Dictionary<string, string>();
            TagValues.Add("column1", "Total");
            TagValues.Add("column2", "Accepted");
            TagValues.Add("column3", "Rejected");
            ReplaceText(TagValues, ref ppt, 1);

            byte[] pptBuffer = null;
            using (var ms = new MemoryStream())
            {
                if (pdf)
                {
                    ppt.SaveToFile(ms, FileFormat.PDF);
                }
                else
                {
                    ppt.SaveToFile(ms, FileFormat.Pptx2010);
                }
                pptBuffer = ms.ToArray();
            }
            return pptBuffer;
        }

        [HttpPost]
        [ResponseType(typeof(byte[]))]
        [Route("document/generatepptins")]
        public HttpResponseMessage GeneratePptIns([FromBody]List<Inspection> inspections, bool pdf, string reportType)
        {
            byte[] pptBuffer = null;
            switch (reportType)
            {
                case "dailyReport":
                    pptBuffer = GenerateDailyInspectionReport(inspections, pdf);
                    break;
                case "vendorReport":
                    pptBuffer = GenerateVendorPerformanceReport(inspections, pdf);
                    break;
                case "inspectorPerformanceReport":
                    pptBuffer = GenerateInspectorPerformanceReport(inspections, pdf);
                    break;
            }

            HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.Created);
            response.Content = new ByteArrayContent(pptBuffer);
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            if(pdf)
            {
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
            }
            else
            {
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.presentationml.presentation");
            }
            response.Content.Headers.ContentLength = pptBuffer.Length;
            return response;
        }

        [HttpPost]
        [ResponseType(typeof(byte[]))]
        [Route("document/generateppt")]
        public HttpResponseMessage GeneratePpt([FromBody]IDictionary<string, int> inspectionData)
        {
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Templates\inspection_report_template - Copy.pptx");

            Presentation ppt = new Presentation();
            ppt.LoadFromFile(filePath, FileFormat.Pptx2010);
            var accepted = 20;//inspectionData["accepted"]; //20;
            var rejected = 20;//inspectionData["rejected"]; //40;
            IChart chart = ppt.Slides[1].Shapes[0] as IChart;
            chart.ChartData["A2"].Text = "Total";
            chart.ChartData["A3"].Text = "Accepted";
            chart.ChartData["A4"].Text = "Rejected";
            chart.ChartData["A5"].Text = "Category4";
            chart.ChartData["B5"].NumberValue = 30;
            chart.ChartData["C5"].NumberValue = 30;
            chart.ChartData["D5"].NumberValue = 30;
            chart.ChartData["E5"].NumberValue = 30;
            chart.Series[0].Values = chart.ChartData["B2", "B5"];
            chart.Series[1].Values = chart.ChartData["C2", "C5"];
            chart.Series[2].Values = chart.ChartData["D2", "D5"];
            chart.Series[3].Values = chart.ChartData["E2", "E5"];
            chart.ChartData["C3"].NumberValue = accepted;
            chart.ChartData["C4"].NumberValue = rejected;
            chart.ChartData["C2"].NumberValue = chart.ChartData["C3"].NumberValue + chart.ChartData["C4"].NumberValue;
            chart.ChartData["B3"].NumberValue = chart.ChartData["C2"].NumberValue - chart.ChartData["C3"].NumberValue;
            chart.ChartData["D2"].NumberValue = chart.ChartData["C2"].NumberValue;
            chart.ChartData["D3"].NumberValue = chart.ChartData["C2"].NumberValue;
            chart.ChartData["E3"].NumberValue = chart.ChartData["C2"].NumberValue - chart.ChartData["C3"].NumberValue;
            chart.ChartData["E4"].NumberValue = chart.ChartData["C2"].NumberValue - chart.ChartData["C3"].NumberValue;

            ppt.Slides.Append();

            Double[] widths = new double[] { 100, 100, 150, 100, 100 };
            Double[] heights = new double[] { 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15 };
            ITable table = ppt.Slides[2].Shapes.AppendTable(ppt.SlideSize.Size.Width / 2 - 275, 80, widths, heights);
            //set the style of table
            table.StylePreset = TableStylePreset.LightStyle1Accent2;

            String[,] dataStr = new String[,]{
            {"Inspection Id",    "Part",  "User",    "Status", "Date"},
            {"1",   "Engine",  "Barath",    "Accepted", "2/3/2018"},
            {"2", "Engine",   "Barath",    "Accepted", "2/3/2018"},
            {"3",  "Engine", "Barath",    "Accepted", "2/3/2018"},
            {"4",  "Engine",   "Barath",    "Accepted", "2/3/2018"},
            {"5",   "Engine", "Barath",    "Accepted", "2/3/2018"},
            {"6",    "Engine",   "Barath",    "Rejected", "2/3/2018"},
            {"7",    "Engine",   "Barath",    "Rejected", "2/3/2018"},
            {"8", "Engine",    "Barath",    "Rejected", "2/3/2018"},
            {"9",    "Engine","Barath", "Accepted", "2/3/2018"},
            {"10",    "Engine", "Barath",    "Rejected", "2/3/2018"},
            {"11", "Engine", "Barath",    "Rejected", "2/3/2018"},
            {"12",  "Engine",  "Barath",    "Rejected", "2/3/2018"},
            };
            for (int i = 0; i < 13; i++)
                for (int j = 0; j < 5; j++)
                {
                    //fill the table with data
                    table[j, i].TextFrame.Text = dataStr[i, j];

                    //set the Font
                    table[j, i].TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Arial Narrow");
                }

            //set the alignment of the first row to Center
            for (int i = 0; i < 5; i++)
            {
                table[i, 0].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
            }

            chart.ChartTitle.TextProperties.Text = "Inspection Report";
            chart.Series[1].DataLabels.LabelValueVisible = true;
            //chart.Series[1].DataLabels.Fill.SolidColor.Color = System.Drawing.Color.White;

            Dictionary<string, string> TagValues = new Dictionary<string, string>();
            TagValues.Add("column1", "Total");
            TagValues.Add("column2", "Accepted");
            TagValues.Add("column3", "Rejected");
            ReplaceText(TagValues, ref ppt, 1);


            HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.Created);
            byte[] pptBuffer = null;
            using (var ms = new MemoryStream())
            {
                ppt.SaveToFile(ms, FileFormat.Pptx2010);
                pptBuffer = ms.ToArray();
            }
            response.Content = new ByteArrayContent(pptBuffer);
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.presentationml.presentation");
            response.Content.Headers.ContentLength = pptBuffer.Length;
            return response;
            //ppt.SaveToHttpResponse("out.pptx", FileFormat.Pptx2010, response);
            //ppt.SaveToFile(@"c:\tmp\spiretest5.pptx", FileFormat.Pptx2010);
        }

        public void ReplaceText(Dictionary<string, string> TagValues, ref Presentation presentation, int slideNumber)
        {
            var pSlide = presentation.Slides[slideNumber];
            foreach (IShape curShape in pSlide.Shapes)
            {
                if (curShape is IAutoShape)
                {
                    foreach (TextParagraph tp in (curShape as IAutoShape).TextFrame.Paragraphs)
                    {
                        foreach (var curKey in TagValues.Keys)
                        {
                            if (tp.Text.Contains(curKey))
                            {
                                tp.Text = tp.Text.Replace(curKey, TagValues[curKey]);
                            }
                        }
                    }
                }
            }
        }
        public void ReplaceTags(ISlide pSlide, Dictionary<string, string> TagValues)
        {
            foreach (IShape curShape in pSlide.Shapes)
            {
                if (curShape is IAutoShape)
                {
                    foreach (TextParagraph tp in (curShape as IAutoShape).TextFrame.Paragraphs)
                    {
                        foreach (var curKey in TagValues.Keys)
                        {
                            if (tp.Text.Contains(curKey))
                            {
                                tp.Text = tp.Text.Replace(curKey, TagValues[curKey]);
                            }
                        }
                    }
                }
            }
        }
    }
}
