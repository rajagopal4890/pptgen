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

        [HttpPost]
        [ResponseType(typeof(byte[]))]
        [Route("document/generatepptins")]
        public HttpResponseMessage GeneratePptIns([FromBody]List<Inspection> inspections, bool pdf)
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
            response.Content = new ByteArrayContent(pptBuffer);
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.presentationml.presentation");
            response.Content.Headers.ContentLength = pptBuffer.Length;
            return response;
        }

        [HttpPost]
        [ResponseType(typeof(byte[]))]
        [Route("document/generateppt")]
        public HttpResponseMessage GeneratePpt([FromBody]IDictionary<string, int> inspectionData)
        {
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Templates\inspection_report_template.pptx");

            Presentation ppt = new Presentation();
            ppt.LoadFromFile(filePath, FileFormat.Pptx2010);
            var accepted = inspectionData["accepted"]; //20;
            var rejected = inspectionData["rejected"]; //40;
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
            {
                //Dictionary<string, string> TagValues = new Dictionary<string, string>();
                //TagValues.Add("Spire.Presentation for .NET", "Spire.PPT");

                //Presentation presentation = new Presentation();

                //presentation.LoadFromFile("Sample.pptx", FileFormat.Pptx2010);

                ReplaceTags(presentation.Slides[slideNumber], TagValues);

                //presentation.SaveToFile("Result.pptx", FileFormat.Pptx2010);
                //System.Diagnostics.Process.Start("Result.pptx");
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
