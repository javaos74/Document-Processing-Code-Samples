using Azure.AI.FormRecognizer.DocumentAnalysis;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using UiPath.DocumentProcessing.Contracts;
using UiPath.DocumentProcessing.Contracts.DataExtraction;
using UiPath.DocumentProcessing.Contracts.Dom;
using UiPath.DocumentProcessing.Contracts.Results;
using UiPath.DocumentProcessing.Contracts.Taxonomy;
using SampleActivities.Basic.OCR;
using static SampleActivities.Basic.OCR.ClovaOCRResultHelper;
using Newtonsoft.Json;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Net;
using System.Text;
using System.Text.Json;

namespace SampleActivities.Basic.DataExtraction
{
    enum ClovaFieldType
    {
        Text =0,
        Date = 1,
        Tel = 2,
        Time = 3
    }

    internal class ClovaField
    {
        [JsonProperty("text")]
        public string Text { get; set; }
        [JsonProperty("formatted")]
        public Formatted Formatted { get; set; }
        [JsonProperty("boundingPolys")]
        public BoundingPoly[] BoundingPolys { get; set; }

        private Rectangle _box = Rectangle.Empty;
        public Rectangle    Box {  
            get
            {
                if (_box != Rectangle.Empty)
                {
                    return _box;
                }
                else
                {
                    foreach (var p in BoundingPolys)
                    {
                        if (_box == Rectangle.Empty)
                        {
                            _box = new Rectangle((int)Math.Floor(p.Vertices[0].X)-2, 
                                                             (int)Math.Floor(p.Vertices[0].Y)-2,
                                                             (int)Math.Ceiling(Math.Abs(p.Vertices[0].X - p.Vertices[1].X))+4,
                                                             (int)Math.Ceiling(Math.Abs(p.Vertices[0].Y - p.Vertices[2].Y))+4);
                        }
                        else
                        {
                            _box = Rectangle.Union(_box, new Rectangle((int)Math.Floor(p.Vertices[0].X)-2, 
                                                                 (int)Math.Floor(p.Vertices[0].Y)-2,
                                                                 (int)Math.Ceiling(Math.Abs(p.Vertices[0].X - p.Vertices[1].X))+4,
                                                                 (int)Math.Ceiling(Math.Abs(p.Vertices[0].Y - p.Vertices[2].Y))+4) );
                        }
                    }
                }
                return _box;
            } 
        }
        override public String ToString()
        {
            return $"text: {Text}, formated: {Formatted} Box: {Box.ToString()}";
        }
    }

    internal class Formatted
    {
        [JsonProperty("year")]
        public string Year { get; set; }
        [JsonProperty("month")]
        public string Month { get; set; }
        [JsonProperty("day")]
        public string Day { get; set; }
        [JsonProperty("value")]
        public string Value { get; set; }
        [JsonProperty("hour")]
        public string Hour { get; set; }
        [JsonProperty("minute")]
        public string Minute { get; set; }
        [JsonProperty("second")]
        public string Second { get; set; }


        private DateTime _date;
        public DateTime Date { 
            get {
                if (!string.IsNullOrEmpty(Year) && !string.IsNullOrEmpty(Month) && !string.IsNullOrEmpty(Day))
                {
                    _date = new DateTime(Convert.ToInt32(Year), Convert.ToInt32(Month), Convert.ToInt32(Day));
                    return _date;
                }
                else
                {
                    return DateTime.MinValue;
                }
            } 
        }
        public ClovaFieldType FieldType { 
            get 
            {
                ClovaFieldType type = ClovaFieldType.Text;
                if ( !string.IsNullOrEmpty(Year) && !string.IsNullOrEmpty(Month) && !string.IsNullOrEmpty(Day) )
                {
                    type = ClovaFieldType.Date;
                    _date = new DateTime( Convert.ToInt32(Year), Convert.ToInt32(Month), Convert.ToInt32(Day) );
                }
                else if ( ! string.IsNullOrEmpty(Hour) && !string.IsNullOrEmpty(Minute) && !string.IsNullOrEmpty(Second) )
                {
                    type = ClovaFieldType.Time;
                }
                return type;
            } 
        }
        public override string ToString()
        {
            if (FieldType == ClovaFieldType.Date)
                return $"date: {_date.ToString("yyyy-MM-dd")}";
            else if (FieldType == ClovaFieldType.Time)
                return $"time : {Hour}:{Minute}:{Second}";
            else
                return $"value: {Value}";
        }

    }

    internal class BoundingPoly
    {
        [JsonProperty("vertices")]
        public Vertex[] Vertices { get; set; }
    }

    internal class Vertex
    {
        [JsonProperty("x")]
        public float X { get; set; }
        [JsonProperty("y")]
        public float Y { get; set; }
    }

   
    [DisplayName("Clova 주민등록증 Extractor")]
    public class ClovaIDCardExtractor : ExtractorAsyncCodeActivity
    {
        [Category("Server")]
        [RequiredArgument]
        [Description("ML모델 서비스 endpoint 정보")]
        public InArgument<string> Endpoint { get; set; }

        [Category("Server")]
        [RequiredArgument]
        [Description("ML모델 서비스 endpoint ApiKey 정보 ")]
        public InArgument<string> ApiKey { get; set; }

        ExtractorResult result;
        List<PageLayout> pages;


        public override Task<ExtractorDocumentTypeCapabilities[]> GetCapabilities()
        {
#if DEBUG2
            Debug.WriteLine("GetCapabilities called");
#endif
            //Azure Form Recognizer invoice fields definition 
            List<ExtractorFieldCapability> fields = new List<ExtractorFieldCapability>();

            fields.Add(new ExtractorFieldCapability { FieldId = "name", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "personalNum", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "address", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "issueDate", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "authority", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });

            return Task.FromResult(new[] {
                new ExtractorDocumentTypeCapabilities{
                    DocumentTypeId = "clova.idCard.ic",
                    Fields = fields.ToArray()
                }
            });
        }
        public override Boolean ProvidesCapabilities()
        {
            return true;
        }

        protected override IAsyncResult BeginExecute(AsyncCodeActivityContext context, AsyncCallback callback, object state)
        {
            //get arguments passed to DataExtractionScope 
            ExtractorDocumentType documentType = ExtractorDocumentType.Get(context);
            ResultsDocumentBounds documentBounds = DocumentBounds.Get(context);
            string text = DocumentText.Get(context);
            Document document = DocumentObjectModel.Get(context);
            string documentPath = DocumentPath.Get(context);
            string endpoint = Endpoint.Get(context);
            string apiKey = ApiKey.Get(context);
            this.pages = new List<PageLayout>();

            var task = new Task(_ => Execute(documentType, documentBounds, text, document, documentPath, endpoint, apiKey), state);
            task.Start();
            if (callback != null)
            {
                task.ContinueWith(s => callback(s));
                task.Wait();
            }
            return task;
        }

        protected override async void EndExecute(AsyncCodeActivityContext context, IAsyncResult result)
        {
            var task = (Task)result;
            ExtractorResult.Set(context, this.result);
            await task;
        }

        protected async void Execute(ExtractorDocumentType documentType, ResultsDocumentBounds documentBounds,
                                    string text, Document document, string documentPath,
                                    string endPoint, string apiKey)
        {

            this.result = await ComputeResult(documentType, documentBounds, text, document, documentPath, endPoint, apiKey);
        }

        private async Task<ExtractorResult>  ComputeResult(ExtractorDocumentType documentType, ResultsDocumentBounds documentBounds, 
                                                string text, Document dom, string documentPath,
                                                string endpoint, string apikey)
        {
            var extractorResult = new ExtractorResult();
            var resultsDataPoints = new List<ResultsDataPoint>();

            var client = new UiPathHttpClient(endpoint);
            client.setSecret(apikey);
            var reqBody = new RequestBody();
            var reqimg = new RequestImage(Path.GetFileName(documentPath));
            reqimg.format = Path.GetExtension(documentPath);
            reqBody.images.Add(reqimg);
            client.AddField("message", JsonConvert.SerializeObject(reqBody));
            client.AddFile(documentPath);
            foreach (var x in dom.Pages)
            {
                this.pages.Add(new PageLayout((double)x.Size.Width, (double)x.Size.Height, string.Empty));
            }
            var resp = await client.Upload();
            if (resp.status == HttpStatusCode.OK)
            {
                StringBuilder sb = new StringBuilder();
                JObject respJson = JObject.Parse(resp.body);
#if DEBUG2
                Console.WriteLine("IDcard result : " + resp.body);
#endif
                JObject blocks = (JObject)respJson["images"][0]["idCard"]["result"]["ic"];
                //name 
                foreach (var du_field in documentType.Fields)
                {
                    if (du_field.FieldId.Equals("name") && blocks.ContainsKey("name"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["name"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("personalNum") && blocks.ContainsKey("personalNum"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["personalNum"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("address") && blocks.ContainsKey("address"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["address"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("issueDate") && blocks.ContainsKey("issueDate"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["issueDate"][0].ToString());
                        resultsDataPoints.Add(CreateDateFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("authority") && blocks.ContainsKey("authority"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["authority"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                }
            }
            extractorResult.DataPoints = resultsDataPoints.ToArray();
            return extractorResult;
        }

        private static ResultsDataPoint CreateTextFieldDataPoint(Field du_field, ClovaField clova_field, Document dom, PageLayout[] pages)
        {
            // how to map word index, now just default 0 
            var resultValue = CreateResultsValue(0, dom, clova_field, pages);
            return new ResultsDataPoint(
                du_field.FieldId,
                du_field.FieldName,
                du_field.Type,
                new[] { resultValue });
        }
         private static ResultsDataPoint CreateDateFieldDataPoint(Field du_field, ClovaField clova_field, Document dom, PageLayout[] pages)
        {
            // TODO
            DateTimeOffset _date = clova_field.Formatted.Date ;
            var derivedFields = ResultsDerivedField.CreateDerivedFieldsForDate(_date.Day, _date.Month, _date.Year);
            var resultValue = CreateResultsValue(0, dom, clova_field, pages);
            resultValue.DerivedFields = derivedFields;

            return new ResultsDataPoint(
                du_field.FieldId,
                du_field.FieldName,
                du_field.Type,
                new[] { resultValue });
        }

        private static ResultsValue CreateResultsValue(int wordIndex, Document dom, ClovaField clova_field, PageLayout[] pages)
        {
            float ocr_confidence = 0f;
            Rectangle rect = clova_field.Box;

            var words = dom.Pages[0].Sections.SelectMany(s => s.WordGroups)
                .SelectMany(w => w.Words).Where(t => rect.Contains(new Rectangle((Int32)t.Box.Left, (Int32)t.Box.Top, (Int32)t.Box.Width, (Int32)t.Box.Height))).ToArray();

#if DEBUG2
            Console.WriteLine($"{words.Length} words found");
#endif
            List<Box> boxes = new List<Box>();
            List<ResultsValueTokens> tokens = new List<ResultsValueTokens>();
            foreach (var w in words)
            {
                boxes.Add(w.Box);
#if DEBUG2
                Console.WriteLine($"Box : {w.Box.Left}, {w.Box.Top}, {w.Box.Width}, {w.Box.Height}");
#endif
                ocr_confidence += w.OcrConfidence;
            }
#if DEBUG2
            Console.WriteLine($"{boxes.Count} is found");
#endif
            if (boxes.Count == 0)
            {
                boxes.Add(Box.CreateChecked(0, 0, 0, 0));
            }
            tokens.Add(new ResultsValueTokens(0, clova_field.Text.Length,
                                0,
                                (float)dom.Pages[0].Size.Width,
                                (float)dom.Pages[0].Size.Height, boxes.ToArray()));
            var reference = new ResultsContentReference(0, clova_field.Text.Length, tokens.ToArray());
            return new ResultsValue(clova_field.Text, reference, 0.0f, (float)ocr_confidence / boxes.Count); 
        }
    }
}
