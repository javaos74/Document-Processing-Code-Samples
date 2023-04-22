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

   
    [DisplayName("Clova 운전면허증 Extractor")]
    public class ClovaDriverLicenseExtractor : ExtractorAsyncCodeActivity
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
#if DEBUG
            Debug.WriteLine("GetCapabilities called");
#endif
            //Azure Form Recognizer invoice fields definition 
            List<ExtractorFieldCapability> fields = new List<ExtractorFieldCapability>();

            fields.Add(new ExtractorFieldCapability { FieldId = "type", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "num", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "name", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "personalNum", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "address", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "renewStartDate", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "renewEndDate", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "condition", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "code", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "organDonation", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "issueDate", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "authority", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });

            return Task.FromResult(new[] {
                new ExtractorDocumentTypeCapabilities{
                    DocumentTypeId = "clova.idCard.dl",
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
            var reqimg = new RequestImage(Path.GetFileNameWithoutExtension(documentPath));
            reqimg.format = Path.GetExtension(documentPath).Substring(1);
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
                JObject blocks = (JObject)respJson["images"][0]["idCard"]["result"]["dl"];
#if DEBUG2
                Console.WriteLine("driverlicense result : " + resp.body);
#endif
                //name 
                foreach (var du_field in documentType.Fields)
                {
#if DEBUG2
                    Console.WriteLine($"field: {du_field.FieldId}");
#endif
                    if (du_field.FieldId.Equals("type") && blocks.ContainsKey("type"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["type"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("num") && blocks.ContainsKey("num"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["num"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("name") && blocks.ContainsKey("name"))
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
                    else if (du_field.FieldId.Equals("renewStartDate") && blocks.ContainsKey("renewStartDate"))   
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["renewStartDate"][0].ToString());
                        if( clova_field.Formatted.FieldType == ClovaFieldType.Date)
                            resultsDataPoints.Add(CreateDateFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                        else
                            resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));

                    }
                    else if (du_field.FieldId.Equals("renewEndDate") && blocks.ContainsKey("renewEndDate"))   
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["renewEndDate"][0].ToString());
                        if( clova_field.Formatted.FieldType == ClovaFieldType.Date)
                            resultsDataPoints.Add(CreateDateFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                        else
                            resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("issueDate") && blocks.ContainsKey("issueDate"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["issueDate"][0].ToString());
                        if( clova_field.Formatted.FieldType == ClovaFieldType.Date)
                            resultsDataPoints.Add(CreateDateFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                        else
                            resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("condition") && blocks.ContainsKey("condition"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["condition"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("code") && blocks.ContainsKey("code"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["code"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("organDonation") && blocks.ContainsKey("organDonation"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["organDonation"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("authority") && blocks.ContainsKey("authority"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["authority"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                }
            }
            else
            {
                Console.WriteLine($"error: {resp.body}");
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
            DateTimeOffset _date = clova_field.Formatted.Date;
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
            float ocr_confidence = 1.0f;
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
                ocr_confidence = Math.Min(w.OcrConfidence, ocr_confidence);
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
            return new ResultsValue(clova_field.Text, reference, 0.0f, (float)ocr_confidence); 
        }
    }
}
