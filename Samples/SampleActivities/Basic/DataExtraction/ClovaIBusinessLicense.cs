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

   
    [DisplayName("Clova 사업자등록증 Extractor")]
    public class ClovaBusinessLicenseExtractor : ExtractorAsyncCodeActivity
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

            fields.Add(new ExtractorFieldCapability { FieldId = "birth", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "bisAddress", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "bisArea", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "bisItem", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "registerNumber", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "bisType", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "companyName", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "coRepName", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "corpName", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "corpRegisterNum", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "coRepSocialNum", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "documentType", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "headAddress", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "issuanceDate", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "issuanceReason", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "openDate", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "repName", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "socialNumber", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "taxType", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });

            return Task.FromResult(new[] {
                new ExtractorDocumentTypeCapabilities{
                    DocumentTypeId = "clova.bizLicense",
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
                JObject blocks = (JObject)respJson["images"][0]["bizLicense"]["result"];
#if DEBUG2
                Console.WriteLine("driverlicense result : " + resp.body);
#endif
                //name 
                foreach (var du_field in documentType.Fields)
                {
#if DEBUG2
                    Console.WriteLine($"field: {du_field.FieldId}");
#endif
                    if (du_field.FieldId.Equals("birth") && blocks.ContainsKey("birth"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["birth"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("bisAddress") && blocks.ContainsKey("bisAddress"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["bisAddress"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("bisArea") && blocks.ContainsKey("bisArea"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["bisArea"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("bisItem") && blocks.ContainsKey("bisItem"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["bisItem"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("registerNumber") && blocks.ContainsKey("registerNumber"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["registerNumber"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("bisType") && blocks.ContainsKey("bisType"))   
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["bisType"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("companyName") && blocks.ContainsKey("companyName"))   
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["companyName"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("coRepName") && blocks.ContainsKey("coRepName"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["coRepName"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("corpName") && blocks.ContainsKey("corpName"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["corpName"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("corpRegisterNum") && blocks.ContainsKey("corpRegisterNum"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["corpRegisterNum"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("coRepSocialNum") && blocks.ContainsKey("coRepSocialNum"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["coRepSocialNum"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("documentType") && blocks.ContainsKey("documentType"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["documentType"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("headAddress") && blocks.ContainsKey("headAddress"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["headAddress"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("issuanceDate") && blocks.ContainsKey("issuanceDate"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["issuanceDate"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("issuanceReason") && blocks.ContainsKey("issuanceReason"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["issuanceReason"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("openDate") && blocks.ContainsKey("openDate"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["openDate"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("repName") && blocks.ContainsKey("repName"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["repName"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("socialNumber") && blocks.ContainsKey("socialNumber"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["socialNumber"][0].ToString());
                        resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, clova_field, dom, pages.ToArray()));
                    }
                    else if (du_field.FieldId.Equals("taxType") && blocks.ContainsKey("taxType"))
                    {
                        ClovaField clova_field = JsonConvert.DeserializeObject<ClovaField>(blocks["taxType"][0].ToString());
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
