using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using UiPath.Core;
using UiPath.DocumentProcessing.Contracts.DataExtraction;
using UiPath.DocumentProcessing.Contracts.Results;
using UiPath.DocumentProcessing.Contracts;
using UiPath.DocumentProcessing.Contracts.DataExtraction;
using UiPath.DocumentProcessing.Contracts.Dom;
using UiPath.DocumentProcessing.Contracts.Results;
using UiPath.DocumentProcessing.Contracts.Taxonomy;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using SampleActivities.Basic.OCR;
using static SampleActivities.Basic.OCR.ClovaOCRResultHelper;
using System.Text.Json;
using System.Drawing;

namespace SampleActivities.Basic.DataExtraction
{
    internal class KBSecField
    {
        [JsonProperty("id")]
        public int Id { get; set; } = 0;

        [JsonProperty("key")]
        public string Key { get; set; }

        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [JsonProperty("value")]
        public string Value { get; set; }

        [JsonProperty("refinedValue")]
        public string RefinedValue { get; set; }
        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("page")]
        public int[] Page { get; set; }

    }

    [DisplayName("Upstage KBSec IB Extractor")]
    public class KBSecIBDoc : ExtractorAsyncCodeActivity
    {

        [Category("Server")]
        [RequiredArgument]
        [Description("ML모델 서비스 endpoint 정보")]
        public InArgument<string> Endpoint { get; set; }

        [Category("Server")]
        [RequiredArgument]
        [Description("ML모델 서비스 endpoint Api Key정보 ")]
        public InArgument<string> ApiKey { get; set; }

        Object lockObj = new Object();

        ExtractorResult result;
        List<PageLayout> pages;


        public override Task<ExtractorDocumentTypeCapabilities[]> GetCapabilities()
        {
#if DEBUG
            Debug.WriteLine("GetCapabilities called");
#endif
            //Azure Form Recognizer invoice fields definition 
            List<ExtractorFieldCapability> fields = new List<ExtractorFieldCapability>();

            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.안건명", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.약어명", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.투자형태_구분1", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.투자형태_구분2", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.투자형태_구분3", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.금융상품분류", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.주요투자대상", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.투자기간(개월)", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.일수일정_증권신고서제출", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.기표일(예정)", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.인수일정_수요예측/청약", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.만기일(예정)", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.인수일정_증권발행일", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.부의기준통화", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.부의금액", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.투자국가", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.적용환율", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.부의금액(원)", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.전체수익", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.인수수수료율", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.수수료수익", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.실권수수료율", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.투자수익", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.채권최고금액", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.기타수수료율", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.고위험산업", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.투자분류_업무구분", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.사업지역", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.주요투자물건", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.투자물건상세", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.사업단계", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.담보여부", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.보증여부", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.책임준공여부", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.ESG리스크심사대상여부", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.ESG평가등급", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.업체명", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.협업유형", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "기초자산.기초자산종류", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "기초자산.기초자산명", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "기초자산.기초자산발행사법인번호", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "기초자산.부의기준통화", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "기초자산.기초자산평가액", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "기초자산.적용환율", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "기초자산.기초자산평가액(원)", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "관계사정보.관계사법인형태", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "관계사정보.법인번호", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "관계사정보.최대주주명", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "담보정보.담보유형", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "담보정보.담보상세", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "담보정보.담보명", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "담보정보.담보취득방식", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "담보정보.담보취득방식상세", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "담보정보.권리순위", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "담보정보.담보평가액", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "보증정보.신용보강보증구분", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "보증정보.보증금액", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "보증정보.보증내용", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "보증정보.법인번호", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "책임준공정보.책임준공기관구분", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "책임준공정보.미이행시의무", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "책임준공정보.법인번호", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "책임준공정보.책임준공기한", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "안건구조.인수일정_증권신고서제출", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });

            return Task.FromResult(new[] {
                new ExtractorDocumentTypeCapabilities{
                    DocumentTypeId = "upstage.kbsec.ibreport",
                    Fields = fields.ToArray()
                }
            });
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

        private async Task<ExtractorResult> ComputeResult(ExtractorDocumentType documentType, ResultsDocumentBounds documentBounds,
                                               string text, Document dom, string documentPath,
                                               string endpoint, string apikey)
        {
            var extractorResult = new ExtractorResult();
            var resultsDataPoints = new List<ResultsDataPoint>();

            var client = new UiPathHttpClient(endpoint);
            client.setAuthorizationToken(apikey);
            client.AddFile(documentPath, "document");
            foreach (var x in dom.Pages)
            {
                this.pages.Add(new PageLayout((double)x.Size.Width, (double)x.Size.Height, string.Empty));
            }
            var resp = await client.UploadUpstage();
            if (resp.status == HttpStatusCode.OK)
            {
                StringBuilder sb = new StringBuilder();
                JObject respJson = JObject.Parse(resp.body);
#if DEBUG2
                Console.WriteLine("IDcard result : " + resp.body);
#endif
                JArray blocks = (JArray)respJson["fields"];
                //name 
                foreach (var du_field in documentType.Fields)
                {
                    var found = blocks.Where( b => b["key"].ToString() == du_field.FieldId).ToList(); 
                    if( found.Count == 1 )
                    {
#if DEBUG2
                        Console.WriteLine( found.ElementAt(0).ToString() );
#endif
                        KBSecField kb_field = JsonConvert.DeserializeObject<KBSecField>(found.ElementAt(0).ToString());
                        if( !string.IsNullOrEmpty(kb_field.Value))
                            resultsDataPoints.Add( CreateTextFieldDataPoint(du_field, kb_field, dom, pages.ToArray()));
                    }
                }
            }
            extractorResult.DataPoints = resultsDataPoints.ToArray();
            return extractorResult;
        }

        private static ResultsDataPoint CreateTextFieldDataPoint(Field du_field, KBSecField kb_field, Document dom, PageLayout[] pages )
        {
            // how to map word index, now just default 0 
            var resultValue = CreateResultsValue(0, dom, kb_field, pages);
            return new ResultsDataPoint(
                du_field.FieldId,
                du_field.FieldName,
                du_field.Type,
                new[] { resultValue });
        }

        private static ResultsValue CreateResultsValue(int wordIndex, Document dom, KBSecField kb_field, PageLayout[] pages)
        {
            float ocr_confidence = 1.0f;
            int word_count = kb_field.Value.Split(" ".ToCharArray()).Length;
            List<Word> words = new List<Word>();
            foreach (var p in kb_field.Page)
            {
               /* var tmp = dom.Pages[p].Sections.SelectMany(s => s.WordGroups)
                    .SelectMany(w => w.Words).ToList(); */
                foreach( var sec in dom.Pages[p].Sections)
                {
                    foreach (var wg in sec.WordGroups)
                    {
                        if( wg.Words.Count() >= word_count && wg.Words.Any( w => kb_field.Value.Contains(w.Text))) // All보다 80% 이상 이렇게 해야 할것 같은데 ㅠㅠ 
                            words.AddRange(wg.Words);
                    }
                }
            }

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
            tokens.Add(new ResultsValueTokens(0, kb_field.Value.Length,
                                0,
                                (float)dom.Pages[0].Size.Width,
                                (float)dom.Pages[0].Size.Height, boxes.ToArray()));
            var reference = new ResultsContentReference(0, kb_field.Value.Length, tokens.ToArray());
            return new ResultsValue(kb_field.Value, reference, 0.0f, ocr_confidence);
        }

        public override Boolean ProvidesCapabilities()
        {
            return true;
        }
    }
}
