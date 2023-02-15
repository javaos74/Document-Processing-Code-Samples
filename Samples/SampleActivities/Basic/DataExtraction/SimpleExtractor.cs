using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using UiPath.DocumentProcessing.Contracts;
using UiPath.DocumentProcessing.Contracts.DataExtraction;
using UiPath.DocumentProcessing.Contracts.Dom;
using UiPath.DocumentProcessing.Contracts.Results;
using UiPath.DocumentProcessing.Contracts.Taxonomy;

namespace SampleActivities.Basic.DataExtraction
{
    [DisplayName("Charles Extractor")]
    public class SimpleExtractor : ExtractorAsyncCodeActivity
    {
        [Category("Custom Model")]
        [RequiredArgument]
        [Description("ML모델 서비스 endpoint 정보")]
        public InArgument<string> Endpoint { get; set; }

        [Category("Custom Model")]
        [RequiredArgument]
        [Description("ML모델 서비스 endpoint Api Key정보 ")]
        public InArgument<string> ApiKey { get; set; }

        ExtractorResult result;

        public override Task<ExtractorDocumentTypeCapabilities[]> GetCapabilities()
        {
#if DEBUG
            Console.WriteLine("GetCapabilities called");
#endif
            // call kakao ml api for field description 
            // make extract document type capability with kakao ml fields info 

            List<ExtractorFieldCapability> fields = new List<ExtractorFieldCapability>();

            fields.Add(new ExtractorFieldCapability { FieldId = "m_name", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "m_age", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "m_birth-date", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "m_family", Components = new[] {
                new ExtractorFieldCapability {FieldId = "m_name", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                new ExtractorFieldCapability {FieldId = "m_relation", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                new ExtractorFieldCapability {FieldId = "m_birth-date", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                }, SetValues = new string[0] });
            return Task.FromResult( new[] { 
                new ExtractorDocumentTypeCapabilities{
                    DocumentTypeId = "charles.doc.type",
                    Fields = fields.ToArray()
                }
            });
            //return Task.FromResult(new ExtractorDocumentTypeCapabilities[0]);

        }
        public override Boolean ProvidesCapabilities()
        {
#if DEBUG
            Console.WriteLine("ProvidesCapabilities called");
#endif
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

            var task = new Task(_ => Execute(documentType, documentBounds, text, document, documentPath, endpoint, apiKey), state);
            task.Start();
            if (callback != null)
                task.ContinueWith(s => callback(s));
            return task;
        }

        protected override async void EndExecute(AsyncCodeActivityContext context, IAsyncResult result)
        {
            var task = (Task)result;
#if DEBUG
            Debug.WriteLine($"task.IsCompleted: {task.IsCompleted}");
#endif
            ExtractorResult.Set(context, this.result);
            await task;
        }


        protected async void Execute(ExtractorDocumentType documentType, ResultsDocumentBounds documentBounds,
                            string text, Document document, string documentPath,
                            string endPoint, string apiKey)
        {
#if DEBUG
            Debug.WriteLine("Executel called with parameters");
#endif
            this.result = await ComputeResult(documentType, documentBounds, text, document, documentPath, endPoint, apiKey);
        }

        private async Task<ExtractorResult> ComputeResult(ExtractorDocumentType documentType, ResultsDocumentBounds documentBounds, string text, Document document, string documentPath, string endpoint, string apiKey)
        {
            var extractorResult = new ExtractorResult();
            var resultsDataPoints = new List<ResultsDataPoint>();

            // call kakao ml extractor with documentPath 
            // match taxonomy fields with kakao ml field 
            // create result data 

            // example of reporting a value with derived parts
            Field firstDateField = documentType.Fields.FirstOrDefault(f => f.Type == FieldType.Date);
            if (firstDateField != null)
            {
                // if any field is of type Date return first word on first page as reference and report Jan 1st 2002 as value
                resultsDataPoints.Add(CreateDateFieldDataPoint(firstDateField, document));
            }

            // example of report a value with no textual reference (only visual reference)
            Field firstBooleanField = documentType.Fields.FirstOrDefault(f => f.Type == FieldType.Boolean);
            if (firstBooleanField != null)
            {
                // if any field is of type Boolean return "true" with a visual reference from pixel position (50, 100) and width 200 and height 300.
                resultsDataPoints.Add(CreateBooleanFieldDataPoint(firstBooleanField, document));
            }

            // example of table value
            Field firstTableField = documentType.Fields.FirstOrDefault(f => f.Type == FieldType.Table);
            if (firstTableField != null)
            {
                // if any field is of type Table return a table with headers referencing the first N words and 2 rows referencing the next N * 2 words.
                // N will be the number of columns in the table field.
                resultsDataPoints.Add(CreateTableFieldDataPoint(firstTableField, document));
            }

            extractorResult.DataPoints = resultsDataPoints.ToArray();
            return extractorResult;
        }

        private static ResultsDataPoint CreateDateFieldDataPoint(Field firstDateField, Document document)
        {
            // TODO
            var derivedFields = ResultsDerivedField.CreateDerivedFieldsForDate(1, 1, 2002);
            var firstDateValue = CreateResultsValue(0, document, "Jan 1st 2002");
            firstDateValue.DerivedFields = derivedFields;

            return new ResultsDataPoint(
                firstDateField.FieldId,
                firstDateField.FieldName,
                firstDateField.Type,
                new[] { firstDateValue });
        }

        private static ResultsDataPoint CreateBooleanFieldDataPoint(Field firstBooleanField, Document document)
        {
            var booleanToken = new ResultsValueTokens(0, (float)document.Pages[0].Size.Width, (float)document.Pages[0].Size.Height, new[] { Box.CreateChecked(50, 100, 200, 300) });
            var reference = new ResultsContentReference(0, 0, new[] { booleanToken });
            var firstBooleanValue = new ResultsValue("Yes", reference, 0.9f, 1f);

            return new ResultsDataPoint(
                firstBooleanField.FieldId,
                firstBooleanField.FieldName,
                firstBooleanField.Type,
                new[] { firstBooleanValue });
        }

        private static ResultsDataPoint CreateTableFieldDataPoint(Field firstTableField, Document document)
        {
            int i = 0;
            var headerCells = firstTableField.Components.Select(c => new ResultsDataPoint(c.FieldId, c.FieldName, c.Type, new[] { CreateResultsValue(i++, document) }));

            var firstRowCells = firstTableField.Components.Select(c => new ResultsDataPoint(c.FieldId, c.FieldName, c.Type, new[] { CreateResultsValue(i++, document) }));
            var secondRowCells = firstTableField.Components.Select(c => new ResultsDataPoint(c.FieldId, c.FieldName, c.Type, new[] { CreateResultsValue(i++, document) }));

            var tableValue = ResultsValue.CreateTableValue(firstTableField, headerCells, new[] { firstRowCells, secondRowCells }, 0.9f, 1f);

            return new ResultsDataPoint(
                firstTableField.FieldId,
                firstTableField.FieldName,
                firstTableField.Type,
                new[] { tableValue });
        }

        private static ResultsValue CreateResultsValue(int wordIndex, Document document, string value = null)
        {
            var word = document.Pages[0].Sections.SelectMany(s => s.WordGroups).SelectMany(w => w.Words).ToArray()[wordIndex];
            var wordValueToken = new ResultsValueTokens(word.IndexInText, word.Text.Length, 0, (float)document.Pages[0].Size.Width, (float)document.Pages[0].Size.Height, new[] { word.Box });
            var reference = new ResultsContentReference(word.IndexInText, word.Text.Length, new[] { wordValueToken });

            return new ResultsValue(value ?? word.Text, reference, 0.9f, word.OcrConfidence);
        }
    }
}
