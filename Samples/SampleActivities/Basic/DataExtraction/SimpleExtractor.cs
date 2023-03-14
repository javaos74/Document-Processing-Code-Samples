using Azure.AI.FormRecognizer.DocumentAnalysis;
using Azure;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
//using System.Diagnostics.Metrics;
using System.Linq;
using System.Threading.Tasks;
using UiPath.DocumentProcessing.Contracts;
using UiPath.DocumentProcessing.Contracts.DataExtraction;
using UiPath.DocumentProcessing.Contracts.Dom;
using UiPath.DocumentProcessing.Contracts.Results;
using UiPath.DocumentProcessing.Contracts.Taxonomy;
using System.IO;
using System.Drawing;

namespace SampleActivities.Basic.DataExtraction
{
    class PageLayout
    {
        public double Width { get; set; }
        public double Height { get; set; }
        public string Unit { get; set; }

        public PageLayout(double width, double height, string unit) {
            Width = width;
            Height = height;
            Unit = unit;
        }
    }

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

            fields.Add(new ExtractorFieldCapability { FieldId = "CustomerName", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "CustomerId", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "PurchaseOrder", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "PurchaseOrder", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "InvoiceId", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "InvoiceDate", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "VendorName", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "VendorTaxId", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "VendorAddress", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "VendorAddressRecipient", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "CustomerAddress", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "CustomerTaxId", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "CustomerAddressRecipient", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "BillingAddress", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "BillingAddressRecipient", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "ShippingAddress", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "ShippingAddressRecipient", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "PaymentTerm", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "SubTotal", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "TotalTax", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "InvoiceTotal", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "AmountDue", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "ServiceAddress", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "ServiceAddressRecipient", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "RemittanceAddress", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "RemittanceAddressRecipient", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "ServiceStartDate", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "ServiceEndDate", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "PreviousUnpaidBalance", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "CurrencyCode", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "PaymentOptions", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "TotalDiscount", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "Items", Components = new[] {
                new ExtractorFieldCapability {FieldId = "Amount", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                new ExtractorFieldCapability {FieldId = "Description", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                new ExtractorFieldCapability {FieldId = "Quantity", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                new ExtractorFieldCapability {FieldId = "UnitPrice", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                new ExtractorFieldCapability {FieldId = "ProductCode", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                new ExtractorFieldCapability {FieldId = "Unit", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                new ExtractorFieldCapability {FieldId = "Date", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                new ExtractorFieldCapability {FieldId = "Tax", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                new ExtractorFieldCapability {FieldId = "TaxRate", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                }, SetValues = new string[0] });
            return Task.FromResult( new[] { 
                new ExtractorDocumentTypeCapabilities{
                    DocumentTypeId = "azure.invoice.charlesdemo",
                    Fields = fields.ToArray()
                }
            });
            //return Task.FromResult(new ExtractorDocumentTypeCapabilities[0]);

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

            var task = new Task( _ => Execute(documentType, documentBounds, text, document, documentPath, endpoint, apiKey), state);
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

        protected void Execute(ExtractorDocumentType documentType, ResultsDocumentBounds documentBounds,
                                    string text, Document document, string documentPath,
                                    string endPoint, string apiKey)
        {

            this.result =  ComputeResult(documentType, documentBounds, text, document, documentPath, endPoint, apiKey);
        }


        private ExtractorResult ComputeResult(ExtractorDocumentType documentType, ResultsDocumentBounds documentBounds, 
                                string text, Document dom, string documentPath, string endpoint, string apiKey)
        {

            var credential = new AzureKeyCredential(apiKey);
            var client = new DocumentAnalysisClient(new Uri(endpoint), credential);
            var extractorResult = new ExtractorResult();
            var resultsDataPoints = new List<ResultsDataPoint>();

            
            AnalyzeDocumentOperation operation = client.AnalyzeDocument(WaitUntil.Completed, "prebuilt-invoice", File.OpenRead(documentPath));
            AnalyzeResult result = operation.Value;
            // call kakao ml extractor with documentPath using endPoint and apiKey 
            // create result data 
            foreach( var x in result.Pages)
            {
                this.pages.Add(new PageLayout((double)x.Width, (double)x.Height, x.Unit.ToString()));
            }

            for (int i = 0; i < result.Documents.Count; i++)
            {
                AnalyzedDocument azformdoc = result.Documents[i];
                foreach (var field in documentType.Fields)
                {
                    if ( azformdoc.Fields.TryGetValue(field.FieldId, out DocumentField outField))
                    {
                        if (field.Type == FieldType.Text)
                        {
                            resultsDataPoints.Add(CreateTextFieldDataPoint(field, outField, dom, pages.ToArray()));
                        }
                        else if (field.Type == FieldType.Date)
                        {
                            resultsDataPoints.Add(CreateDateFieldDataPoint(field, outField, dom, pages.ToArray()));
                        }
                        else if (field.Type == FieldType.Number)
                        {
                            resultsDataPoints.Add(CreateNumberFieldDataPoint(field, outField, dom, pages.ToArray()));
                        }
                        /*
                        else if( field.Type == FieldType.Table)
                        {
                            resultsDataPoints.Add(CreateTableFieldDataPoint(field, outField, dom));
                        }*/
                    }
                }
            }
            
            /*
            // example of reporting a value with derived parts
            Field firstDateField = documentType.Fields.FirstOrDefault(f => f.Type == FieldType.Text);
            if (firstDateField != null)
            {
                // if any field is of type Date return first word on first page as reference and report Jan 1st 2002 as value
                resultsDataPoints.Add(CreateTextFieldDataPoint(firstDateField, dom, "김형수 만세 "));
            }

            
            // example of report a value with no textual reference (only visual reference)
            Field firstBooleanField = documentType.Fields.FirstOrDefault(f => f.Type == FieldType.Boolean);
            if (firstBooleanField != null)
            {
                // if any field is of type Boolean return "true" with a visual reference from pixel position (50, 100) and width 200 and height 300.
                resultsDataPoints.Add(CreateBooleanFieldDataPoint(firstBooleanField, dom));
            }

            // example of table value
            Field firstTableField = documentType.Fields.FirstOrDefault(f => f.Type == FieldType.Table);
            if (firstTableField != null)
            {
                // if any field is of type Table return a table with headers referencing the first N words and 2 rows referencing the next N * 2 words.
                // N will be the number of columns in the table field.
                resultsDataPoints.Add(CreateTableFieldDataPoint(firstTableField, dom));
            }
            */


            extractorResult.DataPoints = resultsDataPoints.ToArray();
            return extractorResult;
        }

        private static ResultsDataPoint CreateTextFieldDataPoint(Field field, Document dom, PageLayout []pages, string value = null)
        {
            // how to map word index, now just default 0 
            var resultValue = CreateResultsValue2(0, dom, value);
            return new ResultsDataPoint(
                field.FieldId,
                field.FieldName,
                field.Type,
                new[] { resultValue });
        }

        private static ResultsDataPoint CreateTextFieldDataPoint(Field field, DocumentField mlfield, Document dom, PageLayout[] pages )
        {
            // how to map word index, now just default 0 
            var resultValue = CreateResultsValue(0, dom, mlfield, pages);
            return new ResultsDataPoint(
                field.FieldId,
                field.FieldName,
                field.Type,
                new[] { resultValue });
        }
        private static ResultsDataPoint CreateNumberFieldDataPoint(Field field, DocumentField mlfield , Document dom, PageLayout[] pages)
        {
            float _num  = mlfield.FieldType == DocumentFieldType.Double ? (float) mlfield.Value.AsDouble() : float.Parse(mlfield .Value.AsString());
            var derivedFields = ResultsDerivedField.CreateDerivedFieldsForNumber( _num);
            var resultValue = CreateResultsValue(0, dom, mlfield, pages);
            resultValue.DerivedFields = derivedFields;
            return new ResultsDataPoint(
                field.FieldId,
                field.FieldName,
                field.Type,
                new[] { resultValue });
        }
        private static ResultsDataPoint CreateDateFieldDataPoint(Field field, DocumentField mlfield, Document dom, PageLayout[] pages)
        {
            // TODO
            DateTimeOffset _date = mlfield.FieldType == DocumentFieldType.Date ? mlfield.Value.AsDate() : DateTimeOffset.Parse(mlfield.Content);
            var derivedFields = ResultsDerivedField.CreateDerivedFieldsForDate( _date.Day, _date.Month, _date.Year);
            var resultValue = CreateResultsValue(0, dom, mlfield, pages);
            resultValue.DerivedFields = derivedFields;

            return new ResultsDataPoint(
                field.FieldId,
                field.FieldName,
                field.Type,
                new[] { resultValue });
        }

        private static ResultsDataPoint CreateBooleanFieldDataPoint(Field field, DocumentField mlfield, Document dom)
        {
            var booleanToken = new ResultsValueTokens(0, (float)dom.Pages[0].Size.Width, (float)dom.Pages[0].Size.Height, new[] { Box.CreateChecked(50, 100, 200, 300) });
            var reference = new ResultsContentReference(0, 0, new[] { booleanToken });
            var firstBooleanValue = new ResultsValue("Yes", reference, 0.9f, 1f);

            return new ResultsDataPoint(
                field.FieldId,
                field.FieldName,
                field.Type,
                new[] { firstBooleanValue });
        }

        private static ResultsDataPoint CreateTableFieldDataPoint(Field firstTableField, DocumentField mlfield, Document dom, PageLayout[] pages )
        {
            int i = 0;
            var headerCells = firstTableField.Components.Select(c => new ResultsDataPoint(c.FieldId, c.FieldName, c.Type, new[] { CreateResultsValue(i++, dom, mlfield, pages) }));

            var firstRowCells = firstTableField.Components.Select(c => new ResultsDataPoint(c.FieldId, c.FieldName, c.Type, new[] { CreateResultsValue(i++, dom, mlfield, pages) }));
            var secondRowCells = firstTableField.Components.Select(c => new ResultsDataPoint(c.FieldId, c.FieldName, c.Type, new[] { CreateResultsValue(i++, dom, mlfield, pages) }));

            var tableValue = ResultsValue.CreateTableValue(firstTableField, headerCells, new[] { firstRowCells, secondRowCells }, 0.9f, 1f);

            return new ResultsDataPoint(
                firstTableField.FieldId,
                firstTableField.FieldName,
                firstTableField.Type,
                new[] { tableValue });
        }

        private static ResultsValue CreateResultsValue(int wordIndex, Document dom, DocumentField mlfield, PageLayout [] pages)
        {
            //TODO - find match word with BoundingBox 
            var word = dom.Pages[0].Sections.SelectMany(s => s.WordGroups).SelectMany(w => w.Words).ToArray()[wordIndex];
            List<Box> boxes = new List<Box>();
            List<ResultsValueTokens> tokens = new List<ResultsValueTokens>();
            foreach( var x in mlfield.BoundingRegions)
            {
                PointF[] points = x.BoundingPolygon.ToArray().Select( p =>  new PointF((float)ConvertSize(p.X, pages[x.PageNumber - 1].Width, dom.Pages[x.PageNumber-1].Size.Width),
                                                                                        (float)ConvertSize(p.Y, pages[x.PageNumber - 1].Height, dom.Pages[x.PageNumber - 1].Size.Height))).ToArray();
                //adjust co-ordinate 
                boxes.Add(Box.CreateUnchecked(points[3].X, points[3].Y, Math.Abs(points[0].X - points[3].X), Math.Abs(points[0].Y - points[1].Y)));
                tokens.Add(new ResultsValueTokens(word.IndexInText, word.Text.Length, 
                                x.PageNumber-1, (float)dom.Pages[ x.PageNumber-1].Size.Width, 
                                (float)dom.Pages[x.PageNumber-1].Size.Height, boxes.ToArray()));
                boxes.Clear();
            }
            //var wordValueToken = new ResultsValueTokens(word.IndexInText, word.Text.Length, 0, (float)dom.Pages[0].Size.Width, (float)dom.Pages[0].Size.Height, boxes.ToArray());
            var reference = new ResultsContentReference(word.IndexInText, word.Text.Length, tokens.ToArray());
            return new ResultsValue(mlfield.Content ?? word.Text, reference, (float)mlfield.Confidence, 0.0f); // word.OcrConfidence);
        }

        private static ResultsValue CreateResultsValue2(int wordIndex, Document dom, string value = null)
        {
            //TODO - find match word with BoundingBox 
            var word = dom.Pages[0].Sections.SelectMany(s => s.WordGroups).SelectMany(w => w.Words).ToArray()[wordIndex];
            var wordValueToken = new ResultsValueTokens(word.IndexInText, word.Text.Length, 0, (float)dom.Pages[0].Size.Width, (float)dom.Pages[0].Size.Height, new[] { word.Box });
            var reference = new ResultsContentReference(word.IndexInText, word.Text.Length, new[] { wordValueToken });
            
            return new ResultsValue(value ?? word.Text, reference, (float)0.9f, word.OcrConfidence);
        }
        private static double ConvertSize(double curX, double curWidth, double baseWidth)
        {
            return curX / curWidth * baseWidth; 
        }
    }
}
