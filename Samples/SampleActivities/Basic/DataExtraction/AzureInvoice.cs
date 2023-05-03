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
using Microsoft.Extensions.Azure;

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

    [DisplayName("Charles Azure-Invoice Extractor")]
    public class AzureInvoice : ExtractorAsyncCodeActivity
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

            fields.Add(new ExtractorFieldCapability { FieldId = "CustomerName", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "CustomerId", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
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
                foreach (var du_field in documentType.Fields)
                {
                    if ( azformdoc.Fields.TryGetValue(du_field.FieldId, out DocumentField out_az_field))
                    {
                        if (du_field.Type == FieldType.Text)
                        {
                            resultsDataPoints.Add(CreateTextFieldDataPoint(du_field, out_az_field, dom, pages.ToArray()));
#if DEBUG
                            Console.WriteLine(du_field.FieldName + ": " + Newtonsoft.Json.JsonConvert.SerializeObject(out_az_field));
#endif
                        }
                        else if (du_field.Type == FieldType.Date)
                        {
                            resultsDataPoints.Add(CreateDateFieldDataPoint(du_field, out_az_field, dom, pages.ToArray()));
#if DEBUG
                            Console.WriteLine(Newtonsoft.Json.JsonConvert.SerializeObject(out_az_field));
#endif
                        }
                        else if (du_field.Type == FieldType.Number)
                        {
                            resultsDataPoints.Add(CreateNumberFieldDataPoint(du_field, out_az_field, dom, pages.ToArray()));
#if DEBUG
                            Console.WriteLine(Newtonsoft.Json.JsonConvert.SerializeObject(out_az_field));
#endif
                        }
                        
                        else if( du_field.Type == FieldType.Table)
                        {
                            resultsDataPoints.Add(CreateTableFieldDataPoint(du_field, out_az_field, dom, pages.ToArray()));
                        }
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
        private static ResultsDataPoint CreateTableFieldDataPoint(Field du_field, DocumentField az_field, Document dom, PageLayout[] pages)
        {
            int i = 0;
            float confidence = 1.0f;
            List<ResultsDataPoint> dataPoints = new List<ResultsDataPoint>();
            List<IEnumerable<ResultsDataPoint>> rows = new List<IEnumerable<ResultsDataPoint>>();
            if (az_field.FieldType == DocumentFieldType.List)
            {
                foreach (DocumentField az_item in az_field.Value.AsList())
                {
                    if (az_item.FieldType == DocumentFieldType.Dictionary)
                    {
                        IReadOnlyDictionary<string, DocumentField> az_item_dictionary = az_item.Value.AsDictionary();

                        var row = du_field.Components.Select(c => new ResultsDataPoint(c.FieldId, c.FieldName, c.Type,
                                                                    new[] { CreateRowResultsValue(i++, c, dom, az_item_dictionary, pages.ToArray()) }));
                        confidence = Math.Min(confidence, (float)az_item.Confidence);
                        rows.Add(row);
                    }
                }
            }
            //az_field.Value.AsList()[j]
            foreach(DocumentField az_item in az_field.Value.AsList()) { 
                var headerCells = du_field.Components.Select(c => new ResultsDataPoint(c.FieldId, c.FieldName, c.Type, new[] { CreateResultsValue(i++, dom, az_item, pages) }));
                dataPoints.AddRange(headerCells);
            }

            var tableValue = ResultsValue.CreateTableValue(du_field, dataPoints, rows.ToArray(), confidence, 0.0f);
            return new ResultsDataPoint(
                du_field.FieldId,
                du_field.FieldName,
                du_field.Type,
                new[] { tableValue });
        }



        private static ResultsDataPoint CreateTextFieldDataPoint(Field du_field, DocumentField az_field, Document dom, PageLayout[] pages )
        {
            // how to map word index, now just default 0 
            var resultValue = CreateResultsValue(0, dom, az_field, pages);
            return new ResultsDataPoint(
                du_field.FieldId,
                du_field.FieldName,
                du_field.Type,
                new[] { resultValue });
        }
        private static ResultsDataPoint CreateNumberFieldDataPoint(Field du_field, DocumentField az_field , Document dom, PageLayout[] pages)
        {
            float _num  = az_field.FieldType == DocumentFieldType.Double ? (float) az_field.Value.AsDouble() : float.Parse(az_field .Value.AsString());
            var derivedFields = ResultsDerivedField.CreateDerivedFieldsForNumber( _num);
            var resultValue = CreateResultsValue(0, dom, az_field, pages);
            resultValue.DerivedFields = derivedFields;
            return new ResultsDataPoint(
                du_field.FieldId,
                du_field.FieldName,
                du_field.Type,
                new[] { resultValue });
        }
        private static ResultsDataPoint CreateDateFieldDataPoint(Field du_field, DocumentField az_field, Document dom, PageLayout[] pages)
        {
            // TODO
            DateTimeOffset _date = az_field.FieldType == DocumentFieldType.Date ? az_field.Value.AsDate() : DateTimeOffset.Parse(az_field.Content);
            var derivedFields = ResultsDerivedField.CreateDerivedFieldsForDate( _date.Day, _date.Month, _date.Year);
            var resultValue = CreateResultsValue(0, dom, az_field, pages);
            resultValue.DerivedFields = derivedFields;

            return new ResultsDataPoint(
                du_field.FieldId,
                du_field.FieldName,
                du_field.Type,
                new[] { resultValue });
        }

        private static ResultsDataPoint CreateBooleanFieldDataPoint(Field du_field, DocumentField az_field, Document dom)
        {
            var booleanToken = new ResultsValueTokens(0, (float)dom.Pages[0].Size.Width, (float)dom.Pages[0].Size.Height, new[] { Box.CreateChecked(50, 100, 200, 300) });
            var reference = new ResultsContentReference(0, 0, new[] { booleanToken });
            var firstBooleanValue = new ResultsValue("Yes", reference, 0.9f, 1f);

            return new ResultsDataPoint(
                du_field.FieldId,
                du_field.FieldName,
                du_field.Type,
                new[] { firstBooleanValue });
        }

       

        private static ResultsValue CreateRowResultsValue( int  rowIdx, Field du_item, Document dom, IReadOnlyDictionary<string, DocumentField> az_item_dictionary, PageLayout[] pages)
        {
            if (!az_item_dictionary.TryGetValue(du_item.FieldId, out DocumentField az_item))
                return null;

            float ocr_confidence = 1.0f;
            Rectangle rect = new Rectangle((Int32)az_item.BoundingRegions[0].BoundingPolygon[0].X - 2,
                      (Int32)az_item.BoundingRegions[0].BoundingPolygon[0].Y - 2,
                      (Int32)(Math.Abs(az_item.BoundingRegions[0].BoundingPolygon[1].X - az_item.BoundingRegions[0].BoundingPolygon[0].X) * 1.1),
                      (Int32)(Math.Abs(az_item.BoundingRegions[0].BoundingPolygon[2].Y - az_item.BoundingRegions[0].BoundingPolygon[0].Y) * 1.1));

            var words = dom.Pages[az_item.BoundingRegions[0].PageNumber - 1].Sections.SelectMany(s => s.WordGroups)
                .SelectMany(w => w.Words).Where(t => rect.Contains(new Rectangle((Int32)t.Box.Left, (Int32)t.Box.Top, (Int32)t.Box.Width, (Int32)t.Box.Height))).ToArray();
#if DEBUG
            Console.WriteLine($"{words.Length} words found");
#endif
            List<Box> boxes = new List<Box>();
            List<ResultsValueTokens> tokens = new List<ResultsValueTokens>();
            foreach (var w in words)
            {
                boxes.Add(w.Box);
#if DEBUG
                Console.WriteLine($"Box : {w.Box.Left}, {w.Box.Top}, {w.Box.Width}, {w.Box.Height}");
#endif
                ocr_confidence = Math.Min(ocr_confidence, w.OcrConfidence);
            }
#if DEBUG
            Console.WriteLine($"{boxes.Count} is found");
#endif
            if (boxes.Count == 0)
            {
                boxes.Add(Box.CreateChecked(0, 0, 0, 0));
            }
            tokens.Add(new ResultsValueTokens(az_item.Spans[0].Index, az_item.Spans[0].Length,
                                az_item.BoundingRegions[0].PageNumber - 1,
                                (float)dom.Pages[az_item.BoundingRegions[0].PageNumber - 1].Size.Width,
                                (float)dom.Pages[az_item.BoundingRegions[0].PageNumber - 1].Size.Height, boxes.ToArray()));
            var reference = new ResultsContentReference(az_item.Spans[0].Index, az_item.Spans[0].Length, tokens.ToArray());

            ResultsValue rv = new ResultsValue(az_item.Content, reference, (float)az_item.Confidence, (float)ocr_confidence);

            if(du_item.Type == FieldType.Number)
            {
                float _num = az_item.FieldType == DocumentFieldType.Double ? (float)az_item.Value.AsDouble() : float.Parse(az_item.Value.AsString());
                var derivedFields = ResultsDerivedField.CreateDerivedFieldsForNumber(_num);
                rv.DerivedFields = derivedFields;
            }

            return rv;
        }

        private static ResultsValue CreateResultsValue(int wordIndex, Document dom, DocumentField az_field, PageLayout[] pages)
        {
            float ocr_confidence = 1.0f;
            Rectangle rect;
            rect = new Rectangle((Int32)az_field.BoundingRegions[0].BoundingPolygon[0].X - 2,
                              (Int32)az_field.BoundingRegions[0].BoundingPolygon[0].Y - 2,
                              (Int32)(Math.Abs(az_field.BoundingRegions[0].BoundingPolygon[1].X - az_field.BoundingRegions[0].BoundingPolygon[0].X) * 1.1),
                              (Int32)(Math.Abs(az_field.BoundingRegions[0].BoundingPolygon[2].Y - az_field.BoundingRegions[0].BoundingPolygon[0].Y) * 1.1));

            var words = dom.Pages[az_field.BoundingRegions[0].PageNumber - 1].Sections.SelectMany(s => s.WordGroups)
                .SelectMany(w => w.Words).Where(t => rect.Contains(new Rectangle((Int32)t.Box.Left, (Int32)t.Box.Top, (Int32)t.Box.Width, (Int32)t.Box.Height))).ToArray();

#if DEBUG
            Console.WriteLine($"{words.Length} words found");
#endif
            List<Box> boxes = new List<Box>();
            List<ResultsValueTokens> tokens = new List<ResultsValueTokens>();
            foreach (var w in words)
            {
                boxes.Add(w.Box);
#if DEBUG
                Console.WriteLine($"Box : {w.Box.Left}, {w.Box.Top}, {w.Box.Width}, {w.Box.Height}");
#endif
                ocr_confidence = Math.Min( ocr_confidence, w.OcrConfidence);
            }
#if DEBUG
            Console.WriteLine($"{boxes.Count} is found");
#endif
            if (boxes.Count == 0)
            {
                boxes.Add(Box.CreateChecked(0, 0, 0, 0));
            }
            tokens.Add(new ResultsValueTokens(az_field.Spans[0].Index, az_field.Spans[0].Length,
                                az_field.BoundingRegions[0].PageNumber - 1,
                                (float)dom.Pages[az_field.BoundingRegions[0].PageNumber - 1].Size.Width,
                                (float)dom.Pages[az_field.BoundingRegions[0].PageNumber - 1].Size.Height, boxes.ToArray()));
            var reference = new ResultsContentReference(az_field.Spans[0].Index, az_field.Spans[0].Length, tokens.ToArray());
            return new ResultsValue(az_field.Content, reference, (float)az_field.Confidence, (float)ocr_confidence / boxes.Count); // word.OcrConfidence);
        }

        private static double ConvertSize(double curX, double curWidth, double baseWidth)
        {
            return curX / curWidth * baseWidth; 
        }
    }
}
