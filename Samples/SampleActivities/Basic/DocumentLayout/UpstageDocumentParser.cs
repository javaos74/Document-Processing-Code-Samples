using Newtonsoft.Json.Linq;
using SampleActivities.Basic.OCR;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using UiPath.Shared.Activities.Services;
using TextTableFormatter;

namespace SampleActivities.Basic.DocumentLayout
{
    [DisplayName("Upstage Document Parser")]
    [Browsable(true)]
    public class UpstageDocumentParser : AsyncCodeActivity
    {
        [Category("Login")]
        [RequiredArgument]
        [Description("Upstage Document Parser API endpoint 정보")]
        public InArgument<string> Endpoint { get; set; }

        [Category("Login")]
        [RequiredArgument]
        [Description("Upstage Document Parser ApiKey")]
        public InArgument<string> ApiKey { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Browsable(true)]
        public InArgument<string> InputFilePath { get; set; }

        [Category("Options")]
        [Description("Upstage Document Parser ApiKey")]
        public InArgument<string> OCR { get; set; } = "auto";

        [Category("Output")]
        [Browsable(true)]
        public OutArgument<string> FullText { get; set; }

        private UiPathHttpClient _httpClient;

        private ClovaResponse _result;
        private string _fullText;

        public UpstageDocumentParser()
        {
            _httpClient = new UiPathHttpClient();
        }

        private async void Execute(string endpoint, string apikey, string ocr, string filePath )
        {
            string[] formats = { "markdown", "text" };
            StringBuilder sb = new StringBuilder();
            ClovaSpeechParam param = new ClovaSpeechParam();

            _httpClient.setEndpoint(endpoint);
            _httpClient.setAuthorizationToken(apikey);
            _httpClient.AddFile(filePath, "document");
            _httpClient.AddField("output_formats", Newtonsoft.Json.JsonConvert.SerializeObject( formats));
            _httpClient.AddField("ocr", ocr);

            this._result = await _httpClient.Upload();
        }
        protected override IAsyncResult BeginExecute(AsyncCodeActivityContext context, AsyncCallback callback, object state)
        {
            var endpoint = Endpoint.Get(context);
            var secret = ApiKey.Get(context);
            var ocr = OCR.Get(context);
            var filePath = InputFilePath.Get(context);

            var task = new Task(_ => Execute(endpoint, secret, ocr, filePath ), state);
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

#if DEBUG
            Console.WriteLine($"status={this._result.status}, body={this._result.body}");
#endif
            if (this._result.status == HttpStatusCode.OK)
            {
                StringBuilder sb = new StringBuilder();
                JObject respJson = JObject.Parse(this._result.body);
                JArray els= (JArray)respJson["elements"];
                foreach (var el in els)
                {
                    if (el["category"].ToString() != "table")
                        sb.Append(el["content"]["text"].ToString() + "\n");
                    else
                    {
                        var lines = el["content"]["markdown"].ToString().Split( "\n".ToCharArray());
                        var txttbl = new TextTable(lines[0].Split("|".ToCharArray()).Length);
                        int idx = 0;
                        foreach ( var line in lines)
                        {
                            if (idx != 1)
                            {
                                var cells = line.Split("|".ToCharArray());
                                foreach (var cell in cells)
                                {
                                    txttbl.AddCell(cell);
                                }
                            }
                            idx++;
                        }
                        sb.Append( txttbl.Render() +"\n");
                        //sb.Append(el["content"]["markdown"].ToString() + "\n");
                    }
                }
                this._fullText = sb.ToString();
            }

            FullText.Set(context, this._fullText);
            await task;
        }
    }
}
