using Newtonsoft.Json.Linq;
using SampleActivities.Basic.OCR;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using UiPath.DocumentProcessing.Contracts.DataExtraction;
using UiPath.DocumentProcessing.Contracts.Dom;
using UiPath.DocumentProcessing.Contracts.Taxonomy;
using UiPath.OCR.Contracts;
using UiPath.OCR.Contracts.Activities;
using UiPath.OCR.Contracts.DataContracts;
using static System.Net.Mime.MediaTypeNames;

namespace SampleActivities.Basic.Speech
{
    [DisplayName("Clova Speech Engine")]
    public class ClovaSpeech: AsyncCodeActivity
    {

        [Category("Login")]
        [RequiredArgument]
        [Description("Clova Speech API endpoint 정보")]
        public InArgument<string> Endpoint { get; set; }

        [Category("Login")]
        [RequiredArgument]
        [Description("Clova Speech Secret")]
        public InArgument<string> Secret { get; set; }

        [Category("Input")]
        [Browsable(false)]
        public InArgument<string> Language { get; set; } = "ko-KR";

        [Category("Input")]
        [RequiredArgument]
        [Browsable(true)]
        public InArgument<string> VoiceFilePath { get; set; }

        [Category("Input")]
        [Browsable(false)]
        public InArgument<string> Completion { get; set; } = "sync";


        [Category("Input")]
        [Browsable(true)]
        public InArgument<string[]> Boostings { get; set; }

        [Category("Output")]
        [Browsable(true)]
        public OutArgument<string>  FullText { get; set; }

        private UiPathHttpClient    _httpClient;

        private ClovaResponse _result;
        private string _fullText;

        public ClovaSpeech()
        {
            _httpClient = new UiPathHttpClient();
        }

        private async void Execute( string endpoint, string secret, string lang, string filePath,
                                    string completion, string[] boostings)
        {
            StringBuilder sb = new StringBuilder();
            ClovaSpeechParam param = new ClovaSpeechParam();

            param.language = lang;
            param.completion = completion;
            if (boostings != null)
            {
                List<ClovaSpeechParamBoosting> words = new List<ClovaSpeechParamBoosting>();
                foreach (var word in boostings)
                {
                    words.Add(new ClovaSpeechParamBoosting(word));
                }
                param.boostings = words.ToArray();
            }
            _httpClient.setEndpoint(endpoint);
            _httpClient.setSpeechSecret(secret);
            _httpClient.AddFile(filePath, "media");
            _httpClient.AddField("params", Newtonsoft.Json.JsonConvert.SerializeObject(param));
            _httpClient.AddField("type", "application/json");

            this._result = await _httpClient.Upload();
        }

        protected override IAsyncResult BeginExecute(AsyncCodeActivityContext context, AsyncCallback callback, object state)
        {
            var endpoint = Endpoint.Get(context);
            var secret = Secret.Get(context);
            var lang = Language.Get(context);
            var filePath = VoiceFilePath.Get(context);
            var completion = Completion.Get(context);
            string[] boostings = Boostings.Get(context);

            var task = new Task(_ => Execute(endpoint, secret, lang, filePath,completion, boostings), state);
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
            if( this._result.status == HttpStatusCode.OK )
            {
                StringBuilder sb = new StringBuilder(); 
                JObject respJson = JObject.Parse(this._result.body);
                JArray segments = (JArray)respJson["segments"];
                foreach (var segment in segments)
                {
                    sb.Append(segment["textEdited"].ToString() + "\n");
                }
                this._fullText = sb.ToString();
            }

            FullText.Set(context, this._fullText);
            await task;
        }

    }
}
