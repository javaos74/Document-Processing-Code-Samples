using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Net;
using System.Net.Http;
using System.Globalization;
using System.IO;
using System.Activities;
using System.Net.Sockets;

namespace SampleActivities.Basic.OCR
{
    public class ClovaResponse
    {
        public HttpStatusCode status { get; set; }
        public string body { get; set; }
    }
    public class UpstageResponse
    {
        public HttpStatusCode status { get; set; }
        public string body { get; set; }
    }
    internal class ClovaSpeechParamBoosting
    {
        public ClovaSpeechParamBoosting( string w) {
            this.words = w;
        }

        public string words { get; set; }
    }
    internal class ClovaSpeechParam
    {
        public string language { get; set; } = "ko-KR";
        public string completion { get; set; } = "sync";
        public bool wordAlignment { get; set; } = false;
        public bool fullText { get; set; } = true;
        public bool resultToObs { get; set; } = false;
        public bool noiseFiltering { get; set; } = true;
        public ClovaSpeechParamBoosting[] boostings { get; set; }
        public bool useDomainBoostings { get; set; } = false;
        public string forbiddens { get; set; }
    }
    public class UiPathHttpClient
    {

        public UiPathHttpClient() :
            this("https://ailab.synap.co.kr")
        {
        }
        public UiPathHttpClient( string endpoint)
        {
            this.url = endpoint;
            this.client = new HttpClient();
            this.content = new MultipartFormDataContent("ocr----" + DateTime.Now.Ticks.ToString());
        }

        public void setEndpoint( string endpoint)
        {
            if (!string.IsNullOrEmpty(endpoint))
            {
                this.url = endpoint;
            }
        }
        public void setSecret( string secret)
        {
            setOCRSecret(secret);
        }
        public void setOCRSecret(string secret)
        {
            this.client.DefaultRequestHeaders.Add("X-OCR-SECRET", secret);
        }
        public void setSpeechSecret(string secret)
        {
            this.client.DefaultRequestHeaders.Add("X-CLOVASPEECH-API-KEY", secret);
        }
        public void setAuthorizationToken(string token)
        {
            this.client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
            this.client.DefaultRequestHeaders.Add("Accept", "*/*");
            //this.client.DefaultRequestHeaders.Add("User-Agent", "dotnet/1.0.0");
        }

        public void AddFile(string fileName, string fieldName = "file")
        {
            var fstream = System.IO.File.OpenRead(fileName);
#if DEBUG
            Console.WriteLine($"file size: {fstream.Length}");
#endif
            byte[] buf = new byte[fstream.Length];
            int read_bytes = 0;
            int offset = 0;
            int remains = (int)fstream.Length;
            do
            {
                read_bytes += fstream.Read(buf, offset, remains);
                offset += read_bytes;
                remains -= read_bytes;
            } while (remains != 0);
            fstream.Close();

            this.content.Add(new StreamContent(new MemoryStream(buf)), fieldName, System.IO.Path.GetFileName(fileName));
        }
 
        public void AddField( string name, string value)
        {
            this.content.Add(new StringContent(value), name);
        }

        public void Clear()
        {
            this.content.Dispose();
            this.content = new MultipartFormDataContent("ocr----" + DateTime.Now.Ticks.ToString());
        }

        public async Task<ClovaResponse> Upload()
        {
#if DEBUG
            Console.WriteLine("http content count :" + this.content.Count());
#endif
            using (var message = this.client.PostAsync(this.url, this.content))
            {
                ClovaResponse resp = new ClovaResponse();
                resp.status = message.Result.StatusCode;
                resp.body = await message.Result.Content.ReadAsStringAsync();
                return resp;
            }
        }

        public async Task<UpstageResponse> UploadUpstage()
        {
#if DEBUG
            Console.WriteLine("http content count :" + this.content.Count());
#endif
            using (var message = this.client.PostAsync(this.url, this.content))
            {
                UpstageResponse resp = new UpstageResponse();
                resp.status = message.Result.StatusCode;
                resp.body = await message.Result.Content.ReadAsStringAsync();
                return resp;
            }
        }


        private HttpClient client;
        private string url;
        private MultipartFormDataContent content;
    }
}
