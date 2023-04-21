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

namespace SampleActivities.Basic.OCR
{
    public class ClovaOCRPair
    {
        public HttpStatusCode status { get; set; }
        public string body { get; set; }
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
            this.content = new MultipartFormDataContent("clova----" + DateTime.Now.ToString(CultureInfo.InvariantCulture));
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
            this.client.DefaultRequestHeaders.Add("X-OCR-SECRET", secret);
            //this.client.DefaultRequestHeaders.Add("Content-Type", "application/json");
        }

        public void AddFile(string fileName)
        {
            var fstream = System.IO.File.OpenRead(fileName);
            byte[] buf = new byte[fstream.Length];
            int read_bytes = 0;
            int offset = 0;
            int remains = (int)fstream.Length;
            do {
                read_bytes += fstream.Read(buf, offset, remains);
                offset += read_bytes;
                remains -= read_bytes;
            } while (remains != 0);
            fstream.Close();

            this.content.Add(new StreamContent(new MemoryStream(buf)), "file", System.IO.Path.GetFileNameWithoutExtension(fileName));
        }

        public void AddFile(string fileName, string fieldName)
        {
            var fstream = System.IO.File.OpenRead(fileName);
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
            this.content = new MultipartFormDataContent("clova----" + DateTime.Now.ToString(CultureInfo.InvariantCulture));
        }

        public async Task<ClovaOCRPair> Upload()
        {
#if DEBUG
            Console.WriteLine("http content count :" + this.content.Count());
#endif
            using (var message = this.client.PostAsync(this.url, this.content))
            {
                ClovaOCRPair resp = new ClovaOCRPair();
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
