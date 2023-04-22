using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using UiPath.OCR.Contracts.DataContracts;
using UiPath.OCR.Contracts;
//using BitMiracle.Docotic.Pdf;

namespace SampleActivities.Basic.OCR
{
    internal static class HancomOCRResultHelper
    {
        internal class RequestBody
        {
            public string key { get; set; }
            public string request_id { get; set; } = Guid.NewGuid().ToString();
            public string file_format { get; set; } = "image";
            public string file_url { get; set; } = String.Empty;
            public string file_bytes { get; set; } = String.Empty;
            public string file_upload { get; set; } = String.Empty;

            public override string ToString()
            {
                return $"request_id={request_id}, file_formt={file_format}";
            }
        }

        internal static  UiPath.OCR.Contracts.OCRRotation GetOCRRotation( Single rot)
        {
 #if DEBUG2
            System.Console.WriteLine(" roation : " + rot);
 #endif
            if ( rot >= 45 && rot < 90+45)
                return OCRRotation.Rotated90;
            else if ( rot >= 90+45 && rot < 180+45)
                return OCRRotation.Rotated180;
            else if( rot >= 180+45 && rot < 270+45)
                return OCRRotation.Rotated270;
            else if ( rot >= 270+45 || rot < 45)
                return OCRRotation.None;
            else
                return OCRRotation.Other;
        }


        internal static async Task<OCRResult> FromHancomClient(string file_path, Dictionary<string, object> options)
        {
            OCRResult ocrResult = new OCRResult();
            var client = new UiPathHttpClient(options["endpoint"].ToString());

            client.AddField("key", options["apikey"].ToString());
            client.AddField("file_format", "image");
            client.AddField("request_id", Guid.NewGuid().ToString());
            client.AddField("file_url", string.Empty);
            client.AddField("file_bytes", string.Empty);
            client.AddFile(file_path, "file_upload");
#if DEBUG
            Console.WriteLine($"key: {options["apikey"].ToString()}");
#endif
            var resp = await client.Upload();
#if DEBUG
            System.Console.WriteLine(resp.status + " == > " + (resp.body.Length > 100 ? resp.body.Substring(0, 100) : resp.body));
            System.IO.File.WriteAllText(@"C:\Temp\hancom_resp.json", resp.body);
#endif
            if (resp.status == HttpStatusCode.OK)
            {
                StringBuilder sb = new StringBuilder();
                JObject respJson = JObject.Parse(resp.body);
                JArray blocks = (JArray)respJson["content"]["ocr_data"][0]["words"];
                ocrResult.Words = blocks.Select(p => new Word
                {
                    Text = (string)p["text"],
                    Confidence =  Convert.ToInt32(100 * (Convert.ToDouble(p["score"]))),
                    PolygonPoints = new [] { 
                                        new PointF( (float)p["bbox"][0], (float)p["bbox"][1]),
                                        new PointF( (float)p["bbox"][2], (float)p["bbox"][3]),
                                        new PointF( (float)p["bbox"][4], (float)p["bbox"][5]),
                                        new PointF( (float)p["bbox"][6], (float)p["bbox"][7])
                                    },
                    Characters = ((string)p["text"]).Select(ch => new Character
                    {
                        Char = ch,
                    }).ToArray()
                }).ToArray();
                foreach (var word in ocrResult.Words)
                {
                    var x = word.PolygonPoints[0].X;
                    var y = word.PolygonPoints[0].Y;
                    var w = Math.Abs(word.PolygonPoints[1].X - x);
                    var y2 = word.PolygonPoints[3].Y;

                    float dx = w / word.Characters.Length;
                    float dy = Math.Abs(y2 - y) / word.Characters.Length;
                    int idx = 0;
#if DEBUG
                   // System.Console.WriteLine(string.Format("{0} has {1} characters", word.Text, word.Characters.Length));
#endif
                    foreach (var c in word.Characters)
                    {
                        c.PolygonPoints = new[] { new PointF(x + dx * idx, y), new PointF(x + dx * (idx + 1), y), new PointF(x + dx * (idx + 1), y2), new PointF(x + dx * idx, y2) };
                        c.Confidence = word.Confidence;
                        c.Rotation = GetOCRRotation(360 - (Single)respJson["content"]["ocr_data"][0]["image_rotation"]);
                        idx++;
                    }
                }
                ocrResult.Text = respJson["content"]["ocr_data"][0]["page_text"].ToString();
                ocrResult.SkewAngle = 0;
                ocrResult.Confidence = 0;
            }
            return ocrResult;
        }
    }
}
