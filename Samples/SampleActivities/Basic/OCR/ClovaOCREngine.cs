using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using UiPath.OCR.Contracts;
using UiPath.OCR.Contracts.Activities;
using UiPath.OCR.Contracts.DataContracts;

namespace SampleActivities.Basic.OCR
{
    [DisplayName("Clova OCR Engine")]
    public class ClovaOCREngine : OCRCodeActivity
    {
        [Category("Input")]
        [Browsable(true)]
        public override InArgument<Image> Image { get => base.Image; set => base.Image = value; }

        [Category("Login")]
        [RequiredArgument]
        [Description("Clova OCR API GW endpoint 정보")]
        public InArgument<string> Endpoint { get; set; }

        [Category("Login")]
        [RequiredArgument]
        [Description("Clova OCR Secret")]
        public InArgument<string> Secret { get; set; }

        [Category("Option")]
        [Browsable(true)]
        [Description("설정 가능한 언어값은 ko ja zh-TW 중 하나를 선택")]
        public InArgument<string> Languages { get; set; } = "ko";


        [Category("Output")]
        [Browsable(true)]
        public override OutArgument<string> Text { get => base.Text; set => base.Text = value; }


        private string file_path;


        /**
         * OCRENgine으로 동작하는데 필요한 함수 구현 
         * Dictionary<string,object> options에 필요한 값을 담아서 넘겨준다. 
         */
        public override Task<OCRResult> PerformOCRAsync(Image image, Dictionary<string, object> options, CancellationToken ct)
        {

            file_path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "clova_ocr_req_image.png");
            if( image != null ) {
                if (System.IO.File.Exists(file_path))
                    System.IO.File.Delete(file_path);
#if DEBUG
                System.Console.WriteLine($"width={image.Width}, height={image.Height} resolution={image.HorizontalResolution} ");
#endif
                image.Save(file_path, System.Drawing.Imaging.ImageFormat.Png);
            } else
            {
                file_path = string.Empty;
            }
 #if DEBUG
            System.Console.WriteLine("temp file path " + file_path);
#endif

            var result =   ClovaOCRResultHelper.FromClovaClient(file_path, options);

            return result;
        }

        /**
         * Output 출력을 설정한다. PeformOCRAsync에서 options에 담겨진 값을 이용해서 최종 Output argument에 값을 설정한다. 
         */
        protected override void OnSuccess(CodeActivityContext context, OCRResult result)
        {

;       }
        //protected override void on

        protected override Dictionary<string, object> BeforeExecute(CodeActivityContext context)
        {
            return new Dictionary<string, object>
            {
                { "endpoint",  Endpoint.Get(context) },
                { "secret", Secret.Get(context) },
                { "lang", Languages.Get(context) }
            };
        }
    }
}
