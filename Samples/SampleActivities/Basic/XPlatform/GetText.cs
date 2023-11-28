using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using UiPath.Core;
using System.ComponentModel;


namespace UiPathTeam.XPlatformTest.Activities
{
    public enum GetTextValueType
    {
        Value = 0,
        Text,
    }


    [Browsable(true)]
    public class GetText : AsyncCodeActivity
    {
        #region Properties

        [Category("Common")]
        [DisplayName("TimeoutMS")]
        public InArgument<int> TimeoutMS { get; set; } = 60000;

        [Category("Input")]
        [DisplayName("ApplicationWindow")]
        [RequiredArgument]
        public InArgument<UiPath.Core.Window> ApplicationWindows { get; set; }

        [Category("Input")]
        [DisplayName("TextValueType")]
        public InArgument<GetTextValueType> Type { get; set; } = GetTextValueType.Value;

        [Category("Input")]
        [DisplayName("Element Name")]
        [RequiredArgument]
        public InArgument<string> ElementName { get; set; }

        [Category("Output")]
        [DisplayName("Text")]
        [RequiredArgument]
        public OutArgument<string> Text { get; set; }

        #endregion


        private Window xpWin;
        private GetTextValueType gettype;
        private String elementName;
        private String outText = string.Empty;
        #region Constructors

        public GetText()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (ApplicationWindows == null) metadata.AddValidationError(string.Format("{0}가 지정되지 않았습니다", nameof(ApplicationWindows)));
            if (Type == null) metadata.AddValidationError(string.Format("{0}이 지정되지 않았습니다", nameof(Type)));
            if (ElementName == null) metadata.AddValidationError(string.Format("{0}가 지정되지 않았습니다.", nameof(ElementName)));

            base.CacheMetadata(metadata);
        }

        
        private void Execute(string reqType)
        {
            int codeCheck = 0;
            if (xpWin is null) xpWin = GlobalVariable.gApplicationWindow;
            int iHandle = XPlatformIntegration.GetValidHandle(xpWin);
            var retval = XPlatformIntegration.buildText((IntPtr)iHandle, elementName, reqType, out codeCheck);
            XPlatformIntegration.getText(out outText, codeCheck);
        }

        protected override IAsyncResult BeginExecute(AsyncCodeActivityContext context, AsyncCallback callback, object state)
        {
            var timeout = TimeoutMS.Get(context);
            xpWin = ApplicationWindows.Get(context);
            gettype = Type.Get(context);
            elementName = ElementName.Get(context).Trim();

            string reqType = "Text";

            if (gettype == GetTextValueType.Value)
                reqType = "Value";

            var task = new Task(_ => Execute(reqType), state);
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
            Text.Set(context, outText);
            await task;
        }

        #endregion
    }
}

