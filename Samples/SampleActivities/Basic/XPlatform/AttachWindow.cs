using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

using UiPath.Core;
using System.ComponentModel;
using System.Drawing.Drawing2D;

namespace UiPathTeam.XPlatformTest.Activities
{

    [Browsable(true)]
    public class AttachWindow : AsyncCodeActivity
    {
        #region Properties
        [Category("Common")]
        [DisplayName("TimeoutMS")]
        public InArgument<int> TimeoutMS { get; set; } = 60000;


        [Category("Input")]
        [DisplayName("UI Selector")]
        public InArgument<string> Selector { get; set; }

        [Category("Input")]
        [DisplayName("ApplicationWindow")]
        public InArgument<UiPath.Core.Window> ApplicationWindows { get; set; }

        #endregion


        private String strSelector;
        private UiPath.Core.Window xpWin;


        #region Constructors

        public AttachWindow()
        {
        }

        #endregion



        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (Selector == null) metadata.AddValidationError(string.Format("{0}가 지정되지 않았습니다", nameof(Selector)));
            if (ApplicationWindows == null) metadata.AddValidationError(string.Format("{0}가 지정되지 않았습니다.", nameof(ApplicationWindows)));

            base.CacheMetadata(metadata);
        }

        protected void Execute()
        {
            if (strSelector != null)
            {
                GlobalVariable.gSelector = strSelector;

                if (!strSelector.Contains("title"))
                {
                    strSelector += "title=''";
                }

                if (!strSelector.Contains("title") || !strSelector.Contains("wnd app"))
                {
                    Console.WriteLine("Application window information is invalid.");
                }

                XPlatformIntegration.DeclareHandleData(strSelector);
                string stitle = XPlatformIntegration.GetGlobalHandleData("title");

                int iiHandle = XPlatformIntegration.GetHandle(strSelector);

                //System.Windows.MessageBox.Show("AssingWindow Excute iiHandle :: " + iiHandle);

                GlobalVariable.gWindowHandle = (IntPtr)iiHandle;

                //System.Windows.MessageBox.Show("GlobalVariable.gWindowHandle :: " + GlobalVariable.gWindowHandle);

                GlobalVariable.gApplicationWindow = (UiPath.Core.Window)xpWin;

                Console.WriteLine("AssignWindow done.");
            }
        }
       
        protected override IAsyncResult BeginExecute(AsyncCodeActivityContext context, AsyncCallback callback, object state)
        {
            var timeout = TimeoutMS.Get(context);
            strSelector = Selector.Get(context);
            xpWin = ApplicationWindows.Get(context);

            var task = new Task(_ => Execute(), state);
            task.Start();
            if (callback != null)
            {
                task.ContinueWith(s => callback(s));
                task.Wait();
            }
            return task;
        }

        protected override void EndExecute(AsyncCodeActivityContext context, IAsyncResult result)
        {
            
        }

        #endregion
    }
}

