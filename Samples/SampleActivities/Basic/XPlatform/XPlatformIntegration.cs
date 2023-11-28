using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using UiPath.Core.Activities;

namespace UiPathTeam.XPlatformTest.Activities
{
    public class GlobalVariable
    {
        // 2020.11.26 batman 
        public static string gAppNm = "";
        public static string gTitle = "";
        public static string cls = "";

        public static string gSelector = "";
        // 2020.11.26 batman 

        public static UiPath.Core.Window gApplicationWindow = null;
        public static IntPtr gWindowHandle = IntPtr.Zero;
    }
    struct COPYDATASTRUCT
    {
        public IntPtr dwData;
        public int cbData;
        //[MarshalAs(UnmanagedType.LPStr)]
        public IntPtr lpData;
    }

    public class ObjectAttribute
    {
        public string Type { set; get; }
        public string Name { set; get; }
        public string Value { get; set; }
        public bool Visible { get; set; }
        public int X { get; set; }
        public int Y { get; set; }
        public int W { get; set; }
        public int H { get; set; }

        public override string ToString()
        {
            return string.Format("Type={0}, Name={1}, Value={2}, Visibility={3} Region.X={4} Region.Y={5} Region.W={6} Region.H={7}",
                this.Type, this.Name, this.Value, this.Visible, this.X, this.Y, this.W, this.H);
        }
    }
    public class VersionInfo
    {
        public string Version { get; set; }
        public string Build { get; set; }
        public string Licensed { get; set; }

        public string ProductName { get; set; }
        public string CustomerName { get; set; }
        public string LicenseLevel { get; set; }

        public override string ToString()
        {
            return string.Format("Version={0}, Build={1}, LicenseDate={2}, ProductName={3}, CustomerName={4}, LicenseLevel={5}", this.Version, this.Build, this.Licensed, this.ProductName, this.CustomerName, this.LicenseLevel);
        }
    }

    public enum ClickType
    {
        Double = 0,
        Single
    }

    public enum CheckType
    {
        Check = 0,
        Uncheck,
        Toggle
    }

    public enum GridCheckType
    {
        Check = 0,
        Uncheck
    }

    public enum GridCellType
    {
        Body = 0,
        Head,
        Summary
    }
    internal class ClipboardAsync
    {
        public static void Clear()
        {
            Clipboard.SetDataObject("", false);
        }


        public static string GetText(int timeoutMS = 30000)
        {
            var result = string.Empty;
            result = Clipboard.GetText();
            var startTime = DateTime.Now.AddMilliseconds(timeoutMS);
            while (result == string.Empty && startTime > DateTime.Now)
            {
                result = Clipboard.GetText();
                Thread.Sleep(20);
            }
            return result;
        }
    }
    internal class XPlatformIntegration
    {
        public const int WM_COPYDATA = 0x004A;
        public static int VERSION = 1;
        public static string RPAProductCode = "UP";

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int FindWindowEx(int hWnd1, int hWnd2, string lpsz1, string lpsz2);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, uint wParam, ref COPYDATASTRUCT lParam);

        static Random reqCode = new Random((int)DateTime.Now.ToFileTimeUtc());
        static int MaxRetryCount = 150;
        static int BEFORE_SEND_MESSAGE = 100;
        static int WAIT_FOR_NEXT_CHECK = 200;


        public static string BuilderHeader(string CmdString, out int RandCode)
        {
            RandCode = reqCode.Next(0, 9999);

            string machine_user = Environment.MachineName + "/" + Environment.UserName;

            string reqString = string.Format("{0}{1:0000}XP{2:00}{3}{4}|", RPAProductCode, RandCode, VERSION, CmdString, machine_user);

            return reqString;

        }

        public static string getResponse()
        {
            var response = ClipboardAsync.GetText();
#if DEBUG
            if (response.Length > 8)
                Console.WriteLine(response.Substring(8));
#endif
            return response;
        }

        public static int buildText(IntPtr winHandle, string elementId, string reqType, out int reqRandCode)
        {
            int retval = 0;
            reqRandCode = 0;
            string CmdString = "CGTVAL";

            string reqString = BuilderHeader(CmdString, out reqRandCode);

            // reqString 만들기
            string reqHdr = string.Format("{0}{1}|{2}", reqString, elementId, reqType);

#if DEBUG
            Console.WriteLine(string.Format("ToConnector: {0}", reqHdr.Substring(6)));
#endif
            try
            {
                if (winHandle != IntPtr.Zero)
                {
                    COPYDATASTRUCT cds = new COPYDATASTRUCT();
                    cds.dwData = IntPtr.Zero;
                    cds.cbData = 2 * System.Text.Encoding.Default.GetBytes(reqHdr).Length;
                    cds.lpData = Marshal.StringToHGlobalUni(reqHdr);

                    // 1125 - Clipboard 확인  ///////////////////////////////////
                    ClipboardAsync.Clear();
                    /////////////////////////////////////////////////////////////

                    Thread.Sleep(BEFORE_SEND_MESSAGE);
                    SendMessage(winHandle, WM_COPYDATA, 0, ref cds);
                }
                else
                {
                    retval = 404;
                }
            }
            catch (Exception e)
            {
                System.Console.WriteLine(string.Format("Exception: {0}", e.Message));
                System.Console.WriteLine("Detail : " + e.StackTrace);
                retval = -1;
            }
            return retval;
        }

        public static void getText(out string text, int codeCheck)
        {
            var resp = getResponse();
            text = string.Empty;
            if (resp.StartsWith(string.Format("{0:0000}", codeCheck)) && resp.Contains("SUCC") && resp.Length >= 8)
            {
#if DEBUG
                Console.WriteLine(string.Format("FromConnector: {0}", resp.Substring(8)));
#endif
                text = resp.Substring(8);
            }
            else
            {
                if (string.IsNullOrEmpty(resp))
                {
                    throw new ActivityTimeoutException();
                }
                else
                {
                    var msg = resp; // resp.Length >= 8 ? resp.Substring(8) : resp;
                    Console.WriteLine(msg);
                    var st = new StackTrace();
                    var sf = st.GetFrame(0);

                    throw new Exception(string.Format("XPlatformConnector {0} 호출시 오류가 발생했습니다. {1}", sf.GetMethod(), msg));
                }
            }
        }

        // 2020.11.26 batman 
        static Dictionary<string, string> SelectorKey = new Dictionary<string, string>
        {
            {"title","title" },
            { "cls","cls" },
            { "wndapp" , "wnd app"}
        };
        public static void DeclareHandleData(string strSelector)
        {
            string tmpSelector = strSelector;
            foreach (KeyValuePair<string, string> selkey in SelectorKey)
            {
                string[] arrTemp = tmpSelector.Split(new string[] { selkey.Value }, StringSplitOptions.None);
                tmpSelector = arrTemp[0];
                //string[] arrData = arrTemp[1].Split('\'');
                int iStartidx = arrTemp[1].IndexOf('\'');
                //int iEndidx = arrTemp[1].LastIndexOf('\'');
                int iEndidx = arrTemp[1].IndexOf('\'', iStartidx + 1);
                string ret = arrTemp[1].Substring(iStartidx + 1, iEndidx - iStartidx - 1);

                switch (selkey.Key)
                {
                    case "title":
                        GlobalVariable.gTitle = ret;
                        break;
                    case "cls":
                        GlobalVariable.cls = ret;
                        break;
                    case "wndapp":
                        if (ret.Contains(".")) //확장자 제거
                        {
                            ret = ret.Substring(0, ret.LastIndexOf("."));
                        }
                        GlobalVariable.gAppNm = ret;
                        break;
                    default:

                        break;

                }
            }
        }

        public static string GetGlobalHandleData(string sName)
        {
            if (sName.Equals("title"))
            {
                return GlobalVariable.gTitle;
            }
            else if (sName.Equals("cls"))
            {
                return GlobalVariable.cls;
            }
            else if (sName.Equals("wndapp"))
            {
                return GlobalVariable.gAppNm;
            }
            else if (sName.Equals("selector"))
            {
                return GlobalVariable.gSelector;
            }
            else
            {
                return "";
            }

        }

        public static int GetHandle(string strSelector)
        {
            DeclareHandleData(strSelector);

            IntPtr ipWindowHandle = FindGetWindowHandleNew(GlobalVariable.gAppNm, GlobalVariable.gTitle, GlobalVariable.cls);
            //            IntPtr ipWindowHandle = DesignerMetadata.FindGetWindowHandle(GlobalDesignerVariable.gAppNm, GlobalDesignerVariable.gTitle);
            GlobalVariable.gWindowHandle = ipWindowHandle;
            if (ipWindowHandle == IntPtr.Zero)
            {
                Console.Write("No application window information is available.");
                return 0;
            }
            int iHandle = getWindowHandle(ipWindowHandle); //ActivX 핸들찾기
            return iHandle;
        }

        public static IntPtr FindGetWindowHandleNew(string strPronm, string strTitle, string strCls)
        {
            IntPtr ipHandle = IntPtr.Zero;
            IntPtr result = IntPtr.Zero;


            //System.Windows.MessageBox.Show(string.Format("strProssNm='{0}' strProTitle='{1}", strPronm, strTitle));
            Process[] pro = Process.GetProcessesByName(strPronm);

            if (pro.Length > 0)
            {
                // Title을 지정하지 않았을 경우에 타이틀은 비교하지 않고 첫번째 프로세스를 선택
                if (strTitle.Equals("") || strTitle.Equals("*"))
                {
                    // 이름이 맞는게 존재하면 첫번째 프로세스의 메인윈도우 핸들을 리턴
                    // 아니면 Null
                    for (int i = 0; i < pro.Length; i++)
                    {
                        if (pro[i].MainWindowTitle == "")
                        {
                            ipHandle = pro[i].MainWindowHandle;
                        }
                    }

                }
                else
                {
                    for (int i = 0; i < pro.Length; i++)
                    {
                        if (pro[i].ProcessName.ToLower() == strPronm.ToLower())
                        {
                            if (
                                (strTitle.LastIndexOf('*') > 0 && pro[i].MainWindowTitle.ToLower().Contains(strTitle.ToLower().Substring(0, strTitle.LastIndexOf('*')))) //타이틀 끝에 '*' 있는경우
                                || pro[i].MainWindowTitle.ToLower() == strTitle.ToLower()
                               )
                            {
                                ipHandle = pro[i].MainWindowHandle;
                                break;
                            }

                        }
                    }
                }
            }

            if (ipHandle != IntPtr.Zero)
                result = ipHandle;
            else
            {
                result = FindWindow(strCls, strTitle);
            }

            return result;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr FindWindow(String lpClassName, String lpWindowName);

        public static int getWindowHandle(IntPtr winHandle)
        {
            int iHandle;

            if (!IsWindow(winHandle))
            {
                var msg = string.Format("지정된 화면은 더 이상 유효하지 않습니다.{0}", winHandle);
                Console.WriteLine(msg);
                var st = new StackTrace();
                var sf = st.GetFrame(0);

                throw new Exception(string.Format("XPlatformConnector {0} 호출시 오류가 발생했습니다. {1}", sf.GetMethod(), msg));
            }

            iHandle = FindWindowEx((int)winHandle, 0, "Frame Tab", null);
            if (iHandle > 0)
            {   // XP Mi 구조가 다름
                iHandle = FindWindowEx(iHandle, 0, "TabWindowClass", null);
                iHandle = FindWindowEx(iHandle, 0, "Shell DocObject View", null);
                iHandle = FindWindowEx(iHandle, 0, "Internet Explorer_Server", null);
                iHandle = FindWindowEx(iHandle, 0, null, null);

                int nhwnd1 = 0;
                while (true)
                {
                    nhwnd1 = FindWindowEx(iHandle, nhwnd1, "CyWindowClass", null);
                    if (nhwnd1 > 0)
                    {
#if DEBUG
                        Console.WriteLine(string.Format("Find iHandle: {0}", nhwnd1.ToString()));
#endif
                        break;
                    }
                    else
                    {
#if DEBUG
                        Console.WriteLine(string.Format("Do not find iHandle: {0}", nhwnd1.ToString()));
#endif
                        break;
                    }
                }

                if (nhwnd1 > 0) iHandle = nhwnd1;
            }

            if (iHandle > 0)
            {
                return iHandle;
            }
            else
            {
                return (int)winHandle;
            }
        }

        public static int GetValidHandle(UiPath.Core.Window xpwin)
        {
            // System.Windows.MessageBox.Show("XP  >>>> GlobalVariable.gWindowHandle  :: " + GlobalVariable.gWindowHandle + " ::: " + xpwin);

            int iHandle = 0;
            if (GlobalVariable.gWindowHandle != IntPtr.Zero)
            {
                iHandle = XPlatformIntegration.getWindowHandle(GlobalVariable.gWindowHandle);
            }
            else if (xpwin != null)
            {
                iHandle = XPlatformIntegration.getWindowHandle(xpwin.Handle);
            }
            else
            {
                Console.WriteLine("No Application window information is available.");
            }
            return iHandle;

        }
    }
}
