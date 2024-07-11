using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Windows.Automation; // UIAutomationClient.dll 참조추가 후 사용가능

namespace KimCGLib
{
    public partial class CSWrapper
    {
        //사용할 API함수를 임포트 한다.
        [DllImport("USER32.DLL")]
        public static extern uint FindWindow(string lpClassName, string lpWindowName); // 주어진 윈도우의 핸들 조회(null, "윈도우 창 제목")

        [DllImport("user32.dll")]
        public static extern uint FindWindowEx(uint hWnd1, uint hWnd2, string lpsz1, string lpsz2); // 주어진 윈도우의 하위 객체(버튼, 텍스트, 라디오, ...) 핸들 조회(핸들1, 핸들2, "객체명", "객체캡션")

        [DllImport("user32.dll")]
        public static extern uint SendMessage(uint hwnd, uint wMsg, uint wParam, uint lParam); // 동기식 메세지(이벤트) 전송(핸들, 메세지, 0, 0)

        [DllImport("user32.dll")]
        public static extern uint PostMessage(uint hwnd, uint wMsg, uint wParam, uint lParam); // 비동기식 메세지(이벤트) 전송(핸들, 메세지, 0, 0)

        [DllImport("user32")]
        public static extern bool SetForegroundWindow(IntPtr handle); // 윈도우 앞으로 올리기

        // 윈도우 핸들 얻기...2024.07.09
        public IntPtr GetWindowHandle(string pProcessName)
        {
            IntPtr handle = null;
            foreach(Process process in Process.GetProcesses())
            {
            	if(process.ProcessName.StartsWith(pProcessName))
                {
                    handle = process.MainWindowHandle;
                    break;
                }
            }

            return handle;
        }

        // AutomationElement를 위한 윈도우 핸들 얻기...2024.07.10
        public AutomationElement GetAutomationElement(string pProcessName)
        {
            AutomationElement automationElement = null;
            
            IntPtr handle = GetWindowHandle(pProcessName);
            if(handle != null && handle != IntPtr.Zero)
            {
                automationElement = AutomationElement.FromHandle(handle);
            }
                
            return automationElement;    
        }

        // 외부 프로그램의 버튼 클릭 처리하기...2024.07.10
        public bool ExternalButtonClickById(string pObjectID) // using spy++
        {
            bool isSuccess = true;
            AutomationElement automationElement = SetMainWindow("---");
            PropertyCondition propCondition = new PropertyCondition(AutomationElement.AutomationProperty, pObjectID);
            AutomationElement aeObject = automationElement.FindFirst(TreeScope.Decendants, propCondition);
            if(aeObject != null)
            {
                try
                {
                    InvokePattern invPattern = aeObject.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                    invPattern?.Invoke();
                }
                catch
                {
                    isSuccess = false;
                }
            }

            return isSuccess;
        }

        // 외부 프로그램의 버튼 클릭 처리하기...2024.07.11
        public bool ExternalButtonClickByName(string pObjectName)
        {
            bool isSuccess = true;
            AutomationElement automationElement = SetMainWindow("---");
            PropertyCondition propCondition = new PropertyCondition(AutomationElement.NameProperty, pObjectName);
            AutomationElement aeObject = automationElement.FindFirst(TreeScope.Decendants, propCondition);
            if(aeObject != null)
            {
                try
                {
                    InvokePattern invPattern = aeObject.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                    invPattern?.Invoke();
                }
                catch
                {
                    isSuccess = false;
                }
            }

            return isSuccess;
        }
    }
}

