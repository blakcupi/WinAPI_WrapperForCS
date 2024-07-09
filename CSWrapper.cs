using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace KimCGLib
{
    public partial class CSWrapper
    {
        //사용할 API함수를 임포트 한다.
        [DllImport("USER32.DLL")]
        public static extern uint FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern uint FindWindowEx(uint hWnd1, uint hWnd2, string lpsz1, string lpsz2);

        [DllImport("user32.dll")]
        public static extern uint SendMessage(uint hwnd, uint wMsg, uint wParam, uint lParam);

        [DllImport("user32.dll")]
        public static extern uint PostMessage(uint hwnd, uint wMsg, uint wParam, uint lParam);
    }
}

