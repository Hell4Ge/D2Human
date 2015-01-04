using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace D2Human
{
    public class D2H
    {
        // IMPORT REQUIRED FUNCTIONS
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool PostMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        private String wName;

        private readonly int[] BTN_CREATE = new int[] { 583, 462 };
        private readonly int[] DIFF_NORM = new int[] { 442, 376 };
        private readonly int[] DIFF_NIGHT = new int[] { 567, 380 };
        private readonly int[] DIFF_HELL = new int[] { 708, 376 };
        private readonly int[] BTN_CREATE_GAME = new int[] { 668, 423 };

        private readonly int[] TB_CREATE_GAMENAME = new int[] { 438, 156 };
        private readonly int[] TB_CREATE_PASSWORD = new int[] { 435, 210 };
        private readonly int[] TB_CREATE_DESCR = new int[] { 464, 265 };
        
        //private readonly int[] CB_LVLDIFF = new int[] { 464, 265 };
        //private readonly int[] TB_LVLDIFF = new int[] { 464, 265 };

        public D2H(String windowName)
        {
            this.wName = windowName;
        }

        public bool PMsg(IntPtr hWnd, int Msg, int wParam, int lParam)
        {
            return PostMessage(hWnd, Msg, wParam, lParam);
        }

        public IntPtr SMsg(IntPtr hWnd, int Msg, int wParam, int lParam)
        {
            return SendMessage(hWnd, Msg, wParam, lParam);
        }

        #region const data
        ///summary>
        /// Virtual Messages
        /// </summary>
        public enum WMessages : int
        {
            WM_LBUTTONDOWN = 0x201, //Left mousebutton down
            WM_LBUTTONUP = 0x202,  //Left mousebutton up
            WM_LBUTTONDBLCLK = 0x203, //Left mousebutton doubleclick
            WM_RBUTTONDOWN = 0x204, //Right mousebutton down
            WM_RBUTTONUP = 0x205,   //Right mousebutton up
            WM_RBUTTONDBLCLK = 0x206, //Right mousebutton doubleclick
            WM_KEYDOWN = 0x100,  //Key down
            WM_KEYUP = 0x101,   //Key up
            WM_CHAR = 0x102   // Char
        }

        /// <summary>
        /// Virtual Keys
        /// </summary>
        public enum VKeys : int
        {
            VK_LBUTTON = 0x01,   //Left mouse button
            VK_RBUTTON = 0x02,   //Right mouse button
            VK_CANCEL = 0x03,   //Control-break processing
            VK_MBUTTON = 0x04,   //Middle mouse button (three-button mouse)
            VK_BACK = 0x08,   //BACKSPACE key
            VK_TAB = 0x09,   //TAB key
            VK_CLEAR = 0x0C,   //CLEAR key
            VK_RETURN = 0x0D,   //ENTER key
            VK_SHIFT = 0x10,   //SHIFT key
            VK_CONTROL = 0x11,   //CTRL key
            VK_MENU = 0x12,   //ALT key
            VK_PAUSE = 0x13,   //PAUSE key
            VK_CAPITAL = 0x14,   //CAPS LOCK key
            VK_ESCAPE = 0x1B,   //ESC key
            VK_SPACE = 0x20,   //SPACEBAR
            VK_PRIOR = 0x21,   //PAGE UP key
            VK_NEXT = 0x22,   //PAGE DOWN key
            VK_END = 0x23,   //END key
            VK_HOME = 0x24,   //HOME key
            VK_LEFT = 0x25,   //LEFT ARROW key
            VK_UP = 0x26,   //UP ARROW key
            VK_RIGHT = 0x27,   //RIGHT ARROW key
            VK_DOWN = 0x28,   //DOWN ARROW key
            VK_SELECT = 0x29,   //SELECT key
            VK_PRINT = 0x2A,   //PRINT key
            VK_EXECUTE = 0x2B,   //EXECUTE key
            VK_SNAPSHOT = 0x2C,   //PRINT SCREEN key
            VK_INSERT = 0x2D,   //INS key
            VK_DELETE = 0x2E,   //DEL key
            VK_HELP = 0x2F,   //HELP key
            VK_0 = 0x30,   //0 key
            VK_1 = 0x31,   //1 key
            VK_2 = 0x32,   //2 key
            VK_3 = 0x33,   //3 key
            VK_4 = 0x34,   //4 key
            VK_5 = 0x35,   //5 key
            VK_6 = 0x36,    //6 key
            VK_7 = 0x37,    //7 key
            VK_8 = 0x38,   //8 key
            VK_9 = 0x39,    //9 key
            VK_A = 0x41,   //A key
            VK_B = 0x42,   //B key
            VK_C = 0x43,   //C key
            VK_D = 0x44,   //D key
            VK_E = 0x45,   //E key
            VK_F = 0x46,   //F key
            VK_G = 0x47,   //G key
            VK_H = 0x48,   //H key
            VK_I = 0x49,    //I key
            VK_J = 0x4A,   //J key
            VK_K = 0x4B,   //K key
            VK_L = 0x4C,   //L key
            VK_M = 0x4D,   //M key
            VK_N = 0x4E,    //N key
            VK_O = 0x4F,   //O key
            VK_P = 0x50,    //P key
            VK_Q = 0x51,   //Q key
            VK_R = 0x52,   //R key
            VK_S = 0x53,   //S key
            VK_T = 0x54,   //T key
            VK_U = 0x55,   //U key
            VK_V = 0x56,   //V key
            VK_W = 0x57,   //W key
            VK_X = 0x58,   //X key
            VK_Y = 0x59,   //Y key
            VK_Z = 0x5A,    //Z key
            VK_NUMPAD0 = 0x60,   //Numeric keypad 0 key
            VK_NUMPAD1 = 0x61,   //Numeric keypad 1 key
            VK_NUMPAD2 = 0x62,   //Numeric keypad 2 key
            VK_NUMPAD3 = 0x63,   //Numeric keypad 3 key
            VK_NUMPAD4 = 0x64,   //Numeric keypad 4 key
            VK_NUMPAD5 = 0x65,   //Numeric keypad 5 key
            VK_NUMPAD6 = 0x66,   //Numeric keypad 6 key
            VK_NUMPAD7 = 0x67,   //Numeric keypad 7 key
            VK_NUMPAD8 = 0x68,   //Numeric keypad 8 key
            VK_NUMPAD9 = 0x69,   //Numeric keypad 9 key
            VK_SEPARATOR = 0x6C,   //Separator key
            VK_SUBTRACT = 0x6D,   //Subtract key
            VK_DECIMAL = 0x6E,   //Decimal key
            VK_DIVIDE = 0x6F,   //Divide key
            VK_F1 = 0x70,   //F1 key
            VK_F2 = 0x71,   //F2 key
            VK_F3 = 0x72,   //F3 key
            VK_F4 = 0x73,   //F4 key
            VK_F5 = 0x74,   //F5 key
            VK_F6 = 0x75,   //F6 key
            VK_F7 = 0x76,   //F7 key
            VK_F8 = 0x77,   //F8 key
            VK_F9 = 0x78,   //F9 key
            VK_F10 = 0x79,   //F10 key
            VK_F11 = 0x7A,   //F11 key
            VK_F12 = 0x7B,   //F12 key
            VK_SCROLL = 0x91,   //SCROLL LOCK key
            VK_LSHIFT = 0xA0,   //Left SHIFT key
            VK_RSHIFT = 0xA1,   //Right SHIFT key
            VK_LCONTROL = 0xA2,   //Left CONTROL key
            VK_RCONTROL = 0xA3,    //Right CONTROL key
            VK_LMENU = 0xA4,      //Left MENU key
            VK_RMENU = 0xA5,   //Right MENU key
            VK_PLAY = 0xFA,   //Play key
            VK_ZOOM = 0xFB, //Zoom key
        }
        #endregion

        public int MakeLParam(int LoWord, int HiWord)
        {
            return ((HiWord << 16) | (LoWord & 0xffff));
        }

        public IntPtr FindWindow(string wName)
        {
            return FindWindow(null, wName);
        }
        /// <summary>
        /// Sends a key to the game window
        /// </summary>
        /// <param name="wName">Window name</param>
        /// <param name="VK">Use WMessages for easier handling with int equivalents</param>
        public void SendVKey(String wName, int VK)
        {
            PMsg(FindWindow(wName), (int)WMessages.WM_KEYDOWN,  VK,     0);
            PMsg(FindWindow(wName), (int)WMessages.WM_CHAR,     VK,     0);
            PMsg(FindWindow(wName), (int)WMessages.WM_KEYUP,    VK,     0);
        }

        /// <summary>
        /// Selects a character from character slot
        /// </summary>
        /// <param name="wName">Window name</param>
        /// <param name="slot">Position of character, odd values are on the right column</param>
        public void SelectCharacter(int slot)
        {
            if (slot == 1)
            {
                SendVKey(wName, (int)VKeys.VK_RETURN);
            }
            else
            {
                if (slot == 2)
                {
                    // right & enter
                    SendVKey(wName, (int)VKeys.VK_RIGHT);
                    SendVKey(wName, (int)VKeys.VK_RETURN);
                    return;
                }

                if (slot % 2 == 0)  // Right column
                {
                    // move right
                    SendVKey(wName, (int)VKeys.VK_RIGHT);
                    for (int i = 0; i < ((slot / 2) - 1); i++)
                    {
                        SendVKey(wName, (int)VKeys.VK_DOWN);
                    }
                    SendVKey(wName, (int)VKeys.VK_RETURN);
                }
                else                // Left column
                {
                    for (int i = 0; i < ((slot / 2) - 1); i++)
                    {
                        SendVKey(wName, (int)VKeys.VK_DOWN);
                    }
                    SendVKey(wName, (int)VKeys.VK_RETURN);
                }
            }
        }
        /// <summary>
        /// Passes you to character select screen with given credentials
        /// </summary>
        /// <param name="wName">Window name</param>
        /// <param name="user">Account name</param>
        /// <param name="password">Account password</param>
        /// <param name="serverLoginWait">Maximum time between clicking on battle.net and character login box</param>
        /// <param name="additionalDelay">Default 2000, you can make it shorter but usually it just sets some delays between actions, pay attention when changing this value</param>
        public void Login(String user, String password, int serverLoginWait)
        {
            
            Click(398, 348);   // Click battle.net
            Sleep(serverLoginWait);
            // Doubleclick on accname field

            Click(335, 340);
            Sleep(50);
            Click(335, 340);

            Sleep(25);
            SendString(user);                // Send username
            Sleep(25); 
            Click( 339, 393);
            Sleep(25); 
            SendString(password);            // Send password
            Sleep(25);
            
            Click(403, 470);           // Click login
            Sleep(500);
        }

        public void JoinGame(String gameName, String password="")
        {
            var hwnd = FindWindow(wName);
            Click( 714, 463);             // Join
            Sleep(150);
            SendString(gameName);        // Type gamename
            Sleep(150);
            if(!password.Equals(""))
            {
                SendVKey(wName, (int)VKeys.VK_TAB);
                Sleep(50);
                SendString(password);
                Sleep(50);
            }
            SendVKey(wName, (int)VKeys.VK_RETURN);
        }

        public void QuitGame()
        {
            SendVKey(wName, (int)VKeys.VK_ESCAPE);
            Sleep(750);
            Click( 400, 260);
        }

        public void CreateGame(String name, String diff="", String pw="")
        {
            Click(BTN_CREATE);
            Sleep(100);
            Click(TB_CREATE_GAMENAME);
            Sleep(50);
            SendString(name);

            if(!pw.Equals(""))
            {
                Sleep(100);
                Click(TB_CREATE_PASSWORD);
                Sleep(50);
                SendString(pw);
                Sleep(100);
            }

            if(!diff.Equals(""))
            {
                if (diff.Equals("Normal") || diff.Equals("Norm"))
                    Click(DIFF_NORM);

                if (diff.Equals("Nightmare") || diff.Equals("Night"))
                    Click(DIFF_NIGHT);

                if (diff.Equals("Hell"))
                    Click(DIFF_HELL);

                Sleep(50);
            }

            Click(BTN_CREATE_GAME);
        }

        public void Click(int[] location)
        {
            IntPtr hWnd = FindWindow(null, wName);
            int LParam = MakeLParam(location[0], location[1]);

            int btnDown = 0;
            int btnUp = 0;

            btnDown = (int)WMessages.WM_LBUTTONDOWN;
            btnUp = (int)WMessages.WM_LBUTTONUP;

            SendMessage(hWnd, 0x0020, 0, LParam);   // set cursor
            System.Threading.Thread.Sleep(50);
            SendMessage(hWnd, btnDown, 0, LParam);
            SendMessage(hWnd, btnUp, 0, LParam);
        }

        public void Click(string button, int x, int y, bool doubleClick)
        {
            IntPtr hWnd = FindWindow(null, wName);
            int LParam = MakeLParam(x, y);

            int btnDown = 0;
            int btnUp = 0;

            if (button == "left")
            {
                btnDown = (int)WMessages.WM_LBUTTONDOWN;
                btnUp = (int)WMessages.WM_LBUTTONUP;
            }

            if (button == "right")
            {
                btnDown = (int)WMessages.WM_RBUTTONDOWN;
                btnUp = (int)WMessages.WM_RBUTTONUP;
            }


            if (doubleClick == true)
            {
                if(button.Equals("left"))
                    SendMessage(hWnd, (int)WMessages.WM_LBUTTONDBLCLK, 0, LParam);

                if(button.Equals("right"))
                    SendMessage(hWnd, (int)WMessages.WM_RBUTTONDBLCLK, 0, LParam);
            }

            if (doubleClick == false)

            {
                SendMessage(hWnd, btnDown, 0, LParam);
                SendMessage(hWnd, btnUp, 0, LParam);
            }

        }
        public void Click(string button, int x, int y)
        {
            IntPtr hWnd = FindWindow(null, wName);
            int LParam = MakeLParam(x, y);

            int btnDown = 0;
            int btnUp = 0;

            if (button == "left")
            {
                btnDown = (int)WMessages.WM_LBUTTONDOWN;
                btnUp = (int)WMessages.WM_LBUTTONUP;
            }

            if (button == "right")
            {
                btnDown = (int)WMessages.WM_RBUTTONDOWN;
                btnUp = (int)WMessages.WM_RBUTTONUP;
            }


            SendMessage(hWnd, btnDown, 0, LParam);
            SendMessage(hWnd, btnUp, 0, LParam);
        }
        public void Click(int x, int y)
        {
            IntPtr hWnd = FindWindow(null, wName);
            int LParam = MakeLParam(x, y);

            int btnDown = 0;
            int btnUp = 0;

            btnDown = (int)WMessages.WM_LBUTTONDOWN;
            btnUp = (int)WMessages.WM_LBUTTONUP;

            SendMessage(hWnd, 0x0020, 0, LParam);   // set cursor
            System.Threading.Thread.Sleep(50);
            SendMessage(hWnd, btnDown, 0, LParam);
            SendMessage(hWnd, btnUp, 0, LParam);
        }

        /// <summary>
        /// Send a character to window, USE SendString INSTEAD!!!
        /// </summary>
        /// <param name="wName">Window name</param>
        /// <param name="key">Character to send</param>
        public void SendChar(char key)
        {
            const Int32 WM_CHAR = 0x102;
            IntPtr hWnd = FindWindow(null, wName);

            PostMessage(hWnd, WM_CHAR, key, 0);
        }
        /// <summary>
        /// Sends a string of chars to the game window
        /// </summary>
        /// <param name="wName">Window name</param>
        /// <param name="str">Characters to send</param>
        public void SendString(string str)
        {
            const Int32 WM_CHAR = 0x102;
            IntPtr hWnd = FindWindow(null, wName);

            foreach (char c in str)
                    PostMessage(hWnd, WM_CHAR, c, 0);
        }

        private void Sleep(int ms)
        {
            System.Threading.Thread.Sleep(ms);
        }
    }
}
