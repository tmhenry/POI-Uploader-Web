using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

//To use the windows API for mouse input and output
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using POILibCommunication;

namespace Communication
{
    public static class EmulateIO
    {
        [DllImport("user32.dll", EntryPoint = "SendInput", SetLastError = true)]
        static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        public static void KeyStroke(string keys)
        {
            //Call sendkeys to emulate keyboard input
            SendKeys.SendWait(keys);
        }

        public static uint MouseClick()
        {
            INPUT input_down = new INPUT();
            input_down.mi.dx = 0;
            input_down.mi.dy = 0;
            input_down.mi.mouseData = 0;
            input_down.mi.dwFlags = (int)MOUSEEVENTF.LEFTDOWN;

            INPUT input_up = input_down;
            input_up.mi.dwFlags = (int)MOUSEEVENTF.LEFTUP;

            INPUT[] input = { input_down, input_up };
            return SendInput(2, input, Marshal.SizeOf(input_down));
        }

        public static uint MouseMove(double x, double y)
        {
            // Bildschirm Auflösung
            float ScreenWidth = Screen.PrimaryScreen.Bounds.Width;
            float ScreenHeight = Screen.PrimaryScreen.Bounds.Height;

            INPUT input_move = new INPUT();
            input_move.mi.dx = (int)Math.Round(x * (65535 / ScreenWidth), 0);
            input_move.mi.dy = (int)Math.Round(y * (65535 / ScreenHeight), 0);

            POIGlobalVar.POIDebugLog(input_move.mi.dx + " " + input_move.mi.dy);

            input_move.mi.mouseData = 0;
            input_move.mi.dwFlags = (int)(MOUSEEVENTF.MOVE | MOUSEEVENTF.ABSOLUTE);

            INPUT[] input = { input_move };
            return SendInput(1, input, Marshal.SizeOf(input_move));
        }

        #region struct for mouse input
        private enum InputType
        {
            INPUT_MOUSE = 0,
            INPUT_KEYBOARD = 1,
            INPUT_HARDWARE = 2,
        }

        [Flags()]
        private enum MOUSEEVENTF
        {
            MOVE = 0x0001,  // mouse move 
            LEFTDOWN = 0x0002,  // left button down
            LEFTUP = 0x0004,  // left button up
            RIGHTDOWN = 0x0008,  // right button down
            RIGHTUP = 0x0010,  // right button up
            MIDDLEDOWN = 0x0020,  // middle button down
            MIDDLEUP = 0x0040,  // middle button up
            XDOWN = 0x0080,  // x button down 
            XUP = 0x0100,  // x button down
            WHEEL = 0x0800,  // wheel button rolled
            VIRTUALDESK = 0x4000,  // map to entire virtual desktop
            ABSOLUTE = 0x8000,  // absolute move
        }

        [Flags()]
        private enum KEYEVENTF
        {
            EXTENDEDKEY = 0x0001,
            KEYUP = 0x0002,
            UNICODE = 0x0004,
            SCANCODE = 0x0008,
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public int mouseData;
            public int dwFlags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public short wVk;
            public short wScan;
            public int dwFlags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct HARDWAREINPUT
        {
            public int uMsg;
            public short wParamL;
            public short wParamH;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct INPUT
        {
            [FieldOffset(0)]
            public int type;
            [FieldOffset(4)]
            public MOUSEINPUT mi;
            [FieldOffset(4)]
            public KEYBDINPUT ki;
            [FieldOffset(4)]
            public HARDWAREINPUT hi;
        }
        #endregion
    }
}
