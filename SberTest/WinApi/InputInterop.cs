using System;
using System.Collections.Generic;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using OpenQA.Selenium.DevTools.V105.Audits;

namespace SberTest
{
    internal static class InputInterop
    {
        [DllImport("user32.dll")]
        public static extern ushort SendInput(ushort inputsLength, INPUT[] inputs, short inputSize);

        [DllImport("user32.dll")]
        public static extern IntPtr GetMessageExtraInfo();

        [DllImport("user32.dll")]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr LoadKeyboardLayout(string pwszKLID, ushort Flags);

        public const string ruLanguage = "00000419";
        public const string enLanguage = "00000409";

        public struct KEYBDINPUT
        {
            public ushort Vk;
            public ushort Scan;
            public uint Flags;
            public uint Time;
            public IntPtr ExtraInfo;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct INPUT
        {
            [FieldOffset(0)]
            public InputType Type;

            [FieldOffset(4)]
            public MOUSEINPUT MouseInputInfo;
            [FieldOffset(4)]
            public KEYBDINPUT KeyboardInputInfo;
            [FieldOffset(4)]
            public HARDWAREINPUT HardwareInputInfo;
        }

        public struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        public struct HARDWAREINPUT
        {
            public ushort uMsg;
            public short wParamL;
            public short wParamH;
        }

        public enum InputType : uint
        {
            Mouse = 0,
            Keyboard = 1,
            Hardware = 2
        }

        public static void SendKeyboardInput(ushort vkCode, uint flag)
        {
            var inputs = new INPUT[1];
            inputs[0].Type = InputType.Keyboard;
            inputs[0].KeyboardInputInfo = new KEYBDINPUT()
            {
                Vk = vkCode,
                Flags = flag,
                ExtraInfo = GetMessageExtraInfo()
            };

            SendInput((ushort)inputs.Length, inputs, (short)Marshal.SizeOf(typeof(INPUT)));
        }

        public static void SendHoldButtonMessage(Key key) => SendKeyboardInput((ushort)KeyInterop.VirtualKeyFromKey(key), 0);

        public static void SendReleaseButtonMessage(Key key) => SendKeyboardInput((ushort)KeyInterop.VirtualKeyFromKey(key), 2);

        public static void SendButtonDownMessage(Key key)
        {
            SendHoldButtonMessage(key);
            SendReleaseButtonMessage(key);
        }

        public static void ChangeInputLanguage(IntPtr handle, string language)
        {
            var WM_INPUTLANGCHANGEREQUEST = (ushort)0x0050;
            var INPUTLANGCHANGE_SYSCHARSET = (IntPtr)0x0001;
            var KLF_ACTIVATE = (ushort)0x00000001;
            var layout = LoadKeyboardLayout(language, KLF_ACTIVATE);

            SendMessage(handle, WM_INPUTLANGCHANGEREQUEST, INPUTLANGCHANGE_SYSCHARSET, layout);
        }
    }
}
