using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;

namespace SberTest
{
    internal class SecondProblemSolver : ISolver
    {
        private readonly string _textToPaste;
        private readonly string _outputPath;

        public SecondProblemSolver(string filename, string testToPaste)
        {
            _textToPaste = testToPaste;
            _outputPath = $"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}/{filename}";

            if (File.Exists(_outputPath))
            {
                throw new ArgumentException($"Файл по пути {_outputPath} уже существует");
            }
        }

        public void Solve()
        {
            using (var notepad = Process.Start("notepad.exe", _outputPath))
            {
                try
                {
                    notepad.WaitForInputIdle();

                    InputInterop.SendButtonDownMessage(Key.Enter);
                    Clipboard.SetText(_textToPaste);

                    BeginParagraph(Key.D1);
                    PasteTextFromClipboard();
                    EndParagraph();

                    BeginParagraph(Key.D2);
                    WriteRandomText(notepad.MainWindowHandle);
                    EndParagraph();

                    Save();
                }
                finally
                {
                    notepad.Close();
                }
            };
        }

        private void BeginParagraph(Key key)
        {
            InputInterop.SendButtonDownMessage(key);
            InputInterop.SendButtonDownMessage(Key.Space);
        }

        private void EndParagraph()
        {
            InputInterop.SendButtonDownMessage(Key.Enter);
        }

        private void Save()
        {
            InputInterop.SendHoldButtonMessage(Key.LeftCtrl);
            InputInterop.SendButtonDownMessage(Key.W);
            InputInterop.SendReleaseButtonMessage(Key.LeftCtrl);
            InputInterop.SendButtonDownMessage(Key.Enter);
        }

        private void WriteRandomText(IntPtr handle)
        {
            var random = new Random();
            var stringLength = random.Next(25, 100);

            for (var i = 0; i < stringLength; i++)
            {
                var randomCharRanges = new List<(int Min, int Max)>()
                {
                    (48, 91),
                    (186, 224)
                };

                var pressShift = random.Next(0, 2) == 1;
                var changeLanguage = random.Next(0, 2) == 1;
                
                if (pressShift)
                {
                    InputInterop.SendHoldButtonMessage(Key.LeftShift);
                }
                else
                {
                    InputInterop.SendReleaseButtonMessage(Key.LeftShift);
                }

                if (changeLanguage)
                {
                    InputInterop.ChangeInputLanguage(handle, InputInterop.ruLanguage);
                }
                else
                {
                    InputInterop.ChangeInputLanguage(handle, InputInterop.enLanguage);
                }

                var position = random.Next(0, randomCharRanges.Count);
                var randomRange = randomCharRanges[position];

                InputInterop.SendButtonDownMessage((Key)random.Next(randomRange.Min, randomRange.Max));
            }

            InputInterop.SendReleaseButtonMessage(Key.LeftShift);
        }

        private void PasteTextFromClipboard()
        {
            if (!Clipboard.ContainsText()) return;

            InputInterop.SendHoldButtonMessage(Key.LeftCtrl);
            InputInterop.SendButtonDownMessage(Key.V);
            InputInterop.SendReleaseButtonMessage(Key.LeftCtrl);
        }
    }
}
