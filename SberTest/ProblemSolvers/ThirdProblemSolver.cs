using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenCvSharp;
using Shell32;
using Point = OpenCvSharp.Point;

namespace SberTest
{
    internal class ThirdProblemSolver : ISolver
    {
        private static readonly string _outputPath = "./output.txt";
        private static readonly string _screenshotPath = "./desktop_screenshot.png";
        private static readonly string _templatePath = "./Resources/template.png";

        public void Solve()
        {
            TakeDesktopScreenshot();
            var matchPoint = FindTemplateMatch();
            File.WriteAllText(_outputPath, matchPoint.ToString());

            Console.WriteLine(matchPoint.ToString());
        }

        private static Point FindTemplateMatch()
        {
            using (var template = new Mat(_templatePath))
            using (var source = new Mat(_screenshotPath))
            using (var result = new Mat())
            {
                Cv2.MatchTemplate(source, template, result, TemplateMatchModes.CCoeffNormed);
                Cv2.MinMaxLoc(result, out var minVal, out var maxVal, out Point min, out Point max);

                if (maxVal < 0.9)
                {
                    throw new ArgumentException("Не удалось найти output.txt на рабочем столе");
                }

                return max;
            };
        }

        private static void TakeDesktopScreenshot()
        {
            var desktopSize = Screen.PrimaryScreen.Bounds.Size;
            using (var bitmap = new Bitmap(desktopSize.Width, desktopSize.Height))
            using (var graphics = Graphics.FromImage(bitmap))
            {
                var shell = new Shell();
                shell.MinimizeAll();
                Task.Delay(250).Wait();
                graphics.CopyFromScreen(0, 0, 0, 0, desktopSize);
                shell.UndoMinimizeALL();
                bitmap.Save(_screenshotPath, ImageFormat.Png);
            }
        }
    }
}
