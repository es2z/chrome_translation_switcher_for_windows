using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenCvSharp;

class Program
{
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    static async Task Main() {
        string directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        string iconPath = Path.Combine(directory, "target.bmp");
        Mat iconImage = new Mat(iconPath, ImreadModes.Color);

        var iconBitmap = new Bitmap(iconPath);
        var icon = BitmapToMat(iconBitmap);

        IntPtr hWnd = IntPtr.Zero;
        RECT rect = new RECT();

        var processes = Process.GetProcessesByName("chrome");
        foreach (var process in processes) {
            if (process.MainWindowHandle != IntPtr.Zero && process.MainWindowTitle.Contains("Google Chrome")) {
                hWnd = process.MainWindowHandle;
                GetWindowRect(hWnd, out rect);
                break;
            }
        }

        if (hWnd != IntPtr.Zero) {
            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;
            var screenCapture = new Bitmap(width, height, PixelFormat.Format24bppRgb);

            using (var buf = Graphics.FromImage(screenCapture)) {
                buf.CopyFromScreen(rect.Left, rect.Top, 0, 0, new System.Drawing.Size(width, height));
            }

            var source = BitmapToMat(screenCapture);

            OpenCvSharp.Point matchPoint = await MatchTemplateAsync(source, icon);

            if (matchPoint.X != -1 && matchPoint.Y != -1) {
                System.Drawing.Point originalPosition = Cursor.Position;
                Cursor.Position = new System.Drawing.Point(matchPoint.X + rect.Left + iconImage.Width / 2, matchPoint.Y + rect.Top + iconImage.Height / 2);
                MouseClick();
                Cursor.Position = originalPosition;
                Thread.Sleep(15);
                SendKeys.SendWait("{ENTER}");
                Thread.Sleep(15);
                SendKeys.SendWait("{ESC}");
            }
        }
    }

    public static Mat BitmapToMat(Bitmap bitmap) {
        using (MemoryStream memory = new MemoryStream()) {
            bitmap.Save(memory, ImageFormat.Bmp);
            memory.Position = 0;
            Mat mat = Mat.FromImageData(memory.ToArray(), ImreadModes.Color);
            return mat;
        }
    }

    private static Task<OpenCvSharp.Point> MatchTemplateAsync(Mat source, Mat template) {
        return Task.Run(() => {
            using (Mat result = new Mat()) {
                Cv2.MatchTemplate(source, template, result, TemplateMatchModes.CCoeffNormed);
                double minVal, maxVal;
                OpenCvSharp.Point minLoc, maxLoc;
                Cv2.MinMaxLoc(result, out minVal, out maxVal, out minLoc, out maxLoc);
                if (maxVal > 0.9) return maxLoc;

                return new OpenCvSharp.Point(-1, -1);
            }
        });
    }

    private static void MouseClick() {
        uint X = (uint)Cursor.Position.X;
        uint Y = (uint)Cursor.Position.Y;

        mouse_event(MOUSEEVENTF_LEFTDOWN, X, Y, 0, 0);
        mouse_event(MOUSEEVENTF_LEFTUP, X, Y, 0, 0);
    }

    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

    private const uint MOUSEEVENTF_LEFTDOWN = 0x02;
    private const uint MOUSEEVENTF_LEFTUP = 0x04;
}
