using System;
using System.Windows.Forms;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using Windows7.DesktopIntegration;
using Windows7.DesktopIntegration.Interop;

namespace MC_Aero_Taskbar_Plugin
{
	/// <summary>
	/// Description of JrMainWindow.
	/// </summary>
	public class JrMainWindow : NativeWindow
    {
        #region Attributes
        private static Bitmap WindowsPeak;
        private bool IsMinimized = false;
        private bool PreviewEnabled = false;
        private ScreenCapture sc = new ScreenCapture();
        private bool CaptureWindow = false;
        #endregion

        #region Constants
        private const int WM_DWMSENDICONICTHUMBNAIL = 0x0323;
        private const int WM_DWMSENDICONICLIVEPREVIEWBITMAP = 0x0326;
        private const int WM_SYSCOMMAND = 0x0112;
        private const int SC_MINIMIZE = 0xF020;
        private const int SC_RESTORE = 0xF120;
        private const int GWL_WNDPROC = -4;
        private const int WM_DESTROY = 0x0002;
        private const UInt32 WS_MINIMIZE = 0x20000000;
        private const int GWL_STYLE = (-16);
        private const int WM_SIZE = 0x0005;
        private const int WM_MOVE = 0x0003;
        private const int WM_EXITSIZEMOVE = 0x0232;
        private const int WM_CLOSE = 0x0010;
        private const int SC_CLOSE = 0xF060;
        #endregion

        #region delegates
        public delegate void JrEventHandler(object sender, EventArgs e);
        #endregion

        #region Events
        public event JrEventHandler RequestThumbnail;
        public event JrEventHandler RequestTrackProgressUpdate;
        public event JrEventHandler WindowClosing;
        #endregion

        #region Pinvokes
        [DllImport("user32.dll")]
        private static extern int SetWindowText(IntPtr hWnd, string text);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("dwmapi.dll")]
        private static extern int DwmInvalidateIconicBitmaps(IntPtr hwnd);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        #endregion

        public JrMainWindow(IntPtr handle)
            : base()
        {
            this.AssignHandle(handle);
            int lStyles = GetWindowLong(this.Handle, GWL_STYLE);
            if ((GetWindowLong(this.Handle, GWL_STYLE) & WS_MINIMIZE) == 0)
                WindowsPeak = (Bitmap)sc.CaptureWindow(this.Handle);
        }

        protected override void WndProc(ref Message m)
        {
            if (CaptureWindow && PreviewEnabled)
            {
                if (WindowsPeak != null) WindowsPeak.Dispose();
                WindowsPeak = (Bitmap)sc.CaptureWindow(this.Handle);
                CaptureWindow = false;
            }
            switch (m.Msg)
            {
                case (WM_SYSCOMMAND):
                    if ((int)m.WParam == SC_CLOSE && WindowClosing != null)
                    {
                        WindowClosing(this, EventArgs.Empty);
                        break;
                    }

                    IsMinimized = (int)m.WParam == SC_MINIMIZE;

                    if (PreviewEnabled)
                    {
                        if (WindowsPeak != null) WindowsPeak.Dispose();
                        WindowsPeak = (Bitmap)sc.CaptureWindow(this.Handle);
                    }
                    break;
                case WM_DWMSENDICONICLIVEPREVIEWBITMAP:
                    if (!IsMinimized)
                    {
                        if (WindowsPeak != null) WindowsPeak.Dispose();
                        WindowsPeak = (Bitmap)sc.CaptureWindow(this.Handle);
                    }
                    Windows7Taskbar.SetPeekBitmap(this.Handle, WindowsPeak, false);
                    break;
                case WM_DWMSENDICONICTHUMBNAIL:
                    if (RequestThumbnail != null)
                    {
                        int width = (int)((long)m.LParam >> 16);
                        int height = (int)(((long)m.LParam) & (0xFFFF));
                        JrThumbArgs args = new JrThumbArgs(new Size(width, height));
                        RequestThumbnail(this, args);
                        SetThumbnailPreview(args.thumbBmp);
                        args.thumbBmp.Dispose();
                    }
                    break;
                case WM_CLOSE:
                    if (WindowClosing != null)
                        WindowClosing(this, EventArgs.Empty);
                    break;
                default:
                    break;
            }

            DwmInvalidateIconicBitmaps(this.Handle);

            if (RequestTrackProgressUpdate != null)
                RequestTrackProgressUpdate(this, EventArgs.Empty);
            base.WndProc(ref m);
        }

        

        #region Windows 7 Thumbnail Wrapper Classes
        public void EnableCustomWindowPreview()
        {
            PreviewEnabled = true;
            Windows7Taskbar.EnableCustomWindowPreview(this.Handle);
        }

        public void DisableCustomWindowPreview()
        {
            PreviewEnabled = false;
            Windows7Taskbar.DisableCustomWindowPreview(this.Handle);
        }

        public void SetProgressState(Windows7Taskbar.ThumbnailProgressState progressState)
        {
            Windows7Taskbar.SetProgressState(this.Handle, progressState);
        }

        public void SetProgressValue(int current, int maximum)
        {
            Windows7Taskbar.SetProgressValue(this.Handle, (ulong)current, (ulong)maximum);
        }

        public void SetWindowTitle(string title)
        {
            SetWindowText(this.Handle, title);
        }

        public string GetWindowTitle()
        {
            StringBuilder sb = new StringBuilder();
            GetWindowText(this.Handle, sb, sb.Capacity);
            return sb.ToString();
        }

        public void SetThumbnailPreview(Bitmap thumbBmp)
        {
            Windows7Taskbar.SetIconicThumbnail(this.Handle, thumbBmp);
        }

        private void cwm_PeekRequested(object sender, BitmapRequestedEventArgs e)
        {
            e.DisplayFrameAroundBitmap = false;
            e.DoNotMirrorBitmap = false;
            e.UseWindowScreenshot = true;
        }

        private void cwm_ThumbnailRequested(object sender, BitmapRequestedEventArgs e)
        {
            e.DoNotMirrorBitmap = true;
            if (RequestThumbnail != null)
            {
                JrThumbArgs args = new JrThumbArgs(new Size(e.Width, e.Height));
                RequestThumbnail(this, args);
                e.Bitmap = args.thumbBmp;
            }
        }

        #endregion
    }
	
	public class JrThumbArgs : EventArgs
    {
        public readonly Size thumbnailSize;
        public Bitmap thumbBmp;

        public JrThumbArgs(Size thumbnailSize)
        {
            this.thumbnailSize = thumbnailSize;
        }
    }
}
