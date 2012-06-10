#region Libraries
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Interop;
using System.Runtime.InteropServices;
using System.IO;
using Windows7.DesktopIntegration;
using Windows7.DesktopIntegration.Interop;
using System.Design;
using MediaCenter;
using System.Drawing.Drawing2D;
using Windows7.DesktopIntegration.WindowsForms;
using MC_Aero_Taskbar_Plugin.Properties;
using System.Threading;
using System.Diagnostics;
using System.Drawing.Imaging;
using MC_Aero_Taskbar_Plugin;
//using Microsoft.WindowsAPICodePack;
#endregion

namespace MC_Aero_Taskbar_Plugin
{
    #region Interop Program ID Registration

    // This string must unique and match the String in the Inno Setup Script
    [System.Runtime.InteropServices.ProgId("MCPlugin.MC_Aero_Taskbar_Plugin")]
    #endregion

    public partial class MainInterface : UserControl
    {
        #region Attributes

        private MediaCenter.MCAutomation mcRef;
        private static string AppId = "MC_Jumpbar";
        private Win32WndProc newWndProc;
        public IntPtr oldWndProc = IntPtr.Zero;
        private static Bitmap screen;
        private Rectangle windowsize;
        private bool windowMinimized = false;
        private StringBuilder oldWindowText = new StringBuilder();
        private IMJFileAutomation nowPlayingFile;
        private string playingFileImgLocation;
        private IMJPlaybackAutomation playback;
        private Settings s = new Settings();
        private static JrMainWindow jrWin;
        //private static CustomWindowsManager cwm;

        #endregion

        #region DLL Imports

        [DllImport("user32")]
        private static extern bool GetWindowRect(IntPtr hWnd, ref Rectangle rect);
        [DllImport("user32")]
        private static extern int CallWindowProc(IntPtr lpPrevWndFunc, IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32")]
        private static extern IntPtr SetWindowLong(IntPtr hWnd, int nIndex, Win32WndProc newProc);
        [DllImport("user32")]
        private static extern IntPtr SetWindowLong(IntPtr hWnd, int nIndex, IntPtr newProc);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetActiveWindow(IntPtr hWnd);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetActiveWindow();
        [DllImport("user32.dll")]
        static extern int SetWindowText(IntPtr hWnd, string text);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        // For Windows Mobile, replace user32.dll with coredll.dll
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        // Find window by Caption only. Note you must pass IntPtr.Zero as the first parameter.
        [DllImport("user32.dll", EntryPoint="FindWindow", SetLastError = true)]
        static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly, string lpWindowName);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, IntPtr windowTitle);

        #endregion

        #region Delegates
        private delegate int Win32WndProc(IntPtr hWnd, int Msg,int wParam, int lParam);
        #endregion

        #region Constructor

        public MainInterface()
        {
            InitializeComponent();
            
            // Don't run unless the taskbar extensions are supported (currently Windows 7+).
            //if (!Windows7Taskbar.)
            //{
            //    MessageBox.Show("This sample requires Windows 7 to run");
            //    Application.Exit();
            //}
            //this.mcRef = mcRef;
            //Init(mcRef);

            
        }
        #endregion

        #region Media Center Initialisation

        // Set all Components on the Panel to invisible
        private void setVisibility(Control ctlControl, Boolean blnVisible)
        {
            while (ctlControl != null)
            {
                ctlControl.Visible = blnVisible;

                ctlControl = mainPanel.GetNextControl(ctlControl, true);
            }
        }

        // Check the version of Media Center.
        // Don't support C# Plugins with versions lower than 11.1
        // Certain Events are not supported with older versions
        private Boolean checkVersion()
        {
            IMJVersionAutomation version;

            version = mcRef.GetVersion();

            if (version.Major >= 12 && version.Minor >= 0 && version.Build >= 213)
            {
                return true;
            }

            //setVisibility(Panel, false);

            Panel.Visible = true;

            return false;
        }

        public void Init(MediaCenter.MCAutomation mcRef)
        {
            try
            {
                // Add buffered graphics possibly
                
                //if (checkVersion())
                //{
                    // Tell MC to call our MJEvent Routine in case of evenst
                this.mcRef = mcRef;
                this.mcRef.FireMJEvent += new IMJAutomationEvents_FireMJEventEventHandler(MJEvent);
                // Init our plugin
                initAll();
                // This is the main entry for MC Automation
                // The application ID is used to group windows together
                //Windows7Taskbar.SetWindowAppId((IntPtr)mcRef.GetWindowHandle(), AppId);
                //Windows7Taskbar.SetCurrentProcessAppId(AppId);
                //txtUserInfo.Visible = true;
                txtUserInfo.Visible = true;
                addUserInfoText("Plugin Initiated OK");
                //cwm = CustomWindowsManager.CreateWindowsManager((IntPtr)mcRef.GetWindowHandle(), IntPtr.Zero);
                //cwm.ThumbnailRequested += new EventHandler<BitmapRequestedEventArgs>(cwm_ThumbnailRequested);
                //cwm.PeekRequested += new EventHandler<BitmapRequestedEventArgs>(cwm_PeekRequested);
                jrWin = new JrMainWindow(mcRef);
                jrWin.RequestThumbnail += new JrMainWindow.JrEventHandler(jrWin_RequestThumbnail);
                //jrWin.RequestPreview += new JrMainWindow.JrEventHandler(jrWin_RequestPreview);
                //jrWin.RequestTrackProgressUpdate += new JrMainWindow.JrEventHandler(jrWin_RequestTrackProgressUpdate);
                //StartSubclass((IntPtr)mcRef.GetWindowHandle());
                ScreenCapture sc = new ScreenCapture();
                screen = new Bitmap(sc.CaptureWindow((IntPtr)mcRef.GetWindowHandle()));
                GetWindowText((IntPtr)mcRef.GetWindowHandle(), oldWindowText, oldWindowText.Capacity);
                //backgroundWorker1.RunWorkerAsync();
                                
            }
            catch (Exception e)
            {
                exceptionHandler(e);
            }
        }

           void jrWin_RequestTrackProgressUpdate(object sender, EventArgs e)
        {
            if (trackProgress.Checked)
            {
                if (playback.State == MJPlaybackStates.PLAYSTATE_PLAYING)
                {
                    //Windows7Taskbar.SetProgressState((IntPtr)mcRef.GetWindowHandle(), Windows7Taskbar.ThumbnailProgressState.Normal);
                    jrWin.SetProgressState(Windows7Taskbar.ThumbnailProgressState.Normal);
                    //addUserInfoText("Duration: " + playback.Duration);
                    if (nowPlayingFile.Duration <= 0) jrWin.SetProgressState(Windows7Taskbar.ThumbnailProgressState.Indeterminate);
                    else
                    {
                        if (playback.Position >= 0) jrWin.SetProgressValue(playback.Position, nowPlayingFile.Duration);
                    }
                }
                else if (playback.State == MJPlaybackStates.PLAYSTATE_PAUSED) jrWin.SetProgressState(Windows7Taskbar.ThumbnailProgressState.Paused);
            }
            else if (playlistProgress.Checked)
            {
                jrWin.SetProgressState(playback.State != MJPlaybackStates.PLAYSTATE_PAUSED ? Windows7Taskbar.ThumbnailProgressState.Normal : Windows7Taskbar.ThumbnailProgressState.Paused);
                jrWin.SetProgressValue(mcRef.GetCurPlaylist().Position, mcRef.GetCurPlaylist().GetNumberFiles());
            }
            else
            {
                jrWin.SetProgressState(Windows7Taskbar.ThumbnailProgressState.NoProgress);
            }
        }

        private void jrWin_RequestPreview(object sender, EventArgs e)
        {
            setWindowsPeak();
        }

        private void jrWin_RequestThumbnail(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(playingFileImgLocation))
            {
                Image coverArt = Image.FromFile(playingFileImgLocation);
                //addUserInfoText("getting image thumbnail at: " + currentFile);
                JrThumbArgs jArgs = (JrThumbArgs)e;
                jArgs.thumb = resizeImage(coverArt, jArgs.thumbnailSize);
                //addUserInfoText(coverArtFile.ToString());
            } 

        }

        private void initAll()
        {
            try
            {
                // Set some skin colors
                setSkinColors();
            }
            catch (Exception e)
            {
                exceptionHandler(e);
            }
        }

        #endregion

        #region ErrorHandler
        private void exceptionHandler(Exception e)
        {
            addUserInfoText("A Fatal error has occured: " + e.Message);
            addUserInfoText("The Failure Occured In Class Object: " + e.Source);
            addUserInfoText("when calling Method " + e.TargetSite);
            addUserInfoText("The following Inner Exception was caused" + e.InnerException);
            addUserInfoText("The Stack Trace Follows:\r\n" + e.StackTrace);

            //txtUserInfo.Dock = DockStyle.Fill;

            this.Enabled = true;
        }
        #endregion

        #region Setting Skin Colors

        private void skinPlugin()
        {
            try
            {
                // MC 12 does not support skinning of C# plugins
                // We do some pseudo skinning by setting the colors used in the skin
                setSkinColors();
            }
            catch (Exception e)
            {
                exceptionHandler(e);
            }
        }

        // Get the Color for an item and attribute (i.e. "Tree", "BackColor"
        // Look in the wiki for the Metamorphis subject for Items and Attributes
        // http://wiki.jrmediacenter.com/index.php/Metamorphis
        private Color getColor(String strItem, String strAttribute)
        {
            Color colReturned = Color.LightGray;
            int intColor;
            int intR;
            int intG;
            int intB;

            try
            {
                intColor = mcRef.GetSkinInfo(strItem, strAttribute);

                // The color is represented as an int in MC. 
                // Windows requires a ARGB Color object
                // Using bitshifting and masking to get the R, G and B values
                if (intColor != -1)
                {
                    intR = intColor & 255;
                    intG = (intColor >> 8) & 255;
                    intB = (intColor >> 16) & 255;

                    colReturned = Color.FromArgb(intR, intG, intB);
                }
            }
            catch (Exception e)
            {
                exceptionHandler(e);
            }

            return colReturned;
        }

        // Setting the Backcolor of a Control (if different)
        private void setBackColor(Control control, Color color)
        {
            if (!control.BackColor.Equals(color))
            {
                control.BackColor = color;
            }
        }

        // Setting the Forecolor of a Control (if different)
        private void setForeColor(Control control, Color color)
        {
            if (!control.ForeColor.Equals(color))
            {
                control.ForeColor = color;
            }
        }

        // Setting the Fore- and Backcolors of all Controls (if different)
        private void setAllColors(Control control)
        {
            while (control != null)
            {
                setBackColor(control, getColor("Tree", "BackColor"));
                setForeColor(control, getColor("Tree", "TextColor"));

                control = mainPanel.GetNextControl(control, true);
            }
        }

        // Setting the Color for Maininterface
        private void setMainInterfaceColors()
        {
            this.BackColor = getColor("Tree", "BackColor");
            this.ForeColor = getColor("Tree", "TextColor");
        }

        // Pseude Skin our Plugin
        private void setSkinColors()
        {
            try
            {
                setMainInterfaceColors();
                setAllColors(Panel);
            }
            catch (Exception e)
            {
                exceptionHandler(e);
            }
        }

        #endregion

        #region Event Handling

        // Just for debugging purposes. Fill a text in the txtUserInfo Textbox
        private void addUserInfoText(String strText)
        {
            txtUserInfo.Text = strText + "\r\n" + txtUserInfo.Text;
        }
        
        // This routine is called by MC in case of an Event
        // s1 - String containing the the MC Command.
        // s2 - String containing the the MC Event
        // s3 - String containing supplement information for the Event (not always filled)
        private void MJEvent(String strCommand, String strEvent, String strEventInfo)
        {
            // Debug info
            addUserInfoText(strCommand + "/" + strEvent + "/" + strEventInfo);
            playback = mcRef.GetPlayback();
            //ScreenCapture.GrabWindowBitmap((IntPtr)mcRef.GetWindowHandle(), new Size(windowsize.Width, windowsize.Height), out screen);
            switch (strCommand)
            {
                case "MJEvent type: MCCommand":
                    switch (strEvent)
                    {
                        case "MCC: NOTIFY_TRACK_CHANGE":
                        case "MCC: NOTIFY_PLAYERSTATE_CHANGE":

                            if (playback.State != MJPlaybackStates.PLAYSTATE_STOPPED)
                            {

                                nowPlayingFile = mcRef.GetCurPlaylist().GetFile(mcRef.GetCurPlaylist().Position);
                                playingFileImgLocation = nowPlayingFile.GetImageFile(MJImageFileFlags.IMAGEFILE_THUMBNAIL_MEDIUM);

                                if (enableCoverArt.Checked)
                                {
                                    jrWin.EnableCustomWindowPreview();
                                    setThumbnail(playingFileImgLocation);
                                }
                            }
                            else jrWin.DisableCustomWindowPreview();
                            //backgroundWorker1.RunWorkerAsync();
                            break;

                        case "MCC: NOTIFY_PLAYLIST_ADDED":
                            // Your code goes here
                            break;

                        case "MCC: NOTIFY_PLAYLIST_INFO_CHANGED":
                            // Your code goes here
                            break;

                        case "MCC: NOTIFY_PLAYLIST_FILES_CHANGED":
                            // Your code goes here
                            break;

                        case "MCC: NOTIFY_PLAYLIST_REMOVED":
                            // Your code goes here
                            break;

                        case "MCC: NOTIFY_PLAYLIST_COLLECTION_CHANGED":
                            // Your code goes here
                            break;

                        case "MCC: NOTIFY_PLAYLIST_PROPERTIES_CHANGED":
                            // Your code goes here
                            break;

                        case "MCC: NOTIFY_SKIN_CHANGED":
                            skinPlugin();
                            break;
                    }

                    break;

                default:
                    break;
            }
        }


 /*           public void MCWin7Taskbar()
            {
                IMJPlaybackAutomation playback;
                playback = mcRef.GetPlayback();
                //this.mcRef = mcRef;
                //this.mcRef.FireMJEvent += new IMJAutomationEvents_FireMJEventEventHandler(MJEvent);
                try
                {

                          
                }
                catch (Exception e)
                {
                    errorHandler(e);
                }
            }*/

        // At the time this template was created the MCC: NOTIFY_SKIN_CHANGED
        // Event didn't function correctly. Doing it on a PAINT Event from Windows.
        private void mainPanel_Paint(object sender, PaintEventArgs e)
        {
            skinPlugin();
        }

        #endregion

        //#region Window Proc stuff
        //private int MyWndProc(IntPtr hWnd, int Msg, int wParam, int lParam)
        //{
        //    //if (enableCoverArt.Checked && mcRef.GetPlayback().State != MJPlaybackStates.PLAYSTATE_STOPPED) Windows7Taskbar.EnableCustomWindowPreview((IntPtr)mcRef.GetWindowHandle());
        //    //else Windows7Taskbar.DisableCustomWindowPreview((IntPtr)mcRef.GetWindowHandle());
        //    //addUserInfoText(Msg.ToString());
        //    switch (Msg)
        //    {
        //        case WM_SYSCOMMAND:
        //            addUserInfoText("System Command");
        //            switch (wParam)
        //            {
        //                case SC_MINIMIZE:
        //                    windowMinimized = true;
        //                    break;
        //                case SC_RESTORE:
        //                    windowMinimized = false;
        //                    break;
        //            }
        //            break;
        //        case WM_DWMSENDICONICLIVEPREVIEWBITMAP:
        //            setWindowsPeak();
        //            return 0;
        //            //break;
        //        case WM_DWMSENDICONICTHUMBNAIL:
        //            string imageFileName = nowPlayingFile.GetImageFile(MJImageFileFlags.IMAGEFILE_THUMBNAIL_MEDIUM);
        //            //addUserInfoText(imageFileName);
        //            int width = (int)((long)lParam >> 16);
        //            int height = (int)(((long)lParam) & (0xFFFF));
        //            thumbnailSize = new Size(width, height);
        //            addUserInfoText(thumbnailSize.ToString());
        //            if (!string.IsNullOrEmpty(imageFileName)) setPreview(imageFileName);
        //            //imageFileName = "";
        //            return 0;
        //            //break;
        //        default:
        //            //generateWindowsPeak = false;
        //            break;
        //    }
        //    return CallWindowProc(oldWndProc, hWnd, Msg, wParam, lParam);
        //}
        //#endregion

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void txtUserInfo_TextChanged(object sender, EventArgs e)
        {

        }

        private void setWindowTitle(IMJFileAutomation nowPlayingFile)
        {

            IMJPlaybackAutomation playback = mcRef.GetPlayback();

            if (playback.State != MJPlaybackStates.PLAYSTATE_STOPPED && displayArtistTrackName.Checked)
            {
                Windows7Taskbar.SetThumbnailTooltip((IntPtr)mcRef.GetWindowHandle(), nowPlayingFile.Artist + " - " + nowPlayingFile.Name);
                SetWindowText((IntPtr)mcRef.GetWindowHandle(), nowPlayingFile.Artist + " - " + nowPlayingFile.Name);
            }
            else
            {
                SetWindowText((IntPtr)mcRef.GetWindowHandle(), oldWindowText.ToString());
                Windows7Taskbar.SetThumbnailTooltip((IntPtr)mcRef.GetWindowHandle(), oldWindowText.ToString());
            }
        }

        private IntPtr getDisplayWindowHandle()
        {
            IntPtr displayHandle = FindWindow("J. River Display Container Window", "Display");
            if (displayHandle == IntPtr.Zero)
            {
                IntPtr MainWindow = FindWindowEx((IntPtr)mcRef.GetWindowHandle(), IntPtr.Zero, "MainUIWnd", IntPtr.Zero);
                IntPtr MCViewContainer = FindWindowEx(MainWindow, IntPtr.Zero, "MC View Container", IntPtr.Zero);
                IntPtr displayContainer = FindWindowEx(MCViewContainer, IntPtr.Zero, "J. River Display Window", IntPtr.Zero);
                displayHandle = FindWindowEx(displayContainer, IntPtr.Zero, "J. River Display Container Window", IntPtr.Zero);
            }

            return displayHandle;
        }
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            playback = mcRef.GetPlayback();
            //customWindow = Windows7.DesktopIntegration.CustomWindowsManager.CreateWindowsManager((IntPtr)mcRef.GetWindowHandle(), IntPtr.Zero);
            //IMJFileAutomation nowPlayingFile;
            //nowPlayingFile = mcRef.GetCurPlaylist().GetFile(mcRef.GetCurPlaylist().Position);

            //jrWin = new JrMainWindow((IntPtr)mcRef.GetWindowHandle());

            string prevFileName = "";
            
            try
            {
                while (playback.State != MJPlaybackStates.PLAYSTATE_STOPPED)
                {
                    
                    //addUserInfoText("Playback state: " + playback.State);

                    if (enableCoverArt.Checked) jrWin.EnableCustomWindowPreview();
                    
                    if (trackProgress.Checked)
                    {
                        if (playback.State == MJPlaybackStates.PLAYSTATE_PLAYING)
                        {
                            //Windows7Taskbar.SetProgressState((IntPtr)mcRef.GetWindowHandle(), Windows7Taskbar.ThumbnailProgressState.Normal);
                            jrWin.SetProgressState(Windows7Taskbar.ThumbnailProgressState.Normal);
                            //addUserInfoText("Duration: " + playback.Duration);
                            if (nowPlayingFile.Duration <= 0) jrWin.SetProgressState(Windows7Taskbar.ThumbnailProgressState.Indeterminate);
                            else
                            {
                                if (playback.Position >= 0) jrWin.SetProgressValue(playback.Position, nowPlayingFile.Duration);
                            }
                        }
                        else if (playback.State == MJPlaybackStates.PLAYSTATE_PAUSED) jrWin.SetProgressState(Windows7Taskbar.ThumbnailProgressState.Paused);
                    }

                    else if (playlistProgress.Checked)
                    {
                        jrWin.SetProgressState(playback.State != MJPlaybackStates.PLAYSTATE_PAUSED ? Windows7Taskbar.ThumbnailProgressState.Normal : Windows7Taskbar.ThumbnailProgressState.Paused);
                        jrWin.SetProgressValue(mcRef.GetCurPlaylist().Position, mcRef.GetCurPlaylist().GetNumberFiles());
                    }

                    else
                    {
                        jrWin.SetProgressState(Windows7Taskbar.ThumbnailProgressState.NoProgress);
                    }
                    //setWindowTitle(nowPlayingFile);
                    System.Threading.Thread.Sleep(250);
                    
                }
                jrWin.DisableCustomWindowPreview();
                jrWin.SetProgressState(Windows7Taskbar.ThumbnailProgressState.NoProgress);
            }
            catch (Exception ex)
            {
                exceptionHandler(ex);
            }
            return; 
        }

        private void setWindowsPeak()
        {
            //if (!windowMinimized)
            //{
            //    GetWindowRect((IntPtr)mcRef.GetWindowHandle(), ref windowsize);
            //    ScreenCapture sc = new ScreenCapture();
            //    screen.Dispose();
            //    screen = null;
            //    screen = new Bitmap(sc.CaptureWindow((IntPtr)mcRef.GetWindowHandle()));
            //}
            
            if (!windowMinimized)
            {
                try
                {
                    //ScreenCapture sc = new ScreenCapture();
                    //IntPtr hb = sc.GetHbitMapWindow((IntPtr)mcRef.GetWindowHandle());
                    Windows7Taskbar.SetPeekBitmap((IntPtr)mcRef.GetWindowHandle(), false);
                    //sc.DeleteHbitmapObject(hb);
                }
                catch (Exception e)
                {
                    addUserInfoText(e.Message);
                }
            }
        }

        [DllImport("dwmapi.dll", PreserveSig = false)]
        public static extern void DwmQueryThumbnailSourceSize(IntPtr hThumbnail, out Size size);

        private void setThumbnail(string currentFile)
        {
            setThumbnail(currentFile, new Size(120, 120));
        }

        private void setThumbnail(string currentFile, Size thumbnailSize)
        {
            try
            {
                if (!string.IsNullOrEmpty(currentFile))
                {
                    Image coverArt = Image.FromFile(currentFile);                    
                    //addUserInfoText("getting image thumbnail at: " + currentFile);
                    Bitmap coverArtFile = resizeImage(coverArt, thumbnailSize);
                    //addUserInfoText(coverArtFile.ToString());
                    Windows7Taskbar.SetIconicThumbnail((IntPtr)mcRef.GetWindowHandle(), coverArtFile);
                }
            }
            catch (Exception ex)
            {
                exceptionHandler(ex);
            }
        }
        private static Bitmap resizeImage(Image imgToResize, Size size)
        {
            int sourceWidth = imgToResize.Width;
            int sourceHeight = imgToResize.Height;

            float nPercent = 0;
            float nPercentW = 0;
            float nPercentH = 0;

            nPercentW = ((float)size.Width / (float)sourceWidth);
            nPercentH = ((float)size.Height / (float)sourceHeight);

            if (nPercentH < nPercentW)
                nPercent = nPercentH;
            else
                nPercent = nPercentW;

            int destWidth = (int)(sourceWidth * nPercent);
            int destHeight = (int)(sourceHeight * nPercent);

            Bitmap b = new Bitmap(destWidth, destHeight);
            Graphics g = Graphics.FromImage((Image)b);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            g.DrawImage(imgToResize, 0, 0, destWidth, destHeight);
            g.Dispose();

            return b;
        }

        private void Panel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {
            s.enableCoverArt = ((CheckBox)sender).Checked;
            s.Save();
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            base.OnHandleDestroyed(e);
        }

        protected override void OnLoad(EventArgs e)
        {
            enableCoverArt.Checked = s.enableCoverArt;
            displayArtistTrackName.Checked = s.displayArtistTrackName;
            trackProgress.Checked = s.trackProgress;
            playlistProgress.Checked = s.playlistProgress;
            noProgressTrack.Checked = s.noProgressTrack;
            base.OnLoad(e);
        }

        //[System.Security.Permissions.PermissionSet(System.Security.Permissions.SecurityAction.Demand, Name = "FullTrust")]
        //protected override void WndProc(ref Message m)
        //{
        //    // Listen for operating system messages.
        //    switch (m.Msg)
        //    {
        //        case WM_DESTROY:
        //            StopSubclass((IntPtr)mcRef.GetWindowHandle());
        //        break;
        //    }
        //    base.WndProc(ref m);
        //}

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            s.displayArtistTrackName = ((CheckBox)sender).Checked;
            s.Save();
        }

        private void trackProgress_CheckedChanged(object sender, EventArgs e)
        {
            s.trackProgress = ((RadioButton)sender).Checked;
            s.Save();
        }

        private void playlistProgress_CheckedChanged(object sender, EventArgs e)
        {
            s.playlistProgress = ((RadioButton)sender).Checked;
            s.Save();
        }

        private void noProgressTrack_CheckedChanged(object sender, EventArgs e)
        {
            s.noProgressTrack = ((RadioButton)sender).Checked;
            s.Save();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }

    class JrMainWindow : NativeWindow
    {
        #region Attributes
        private CustomWindowsManager cwm;
        private Bitmap coverArtImage;
        #endregion

        #region Constants
        public const int WM_DWMSENDICONICTHUMBNAIL = 0x0323;
        public const int WM_DWMSENDICONICLIVEPREVIEWBITMAP = 0x0326;
        public const int WM_SYSCOMMAND = 0x0112;
        public const int SC_MINIMIZE = 0xF020;
        public const int SC_RESTORE = 0xF120;
        private const int GWL_WNDPROC = -4;
        private const int WM_DESTROY = 0x0002;
        #endregion

        #region delegates
        public delegate void JrEventHandler(object sender, EventArgs e);
        #endregion

        #region Events
        public event JrEventHandler RequestThumbnail;
        public event JrEventHandler RequestPreview;
        public event JrEventHandler RequestTrackProgressUpdate;
        #endregion

        public JrMainWindow(MCAutomation mcRef)
            : base()
        {
            this.AssignHandle((IntPtr)mcRef.GetWindowHandle());
            cwm = CustomWindowsManager.CreateWindowsManager(this.Handle);
            cwm.PeekRequested += new EventHandler<BitmapRequestedEventArgs>(cwm_PeekRequested);
            cwm.ThumbnailRequested += new EventHandler<BitmapRequestedEventArgs>(cwm_ThumbnailRequested);
        }

        protected override void WndProc(ref Message m)
        {

            switch (m.Msg)
            {
                case WM_SYSCOMMAND:
                    //switch (wParam)
                    //{
                    //}
                    break;
                //case WM_DWMSENDICONICLIVEPREVIEWBITMAP:
                //    if (RequestPreview != null)
                //        RequestPreview(this, EventArgs.Empty);
                //    break;
                case WM_DWMSENDICONICTHUMBNAIL:
                    int width = (int)((long)m.LParam >> 16);
                    int height = (int)(((long)m.LParam) & (0xFFFF));
                    if (RequestThumbnail != null)
                    {
                        JrThumbArgs args = new JrThumbArgs(new Size(width, height));
                        RequestThumbnail(this, args);
                        coverArtImage = args.thumb;
                    }
                    break;
                default:
                    break;
            }
            
            if (cwm != null)
            {
                cwm.EnablePreview();
                if (coverArtImage == null) cwm.DisablePreview();
                cwm.DispatchMessage(ref m);
                cwm.InvalidatePreviews();
            }
            //if (RequestTrackProgressUpdate != null)
            //    RequestTrackProgressUpdate(this, EventArgs.Empty);
            base.WndProc(ref m);
        }

        #region Windows 7 Thumbnail Wrapper Classes
        public void EnableCustomWindowPreview()
        {
            Windows7Taskbar.EnableCustomWindowPreview(this.Handle);
        }

        public void DisableCustomWindowPreview()
        {
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

        void cwm_PeekRequested(object sender, BitmapRequestedEventArgs e)
        {
            e.DisplayFrameAroundBitmap = false;
            e.DoNotMirrorBitmap = false;
            e.UseWindowScreenshot = true;
        }

        void cwm_ThumbnailRequested(object sender, BitmapRequestedEventArgs e)
        {
            e.DoNotMirrorBitmap = true;
            if (coverArtImage != null)
                e.Bitmap = coverArtImage;
        }

        #endregion
    }

    class JrThumbArgs : EventArgs
    {
        public readonly Size thumbnailSize;
        public Bitmap thumb;

        public JrThumbArgs(Size thumbnailSize)
        {
            this.thumbnailSize = thumbnailSize;
        }

    }
}