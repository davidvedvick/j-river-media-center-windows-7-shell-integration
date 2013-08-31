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
using MediaCenter;
using System.Drawing.Drawing2D;
using Windows7.DesktopIntegration.WindowsForms;
using MC_Aero_Taskbar_Plugin.Properties;
using System.Threading;
using System.Diagnostics;
using System.Drawing.Imaging;
using MC_Aero_Taskbar_Plugin;
using Microsoft.WindowsAPICodePack.Taskbar;
using System.Reflection;
using System.Runtime.Serialization;
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
        private static string appId = "MC_Jumpbar";
        public IntPtr oldWndProc = IntPtr.Zero;
        private string oldWindowText;
        private IMJFileAutomation nowPlayingFile;
        private string playingFileImgLocation;
        private IMJPlaybackAutomation playback;
        private Settings settings = new Settings();
        private AppSettings<McAeroTaskbarSettings> appSettings;
        private static JrMainWindow jrWin;
        private static JumpList jumpList;
        private string exePath;
        private const string RECENTLY_IMPORTED = "Recently Imported";
        private const string TOP_HITS = "Top Hits";
        private JumpListCustomCategory playlistCategory;
        //private static CustomWindowsManager cwm;

        #endregion

        #region DLL Imports

        [DllImport("user32")]
        private static extern bool GetWindowRect(IntPtr hWnd, ref Rectangle rect);
        [DllImport("user32")]
        private static extern int CallWindowProc(IntPtr lpPrevWndFunc, IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32")]
        private static extern IntPtr SetWindowLong(IntPtr hWnd, int nIndex, IntPtr newProc);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetActiveWindow(IntPtr hWnd);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetActiveWindow();
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        // For Windows Mobile, replace user32.dll with coredll.dll
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        // Find window by Caption only. Note you must pass IntPtr.Zero as the first parameter.
        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly, string lpWindowName);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, IntPtr windowTitle);

        #endregion

        #region Constructor

        public MainInterface()
        {
            InitializeComponent();

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
                playback = mcRef.GetPlayback();
                // Init our plugin
                initAll();

                

                // This is the main entry for MC Automation
                // The application ID is used to group windows together
                txtUserInfo.Visible = true;
                addUserInfoText("Plugin Initiated OK");

                exePath = Environment.CurrentDirectory;
                
                jrWin = new JrMainWindow((IntPtr)mcRef.GetWindowHandle());
                jrWin.RequestThumbnail += new JrMainWindow.JrEventHandler(jrWin_RequestThumbnail);
                jrWin.RequestTrackProgressUpdate += new JrMainWindow.JrEventHandler(jrWin_RequestTrackProgressUpdate);
                jrWin.WindowClosing += (ss, ee) =>
                {
                    playback = null;
                    this.mcRef = null;
                    this.Dispose();
                };
                oldWindowText = jrWin.GetWindowTitle();
                setWindowsPreview();

                jumpList = JumpList.CreateJumpList();
                playlistCategory = new JumpListCustomCategory("Playlists");
                jumpList.ClearAllUserTasks();
                jumpList.AddCustomCategories(playlistCategory);

                string appSettingsFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "mc_aero_taskbar_plugin");
                if (!Directory.Exists(appSettingsFile)) Directory.CreateDirectory(appSettingsFile);
                appSettingsFile = Path.Combine(appSettingsFile, "settings.jsn");
                appSettings = new AppSettings<McAeroTaskbarSettings>(appSettingsFile);

                if (!File.Exists(appSettingsFile))
                {
                    appSettings.Settings.DisplayArtistTrackName = settings.displayArtistTrackName;
                    appSettings.Settings.EnableCoverArt = settings.enableCoverArt;
                    appSettings.Settings.NoProgressTrack = settings.noProgressTrack;
                    appSettings.Settings.PlaylistProgress = settings.playlistProgress;
                    appSettings.Settings.TrackProgress = settings.trackProgress;
                    appSettings.Save();
                }

                enableCoverArt.Checked = appSettings.Settings.EnableCoverArt;
                displayArtistTrackName.Checked = appSettings.Settings.DisplayArtistTrackName;
                trackProgress.Checked = appSettings.Settings.TrackProgress;
                playlistProgress.Checked = appSettings.Settings.PlaylistProgress;
                noProgressTrack.Checked = appSettings.Settings.NoProgressTrack;
                BuildPlaylistTree(mcRef.GetPlaylists());
            }
            catch (Exception e)
            {
                exceptionHandler(e);
            }
        }

        void m_jumpList_UserRemovedItems(object sender, UserRemovedItemsEventArgs e)
        {
            
        }

        private void BuildPlaylistTree(IMJPlaylistsAutomation playlists)
        {
            tvPlaylists.Nodes.Clear();
            TreeNode newNode;

            IMJPlaylistAutomation currentPlaylist = playlists.GetPlaylist(-2);
            newNode = tvPlaylists.Nodes.Add(RECENTLY_IMPORTED);
            newNode.Tag = RECENTLY_IMPORTED;
            newNode.Checked = appSettings.Settings.PinnedPlaylists.Contains(RECENTLY_IMPORTED);

            currentPlaylist = playlists.GetPlaylist(-1);
            newNode = tvPlaylists.Nodes.Add(TOP_HITS);
            newNode.Tag = TOP_HITS;
            newNode.Checked = appSettings.Settings.PinnedPlaylists.Contains(TOP_HITS);

            for (int i = 0; i < playlists.GetNumberPlaylists(); i++)
            {
                currentPlaylist = playlists.GetPlaylist(i);
                newNode = String.IsNullOrEmpty(currentPlaylist.Path) ? tvPlaylists.Nodes.AddUniqueNode(currentPlaylist.Name) : GetFolderNode(tvPlaylists.Nodes, currentPlaylist.Path).Nodes.AddUniqueNode(currentPlaylist.Name);
                newNode.Tag = currentPlaylist.Path + "\\" + currentPlaylist.Name;
                newNode.Checked = appSettings.Settings.PinnedPlaylists.Contains(newNode.Tag.ToString());
            }
        }

        private TreeNode GetFolderNode(TreeNodeCollection nodes, string path)
        {
            TreeNode returnNode = null;

            string[] findPath = path.Split(new char[] { '\\' }, 2);

            returnNode = nodes.AddUniqueNode(findPath[0]);
            if (findPath[0] != path) return GetFolderNode(returnNode.Nodes, findPath[1]);

            return returnNode;
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

                control = this.GetNextControl(control, true);
            }
        }

        // Setting the Color for Maininterface
        private void setMainInterfaceColors()
        {
            this.BackColor = getColor("Tree", "BackColor");
            this.ForeColor = getColor("Tree", "Text");
        }

        // Pseude Skin our Plugin
        private void setSkinColors()
        {
            try
            {
                setMainInterfaceColors();
                setAllColors(this.Panel);
            }
            catch (Exception e)
            {
                exceptionHandler(e);
            }
        }

        #endregion

        #region Event Handling
        #region JR MC Event Handlers
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

            //ScreenCapture.GrabWindowBitmap((IntPtr)mcRef.GetWindowHandle(), new Size(windowsize.Width, windowsize.Height), out screen);
            switch (strCommand)
            {
                case "MJEvent type: MCCommand":
                    switch (strEvent)
                    {
                        case "MCC: NOTIFY_TRACK_CHANGE":
                        case "MCC: NOTIFY_PLAYERSTATE_CHANGE":
                            setWindowsPreview();

                            //backgroundWorker1.RunWorkerAsync();
                            break;

                        case "MCC: NOTIFY_PLAYLIST_ADDED":
                        case "MCC: NOTIFY_PLAYLIST_INFO_CHANGED":
                        case "MCC: NOTIFY_PLAYLIST_FILES_CHANGED":
                        case "MCC: NOTIFY_PLAYLIST_REMOVED":
                        case "MCC: NOTIFY_PLAYLIST_COLLECTION_CHANGED":
                        case "MCC: NOTIFY_PLAYLIST_PROPERTIES_CHANGED":
                            BuildPlaylistTree(mcRef.GetPlaylists());
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

        // At the time this template was created the MCC: NOTIFY_SKIN_CHANGED
        // Event didn't function correctly. Doing it on a PAINT Event from Windows.
        private void mainPanel_Paint(object sender, PaintEventArgs e)
        {
            skinPlugin();
        }
        #endregion

        #region Plug-in Form Handlers
        
        protected override void OnHandleDestroyed(EventArgs e)
        {
            base.OnHandleDestroyed(e);
        }

        protected override void OnLoad(EventArgs e)
        {
            
            

            base.OnLoad(e);
        }

        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {
            appSettings.Settings.EnableCoverArt = ((CheckBox)sender).Checked;
            appSettings.Save();
            setWindowsPreview();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            appSettings.Settings.DisplayArtistTrackName = ((CheckBox)sender).Checked;
            appSettings.Save();
        }

        private void trackProgress_CheckedChanged(object sender, EventArgs e)
        {
            appSettings.Settings.TrackProgress = ((RadioButton)sender).Checked;
            appSettings.Save();
        }

        private void playlistProgress_CheckedChanged(object sender, EventArgs e)
        {
            appSettings.Settings.PlaylistProgress = ((RadioButton)sender).Checked;
            appSettings.Save();
        }

        private void noProgressTrack_CheckedChanged(object sender, EventArgs e)
        {
            appSettings.Settings.NoProgressTrack = ((RadioButton)sender).Checked;
            appSettings.Save();
        }

        private void tvPlaylists_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Checked)
            {
                appSettings.Settings.PinnedPlaylists.Add(e.Node.Tag.ToString());
            }
            else
            {
                appSettings.Settings.PinnedPlaylists.Remove(e.Node.Tag.ToString());
            }
            appSettings.Save();
            refreshUserTasks();
        }

        private void refreshUserTasks()
        {
            jumpList.ClearAllUserTasks();
            playlistCategory.ClearJumplist();
            foreach (string playlistPath in appSettings.Settings.PinnedPlaylists)
            {
                string link = "MC" + mcRef.GetVersion().Major.ToString() + ".exe";
                JumpListLink item = new JumpListLink(link, playlistPath.Remove(0, playlistPath.LastIndexOf('\\') + 1));
                item.Arguments = "/Play TREEPATH=\"Playlists\\" + playlistPath.TrimStart('\\') + "\"";
                playlistCategory.AddJumpListItems(item);
            }
            
            jumpList.Refresh();
        }
        #endregion

        #region MC Window Handlers
        private void jrWin_RequestTrackProgressUpdate(object sender, EventArgs e)
        {
            try
            {
                if (nowPlayingFile == null || playback == null) return;

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
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex);
            }
        }

        private void jrWin_RequestThumbnail(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(playingFileImgLocation))
            {
                Image coverArt = Image.FromFile(playingFileImgLocation);
                //addUserInfoText("getting image thumbnail at: " + currentFile);
                JrThumbArgs jArgs = (JrThumbArgs)e;
                jArgs.thumbBmp = resizeImage(coverArt, jArgs.thumbnailSize);
                //addUserInfoText(coverArtFile.ToString());
            }

        }
        #endregion

        #endregion

        #region Miscellaneous Functions

        private void setWindowsPreview()
        {
            if (playback == null) return;
            switch (playback.State)
            {
                case MJPlaybackStates.PLAYSTATE_STOPPED:
                    jrWin.DisableCustomWindowPreview();
                    jrWin.SetProgressState(Windows7Taskbar.ThumbnailProgressState.NoProgress);
                    break;
                case MJPlaybackStates.PLAYSTATE_PAUSED:
                case MJPlaybackStates.PLAYSTATE_PLAYING:
                    jrWin.SetProgressState(playback.State == MJPlaybackStates.PLAYSTATE_PLAYING ? Windows7Taskbar.ThumbnailProgressState.Normal : Windows7Taskbar.ThumbnailProgressState.Paused);
                    nowPlayingFile = mcRef.GetCurPlaylist().GetFile(mcRef.GetCurPlaylist().Position);
                    playingFileImgLocation = nowPlayingFile.GetImageFile(MJImageFileFlags.IMAGEFILE_THUMBNAIL_MEDIUM);
                    if (string.IsNullOrEmpty(playingFileImgLocation)) playingFileImgLocation = nowPlayingFile.GetImageFile(MJImageFileFlags.IMAGEFILE_DISPLAY);

                    if (enableCoverArt.Checked)
                    {
                        jrWin.EnableCustomWindowPreview();
                        setThumbnail(playingFileImgLocation);
                    }

                    jrWin.SetWindowTitle(displayArtistTrackName.Checked ? (nowPlayingFile.Artist + " - " + nowPlayingFile.Name) : oldWindowText);
                    break;
                case MJPlaybackStates.PLAYSTATE_WAITING:
                    jrWin.SetProgressValue(0, 1);
                    jrWin.SetProgressState(Windows7Taskbar.ThumbnailProgressState.Indeterminate);
                    nowPlayingFile = mcRef.GetCurPlaylist().GetFile(mcRef.GetCurPlaylist().Position);
                    jrWin.SetWindowTitle(displayArtistTrackName.Checked ? (nowPlayingFile.Artist + " - " + nowPlayingFile.Name) : oldWindowText);
                    break;
            }
            
        }

        [DllImport("dwmapi.dll", PreserveSig = false)]
        public static extern void DwmQueryThumbnailSourceSize(IntPtr hThumbnail, out Size size);
        private void setThumbnail(string currentFile)
        {
            // Can't imagine a thumnbail having much smaller size than this.
            setThumbnail(currentFile, new Size(1, 1));
        }

        private void setThumbnail(string currentFile, Size thumbnailSize)
        {
            try
            {
                if (!string.IsNullOrEmpty(currentFile))
                {
                    using (Image coverArt = Image.FromFile(currentFile))
                    using (Bitmap resizedImg = resizeImage(coverArt, thumbnailSize))
                        jrWin.SetThumbnailPreview(resizedImg);
                }
            }
            catch (Exception ex)
            {
                exceptionHandler(ex);
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

        #endregion

        [DataContract]
        private class McAeroTaskbarSettings
        {
            [DataMember]
            public bool EnableCoverArt;
            [DataMember]
            public bool DisplayArtistTrackName;
            [DataMember]
            public bool TrackProgress;
            [DataMember]
            public bool PlaylistProgress;
            [DataMember]
            public bool NoProgressTrack;
            [DataMember]
            private HashSet<string> _pinnedPlaylists;

            public HashSet<string> PinnedPlaylists
            {
                get
                {
                    if (_pinnedPlaylists == null) _pinnedPlaylists = new HashSet<string>(StringComparer.Ordinal);
                    return _pinnedPlaylists;
                }
                set
                {
                    _pinnedPlaylists = value;
                }
            }
        }
    }

    public static class TreeNodeExtensions
    {
        public static TreeNode AddUniqueNode(this TreeNodeCollection nodes, string name)
        {
            foreach (TreeNode node in nodes)
            {
                if (node.Text == name)
                    return node;
            }

            return nodes.Add(name);
        }
    }
}