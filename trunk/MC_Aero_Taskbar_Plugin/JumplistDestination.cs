using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Windows7.DesktopIntegration;

namespace MC_Aero_Taskbar_Plugin
{
    public class PlaylistJlDestination : IJumpListDestination
    {
        private const string _category = "Playlists";
        private string _title;
        private string _path;

        public string Category
        {
            get
            {
                return _category;
            }
        }

        /// <summary>
        /// Gets or sets the object's title.
        /// </summary>
        public string Title
        {
            get
            {
                return _title;
            }
        }
        /// <summary>
        /// Gets or sets the object's path.
        /// </summary>
        public string Path
        {
            get
            {
                return _path;
            }
        }

        /// <summary>
        /// Gets the shell representation of an object, such as
        /// <b>IShellLink</b> or <b>IShellItem</b>.
        /// </summary>
        /// <returns></returns>
        public object GetShellRepresentation()
        {

            return null;
        }

        public PlaylistJlDestination(string title, string path)
        {
            _title = title;
            _path = path;
        }
    }
}
