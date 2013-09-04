using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MC_Aero_Taskbar_Plugin
{
    public class McAeroTaskbarSettings
    {
        public bool EnableCoverArt;
        public bool DisplayArtistTrackName;
        public bool TrackProgress;
        public bool PlaylistProgress;
        public bool NoProgressTrack;
        public Dictionary<string, HashSet<string>> _pinnedServerPlaylists;

        public HashSet<string> GetPinnedPlaylists(string library)
        {
            if (_pinnedServerPlaylists == null) _pinnedServerPlaylists = new Dictionary<string, HashSet<string>>();
            if (!_pinnedServerPlaylists.ContainsKey(library)) _pinnedServerPlaylists[library] = new HashSet<string>(StringComparer.Ordinal);
            return _pinnedServerPlaylists[library];
        }
    }
}
