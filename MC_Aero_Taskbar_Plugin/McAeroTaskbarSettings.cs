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
        public Dictionary<string, HashSet<string>> _pinnedPlaylists;

        public HashSet<string> GetPinnedPlaylists(string library)
        {
            if (_pinnedPlaylists == null) _pinnedPlaylists = new Dictionary<string, HashSet<string>>();
            if (!_pinnedPlaylists.ContainsKey(library)) _pinnedPlaylists[library] = new HashSet<string>(StringComparer.Ordinal);
            return _pinnedPlaylists[library];
        }
    }
}
