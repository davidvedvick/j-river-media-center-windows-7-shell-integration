using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml.Serialization;

namespace MC_Aero_Taskbar_Plugin
{
    public class AppSettings<T> where T : new()
    {
        private const string DEFAULT_FILENAME = "settings.jsn";
        private string _fileName;

        private T _settings;

        public T Settings
        {
            get
            {
                if (_settings == null) _settings = new T();
                return _settings;
            }
        }

        public AppSettings(string fileName = DEFAULT_FILENAME)
        {
            _fileName = fileName;
            _settings = Load(_fileName);
        }

        public void Save()
        {
            Save(_settings, _fileName);
        }

        public static void Save(T pSettings, string fileName = DEFAULT_FILENAME)
        {
            using (StreamWriter sw = new StreamWriter(fileName, false))
            {
                sw.Write(JsonConvert.SerializeObject(pSettings));
            }
        }

        public static T Load(string fileName = DEFAULT_FILENAME)
        {
            T t = new T();
            if (File.Exists(fileName))
            {
                t = JsonConvert.DeserializeObject<T>(File.ReadAllText(fileName));
            }
            return t;
        }
    }
}
