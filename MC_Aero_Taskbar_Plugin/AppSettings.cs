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
            _settings = Load(fileName);
        }

        public void Save(string fileName = DEFAULT_FILENAME)
        {
            Save(_settings, fileName);
        }

        public static void Save(T pSettings, string fileName = DEFAULT_FILENAME)
        {
            using (FileStream fs = new FileStream(fileName, File.Exists(fileName) ? FileMode.Truncate : FileMode.OpenOrCreate))
            {
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(T));
                ser.WriteObject(fs, pSettings);
            }
        }

        public static T Load(string fileName = DEFAULT_FILENAME)
        {
            T t = new T();
            if (File.Exists(fileName))
            {
                using (FileStream fs = new FileStream(fileName, FileMode.OpenOrCreate))
                {
                    DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(T));
                    t = (T)ser.ReadObject(fs);
                }
            }
            return t;
        }
    }
}
