using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.IO;
using System.Windows.Forms;
using System.Xml;

namespace ERespondent.Config
{
    class Configuration
    {
        public class ConfigKey
        {
            private string _id;
            public ConfigKey(string id)
            {
                this._id = id;
                Configuration.Register(this._id, this);
            }

            public string GetValue() 
            {
                return Configuration.GetKeyValue(this._id);
            }
            protected void SetValue(string value)
            {
                Configuration.SetKeyValue(this._id, value);
            }
        }

        //библиотека ключей
        private static Dictionary<String, ConfigKey> _libKey = new Dictionary<String, ConfigKey>();
        //хранилище данных
        private static Dictionary<String, String> _keyValueStorage = new Dictionary<String, String>();      
        
        public Configuration()
        { }

        public void Init()
        { }

        public void Exit()
        { }

        /// <summary>
        /// Регистрация нового параметра ключ-значение
        /// </summary>
        /// <param name="id"></param>
        /// <param name="key"></param>
        private static void Register(String id, ConfigKey key)
        {
            _libKey.Add(id, key);
            _keyValueStorage.Add(id, key.GetValue());
        }

        /// <summary>
        /// Возвращает занчение ключа
        /// </summary>
        /// <param name="id">Id ключа</param>
        /// <returns>Соответствующее ключу значение</returns>
        private static string GetKeyValue(string id)
        {
            string result = null;
            if (_keyValueStorage.TryGetValue(id, out result ))
            {
                return result;
            }
            return result;
        }

        /// <summary>
        /// Установка нового значения ключа
        /// </summary>
        /// <param name="id">id ключа</param>
        /// <param name="value">новое значение</param>
        private static void SetKeyValue(string id, string value)
        {
            _keyValueStorage[id] = value;
        }

        private void Load()
        {
            //получаем путь, откуда было запущено приложение
            string filePath = Application.StartupPath + @"\App.config";
            if (!File.Exists(filePath))
            {
                throw new IOException("Файл настроек не существует или путь указан неверно!");
            }
            XmlDocument document = new XmlDocument();
            document.Load(filePath);

            XmlNode root = document.FirstChild;
            if (null == root || root.Name != "keys")
            {
                foreach (XmlNode child in root.ChildNodes)
                    _keyValueStorage[child.Attributes["id"].InnerText] = child.Attributes["val"].InnerText;
            }

        }

    }
}
