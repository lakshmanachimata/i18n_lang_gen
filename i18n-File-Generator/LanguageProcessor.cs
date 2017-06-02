using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace i18n_File_Generator
{
    public class LanguageProcessor
    {
        private string _EngStr = "";
        public string EngStr
        {
            get
            {
                return _EngStr;
            }
            set
            {
                _EngStr = value;
                AddLanguageStr("English", _EngStr);
            }
        }
        public string HashID = "";
        public Dictionary<string, string> _LangTokens = new Dictionary<string, string>();

        private bool _isMatchFound = false;
        public bool isMacthFound
        {
            get
            {
                return _isMatchFound;
            }
        }
        public LanguageProcessor(string _hashID)
        {
            HashID = _hashID;
        }
        

        public void AddLanguageStr(string Language,string LangStr)
        {
            lock (_LangTokens)
            {
                if (!_LangTokens.ContainsKey(Language))
                {
                    _LangTokens.Add(Language, LangStr);
                }
            }
            _isMatchFound = true;
        }

        public string GetLanguageStr(string Language)
        {
            try
            {
                return _LangTokens[Language];
            }
            catch { }
            return "";
        }

    }

    
    public class LanguageToken
    {
        public string Language;
        public string LangStr;

        public LanguageToken()
        {

        }

    }

    public class LanguageRef
    {

        public static readonly Dictionary<string, LanguageRef> _LangReferences = new Dictionary<string, LanguageRef>();

        static LanguageRef()
        {
            _LangReferences.Add("English", new LanguageRef("English", "en"));
            _LangReferences.Add("German", new LanguageRef("German", "de"));
            _LangReferences.Add("Dutch", new LanguageRef("Dutch", "nl"));
            _LangReferences.Add("French", new LanguageRef("French", "fr"));
            _LangReferences.Add("Italian", new LanguageRef("Italian", "it"));
            _LangReferences.Add("Russian", new LanguageRef("Russian", "ru"));
            _LangReferences.Add("Polish", new LanguageRef("Polish", "pl"));
            _LangReferences.Add("Swedish", new LanguageRef("Swedish", "sv"));
            _LangReferences.Add("Norwegian", new LanguageRef("Norwegian", "no"));
            _LangReferences.Add("Finnish", new LanguageRef("Finnish", "fi"));
            _LangReferences.Add("Spanish", new LanguageRef("Spanish", "es"));
            _LangReferences.Add("Chinese", new LanguageRef("Chinese", "zh"));
            _LangReferences.Add("Danish", new LanguageRef("Danish", "da"));
            _LangReferences.Add("Turkish", new LanguageRef("Turkish", "tr"));
            _LangReferences.Add("Portuguese", new LanguageRef("Portuguese", "pt"));
            _LangReferences.Add("Greek", new LanguageRef("Greek", "el"));
            _LangReferences.Add("Slovak", new LanguageRef("Slovak", "sk"));
            _LangReferences.Add("Czech", new LanguageRef("Czech", "cs"));
        }

        public string LanguageName;
        public string LanguageCode;

        public LanguageRef(string _LanguageName,string _LanguageCode)
        {
            LanguageName = _LanguageName;
            LanguageCode = _LanguageCode;
        }

    }

}

