using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using System.Text;

namespace i18n_File_Generator
{
    public partial class Form1 : Form
    {
        private string i18n_template = @"import {Injectable} from '@angular/core';

@Injectable()
export class i18n_Lang_Defs
        {
            public userLanguageCode:string = 'en';
            public langTranslations:any = []
            constructor()
            {
                this.userLanguageCode = this.getUserLanguage(navigator.language);
                this.loadStaticData();
            }
    getUserLanguage(_langStr:string)
            {
              console.log(_langStr);
                if(_langStr!=''){
                    if(_langStr.indexOf('-')>0){
                        return _langStr.substring(0,_langStr.indexOf('-'));
                    }
                    else if(_langStr.length==2){
                      return _langStr;
                    }
                }
                return 'en';
            }
    getTranslations()
            {
                if(this.langTranslations[this.userLanguageCode])
                    return this.langTranslations[this.userLanguageCode];
                else
                    return this.langTranslations['en'];
            }
            private loadStaticData()
            {
                   [TEMPLATE_TOKEN]                
            }

        }
";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (ofdDataExcel.ShowDialog(this) == DialogResult.OK)
            {
                txtDataExcelFile.Text = ofdDataExcel.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (fbdExportDir.ShowDialog(this) == DialogResult.OK)
            {
                txtExportDir.Text = fbdExportDir.SelectedPath;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (ofdENSampleFile.ShowDialog(this) == DialogResult.OK)
            {
                txtENSample.Text = ofdENSampleFile.FileName;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(txtDataExcelFile.Text, false, true);

            LogInfo("");

            try
            {
                this.Enabled = false;
                if (ProcessData(xlWorkbook))
                {
                    LogInfo("SUCCESS");
                }
                else
                {
                    LogInfo("FAILED");
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show("Error Occurred, Details : " + exc.ToString());
            }
            finally
            {
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
            this.Enabled = true;
        }

        private bool ProcessData(Excel.Workbook xlWorkbook)
        {
            int workbookCount = xlWorkbook.Sheets.Count;
            Dictionary<string, LanguageProcessor> _wordTokens = new Dictionary<string, LanguageProcessor>();
            _wordTokens = LoadFromSampleFile();
            LogInfo("Total Words Found in Project File : " + _wordTokens.Count.ToString());

            List<string> LanguagesFound = new List<string>();
            for (int sheetIndex = 2; sheetIndex <= workbookCount; sheetIndex++)
            {
                #region get workbook reference
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[sheetIndex];
                //Excel.Range xlRange = xlWorksheet.get_Range("A1", "AA88000"); 
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                List<string> Languages = new List<string>();
                #endregion
                #region verify languages
                for (int colIndex = 3; colIndex <= colCount; colIndex++)
                {
                    string language = xlRange.Cells[1, colIndex].Text.ToString();
                    if (language.Trim() != "")
                    {
                        if (LanguageRef._LangReferences.ContainsKey(language))
                        {
                            LanguagesFound.Add(language);
                            Languages.Add(language);
                        }
                        else
                        {
                            string errMessage = "Language : " + language + ", doesn't exist in definition, Please check and try again.";
                            LogInfo(errMessage);
                            MessageBox.Show(errMessage);
                            return false;
                        }
                    }
                }

                LogInfo("Total Languages Found in sheet(" + (sheetIndex - 1).ToString() + ") : " + Languages.Count.ToString());
                #endregion
                #region process word tokens sample document and excel source
                bool isTokensMismatch = false;
                for (int rowIndex = 2; rowIndex <= rowCount; rowIndex++)
                {
                    //object currWord = xlRange.Cells[rowIndex, 3].Value2;
                    object currWord = xlRange.Cells[rowIndex, 3].Text;
                    if (currWord == null)
                        continue;
                    string currWordStr = currWord.ToString().Trim();
                    if (rowIndex > 89)
                    {
                        int brk = 0;
                    }

                    foreach (string hashID in _wordTokens.Keys)
                    {
                        LanguageProcessor currWordProcessor = _wordTokens[hashID];

                        string wordStr = currWordProcessor.EngStr;

                        
                        if (wordStr == currWordStr)
                        {
                            LoadLangProcessorData(currWordProcessor, rowIndex, colCount, xlRange);
                        }
                        else if(wordStr.ToLower() == currWordStr.ToLower())
                        {
                            string message = "Mismatch in tokens please check, Angular value  = \"" + wordStr + "\", Excel value = \"" + currWordStr + "\" ";
                            LogInfo(message);
                            isTokensMismatch = true;
                        }
                    }
                }
                if (isTokensMismatch)
                {
                    MessageBox.Show("Tokens Mismatch, please check log");
                    return false;
                }
                //foreach (string hashID in _wordTokens.Keys)
                //{
                //    LanguageProcessor currWordProcessor = _wordTokens[hashID];

                //    string wordStr = currWordProcessor.EngStr;
                //    for (int rowIndex = 2; rowIndex <= rowCount; rowIndex++)
                //    {

                //        object currWord = xlRange.Cells[rowIndex, 3].Value2;
                //        if (currWord == null)
                //            continue;
                //        string currWordStr = currWord.ToString().Trim();

                //        if (wordStr == currWordStr)// || wordStr == currWordStr.Trim())
                //        {
                //            LoadLangProcessorData(currWordProcessor, rowIndex, colCount, xlRange);
                //        }
                //    }


                //}
                #endregion


                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
            }
            #region validate data
            foreach (string hashID in _wordTokens.Keys)
            {
                if (_wordTokens[hashID].isMacthFound == false)
                {
                    string message = "Word \"" + _wordTokens[hashID].EngStr + "\" not found in excel, operation aborted";
                    LogInfo(message);
                    MessageBox.Show(message);
                    return false;
                }

            }
            #endregion
            #region prepare the language documents
            StringBuilder builder = new StringBuilder();
            string sampleFileData = File.ReadAllText(txtENSample.Text);
            Dictionary<string, string> _FileList = new Dictionary<string, string>();
            for (int i = 0; i < LanguagesFound.Count; i++)
            {
                //if (LanguagesFound[i].ToLower() == "english")
                //continue;
                string FileName = CopyLanguageFileFromSample(LanguageRef._LangReferences[LanguagesFound[i]].LanguageCode, sampleFileData, "messages");
                _FileList.Add(LanguagesFound[i], FileName);
            }

            foreach (string language in _FileList.Keys)
            {
                UpdateLanguageFile(language, _FileList[language], _wordTokens);
            }

            #region generate i18n-data.ts
            StringBuilder build = new StringBuilder();

            foreach (string language in _FileList.Keys)
            {
                string langCode = LanguageRef._LangReferences[language].LanguageCode.ToLower();
                build.AppendLine("this.langTranslations['" + langCode + "'] = `" + File.ReadAllText(_FileList[language]) + "`;");
            }

            TextWriter tw = new StreamWriter(txtExportDir.Text + "\\i18n-data.ts", false);
            tw.Write(i18n_template.Replace("[TEMPLATE_TOKEN]", build.ToString()));
            tw.Close();
            #endregion

            //foreach (string hashID in _wordTokens.Keys)
            //{
            //    LanguageProcessor wordToken = _wordTokens[hashID];


            //}

            #endregion
            return true;
        }



        private string CopyLanguageFileFromSample(string languageCode, string sampleFileData, string fileStartName)
        {
            string fileName = txtExportDir.Text + "\\" + fileStartName + "." + languageCode + ".xlf";
            File.WriteAllText(fileName, sampleFileData);
            return fileName;
        }

        private void LoadLangProcessorData(LanguageProcessor currWordProcessor, int rowIndex, int colCount, Excel.Range xlRange)
        {

            for (int colIndex = 4; colIndex <= colCount; colIndex++)
            {
                //try
                //{
                    string language = xlRange.Cells[1, colIndex].Text.ToString();
                    string langWordStr = "";
                    object tmpObj = xlRange.Cells[rowIndex, colIndex].Text;
                    if (tmpObj != null)
                    {
                        langWordStr = tmpObj.ToString();
                    }
                    else
                    {
                        langWordStr = currWordProcessor.EngStr;
                    }
                    if (langWordStr == "%100")
                    {
                        int brk = 0;
                    }
                    currWordProcessor.AddLanguageStr(language, langWordStr);
                //}
                //catch
                //{
                //    int brk = 0;
                //}
            }
        }

        private void UpdateLanguageFile(string language, string filepath, Dictionary<string, LanguageProcessor> wordTokens)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(filepath);

            //MessageBox.Show(doc.ChildNodes.Count.ToString());
            XmlNodeList units = doc.GetElementsByTagName("trans-unit");

            for (int i = 0; i < units.Count; i++)
            {
                string id = units.Item(i).Attributes["id"].Value;

                //string str = wordTokens[id].EngStr;

                string langSpeStr = wordTokens[id].GetLanguageStr(language);
                if (langSpeStr == "")
                {
                    LogInfo("lang Specific string not found for --> " + wordTokens[id].EngStr + " (" + language + ") <-- using default");
                    langSpeStr = wordTokens[id].EngStr;
                }

                for (int j = 0; j < units.Item(i).ChildNodes.Count; j++)
                {
                    if (units.Item(i).ChildNodes.Item(j).Name == "target")
                    {
                        units.Item(i).ChildNodes.Item(j).InnerText = langSpeStr;
                        break;
                    }
                }
            }

            doc.Save(filepath);


            doc = null;

        }

        private Dictionary<string, LanguageProcessor> LoadFromSampleFile()
        {
            string fileName = txtENSample.Text;
            Dictionary<string, LanguageProcessor> _retValue = new Dictionary<string, LanguageProcessor>();

            XPathNavigator nav;
            XPathDocument docNav;

            docNav = new XPathDocument(txtENSample.Text);
            nav = docNav.CreateNavigator();
            nav.MoveToRoot();
            XmlReader reader = nav.ReadSubtree();

            LanguageProcessor tmpCurrProcessor = null;
            string tmpHashID = "";
            bool isInsideUnit = false;
            bool duplicatesFound = false;
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    if(reader.Name== "trans-unit")
                    {
                        isInsideUnit = true;

                        tmpHashID = reader.GetAttribute("id");
                        tmpCurrProcessor = new LanguageProcessor(tmpHashID);
                    }
                    if(reader.Name== "source" && isInsideUnit)
                    {
                        //try
                        //{
                            tmpCurrProcessor.EngStr = reader.ReadElementContentAsString().Trim();
                        //}
                        //catch {
                        //    int brk = 0;
                        //}
                    }

                    if(reader.Name== "target")
                    {
                        if (!_retValue.ContainsKey(tmpCurrProcessor.HashID))
                        {
                            _retValue.Add(tmpCurrProcessor.HashID, tmpCurrProcessor);
                        }
                        else
                        {
                            duplicatesFound = true;
                            LogInfo("Warning : IGNORING DUPLICATE STRING --> " + tmpCurrProcessor.EngStr+"("+ tmpHashID+")");
                        }
                        tmpHashID = "";
                        isInsideUnit = false;
                        tmpCurrProcessor = null;
                    }
                    //MessageBox.Show(reader.Name.ToString());
                }
                
            }
            LogInfo("Total " + (duplicatesFound ? "UNIQUE" : "") + "Strings Found = " + _retValue.Count.ToString());
            return _retValue;

        }

        private void LogInfo(string info)
        {
            if (info == "")
            {
                txtLog.Text = "";
            }
            else
            {
                txtLog.Text = DateTime.Now.ToString("HH:mm:ss") + " --> " + info + "\r\n" + txtLog.Text;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                TextWriter tw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "opts.dat", false);
                tw.Write(txtDataExcelFile.Text + "," + txtENSample.Text + "," + txtExportDir.Text);
                tw.Close();
            }
            catch { }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "opts.dat"))
                {
                    string[] tokens = File.ReadAllText(AppDomain.CurrentDomain.BaseDirectory + "opts.dat").Split(new string[] { "," }, StringSplitOptions.None);

                    txtDataExcelFile.Text = tokens[0];
                    txtENSample.Text = tokens[1];
                    txtExportDir.Text = tokens[2];
                }
            }
            catch { }
        }
    }
}
