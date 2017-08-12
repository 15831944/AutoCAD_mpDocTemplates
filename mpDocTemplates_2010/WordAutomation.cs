using System;
using System.IO;
using System.Reflection;

namespace mpDocTemplates
{
    /// <summary>
    /// Cass for automating word.
    /// </summary>
    public class WordAutomation
    {
        #region Member Variables
        //Dynamic object for word application
        private dynamic _wordApplication;
        //Dynamic object for word document
        private dynamic _wordDoc;
        #endregion 

        #region Member Functions

        /// <summary>
        /// Function to create a word document.
        /// </summary>
        /// <param name="fileName">Name of word</param>
        /// <param name="isReadonly">open mode</param>
        /// <returns>If word document exist, return word doc object
        /// else return null</returns>		
        public void CreateWordDoc(object fileName, bool isReadonly)
        {
            if (File.Exists(fileName.ToString()) && _wordApplication != null)
            {
                object isVisible = false;
                object missing = Missing.Value;
                // Add document.
                _wordDoc = _wordApplication.Documents.Add(fileName, missing, missing, isVisible);
            }
        }
        /// <summary>
        /// Function to create a word application
        /// </summary>
        /// <returns>Returns a word application</returns>		
        public void CreateWordApplication()
        {
            const string message = "Не удалось создать объект Word! Убедитесь, что установлен Microsoft Office";
            var wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
            {
                throw new Exception(message);
            }
            _wordApplication = Activator.CreateInstance(wordType);
            if (_wordApplication == null)
            {
                throw new Exception(message);
            }
        }
        public void MakeWordAppVisible()
        {
            _wordApplication.Visible = true;
        }
        /// <summary>
        /// Function to close a word document.
        /// </summary>
        /// <param name="canSaveChange">Need to save changes or not</param>
        /// <returns>True if successfully closed.</returns>		
        public bool CloseWordDoc(bool canSaveChange)
        {
            var isSuccess = false;
            if (_wordDoc != null)
            {
                object saveChanges;
                if (canSaveChange)
                {
                    saveChanges = -1; // Save Changes
                }
                else
                {
                    saveChanges = 0; // No changes
                }
                _wordDoc.Close(saveChanges);
                //InvokeMember("Close",wordDocument,new object[]{saveChanges});	
                _wordDoc = null;
                isSuccess = true;
            }
            return isSuccess;
        }
        /// <summary>
        /// Function to close word application
        /// </summary>
        /// <returns>True if successfully closed</returns>		
        public bool CloseWordApp()
        {
            var isSuccess = false;
            if (_wordApplication != null)
            {
                object saveChanges = 0;
                _wordApplication.Quit(saveChanges);
                _wordApplication = null;
                isSuccess = true;
            }
            return isSuccess;
        }
        /// <summary>
        /// Function to find and replace a given text.
        /// </summary>
        /// <param name="findText">Text for finding</param>
        /// <param name="replaceText">Text for replacing</param>
        /// <returns>True if successfully replaced.</returns>
        public bool FindReplace(
            string findText, string replaceText)
        {

            var isSuccess = false;
            do
            {
                if (_wordDoc == null)
                {
                    break;
                }
                if (_wordApplication == null)
                {
                    break;
                }
                if (replaceText.Trim().Length == 0)
                {
                    break;
                }
                if (findText.Trim().Length == 0)
                {
                    break;
                }

                ReplaceRange(_wordDoc.Content, findText, replaceText);

                int rangeCount = _wordDoc.Comments.Count;
                for (var i = 1; i <= rangeCount; i++)
                {
                    ReplaceRange(_wordDoc.Comments.Item(i).Range,
                        findText, replaceText);
                }

                for (var s = 1; s <= _wordDoc.Sections.Count; s++)
                {
                    rangeCount = _wordDoc.Sections[s].Headers.Count;
                    for (var i = 1; i <= rangeCount; i++)
                    {
                        ReplaceRange(_wordDoc.Sections[s].Headers.Item(i).Range,
                            findText, replaceText);
                    }
                    rangeCount = _wordDoc.Sections[s].Footers.Count;
                    for (var i = 1; i <= rangeCount; i++)
                    {
                        ReplaceRange(_wordDoc.Sections[s].Footers.Item(i).Range,
                            findText, replaceText);
                    }
                }

                rangeCount = _wordDoc.Shapes.Count;
                for (var i = 1; i <= rangeCount; i++)
                {
                    dynamic textFrame = _wordDoc.Shapes.Item(i).TextFrame;
                    int hasText = textFrame.HasText;
                    if (hasText < 0)
                    {
                        ReplaceRange(textFrame.TextRange, findText, replaceText);
                    }
                }
                isSuccess = true;
            }
            while (false);
            return isSuccess;
        }
        /// <summary>
        /// Replace a word with in a range.
        /// </summary>
        /// <param name="range">Range to replace</param>
        /// <param name="findText">Text to find</param>
        /// <param name="replaceText">Text to replace</param>
        private void ReplaceRange(dynamic range,
            string findText, string replaceText)
        {
            object missing = Missing.Value;
            _wordDoc.Activate();
            object item = 1;
            object whichItem = 1;
            _wordDoc.GoTo(item, whichItem);
            object replaceAll = 2;
            dynamic find = range.Find;
            find.ClearFormatting();
            find.Replacement.ClearFormatting();
            find.Execute(findText, false, true,
                                missing, missing, missing, true, missing, missing
                                , replaceText, replaceAll);
        }
        #endregion
    }
}