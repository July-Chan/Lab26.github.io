using System;
using System.Collections.Generic;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace Lab26
{
    internal class WordHelper
    {
        private FileInfo _fileInfo;

        // Метод, що перевіряє наявність файлу
        public WordHelper(string fileName)
        {
            if (File.Exists(fileName))
            {
                _fileInfo = new FileInfo(fileName);
            }
            else
            {
                throw new ArgumentException("Файл не знайдено");
            }
        }

        // Метод, що створює документ на основі шаблону
        internal string Process(Dictionary<string, string> items)
        {
            Word.Application app = null;
            try
            {
                app = new Word.Application();
                Object file = _fileInfo.FullName;
                Object missing = Type.Missing;

                Word.Document doc = app.Documents.Open(ref file, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                foreach (var item in items)
                {
                    Word.Find find = app.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;

                    Object wrap = Word.WdFindWrap.wdFindContinue;
                    Object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(
                        FindText: item.Key,
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: item.Value,
                        Replace: replace);
                }

                Object newFileName = Path.Combine(_fileInfo.DirectoryName, DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + _fileInfo.Name);
                doc.SaveAs2(ref newFileName);
                doc.Close(ref missing, ref missing, ref missing);

                return newFileName.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
            finally
            {
                if (app != null)
                {
                    app.Quit();
                }
            }
        }

        // Метод для пошуку і заміни у останньо створеному документі
        internal bool FindAndReplace(string findText, string replaceText)
        {
            Word.Application app = null;
            try
            {
                app = new Word.Application();
                Object file = _fileInfo.FullName;
                Object missing = Type.Missing;

                Word.Document doc = app.Documents.Open(ref file, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                foreach (Word.Range range in doc.StoryRanges)
                {
                    Word.Find find = range.Find;
                    find.Text = findText;
                    find.Replacement.Text = replaceText;

                    Object wrap = Word.WdFindWrap.wdFindContinue;
                    Object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(
                        FindText: findText,
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: replaceText,
                        Replace: replace);
                }

                Object newFileName = Path.Combine(_fileInfo.DirectoryName, DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_Replaced_" + _fileInfo.Name);
                doc.SaveAs2(ref newFileName);
                doc.Close(ref missing, ref missing, ref missing);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
            finally
            {
                if (app != null)
                {
                    app.Quit();
                }
            }
        }
    }
}
