using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Word;

namespace MailExcel.Model
{
    class WordGenerator
    {
        WordModel _wordModel;

        public WordGenerator(WordModel template)
        {
            _wordModel = template;
        }

        public void Generate1()
        {
               Document doc = null;
               try
               {
                   string path = "D:\\1.docx";
                   Microsoft.Office.Interop.Word.Application application
                       = new Microsoft.Office.Interop.Word.Application();
                   doc = application.Documents.Open(path);
                   doc.Activate();
                   Bookmarks bookmarks = doc.Bookmarks;
                   int n = bookmarks.Count;
                List<string> text = new List<string>();
                text.Add("" + text + "");
                int i = 0;
                Range range;
                foreach (Bookmark bookmark in bookmarks)
                {
                    range = bookmark.Range;
                    range.Text = text[i++];
                }
                doc.Close();
               }
               catch (Exception ex)
               {
                    MessageBox.Show(ex.Message);
                    doc.Close();
               }
        }
    }
}
