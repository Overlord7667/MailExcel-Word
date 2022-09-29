using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using MailExcel.Model; 

namespace MailExcel.ViewModel
{
    class MailnVM : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        ObservableCollection<string> _emailList;
        string _email;
        ExcelModel _excelModel;
        ExcelGenerator _excelGenerator;
        WordModel _wordModel;
        WordGenerator _wordGenerator;
        MailMessager messager;

        public ExcelModel ExcelModelTemplate
        {
            get { return _excelModel; }
            set
            {
                _excelModel = value;
                Notify("ExcelModelTemplate");
            }
        }
        public WordModel WordModelTemplate
        {
            get { return _wordModel; }
            set
            {
                _wordModel = value;
                Notify("ExcelModelTemplate");
            }
        }
        public ButtonCommand AddButtonClick
        {
            get
            {
                return new ButtonCommand(()=> {
                    EmailList.Add(Email);
                    Email = "";
                },
                    ()=> { return Email.Contains('@') && Email.Contains('.'); });
            }
        }
        public string Email
        {
            get { return _email; }
            set
            {
                _email = value;
                Notify("Email");
            }
        }
        public ObservableCollection<string> EmailList
        {
            get { return _emailList; }
            set
            {
                _emailList = value;
                Notify("EmailList");
            }
        }
        public ButtonCommand SendButtonClick
        {
            get
            {
                return new ButtonCommand(
                    () =>
                    {
                        _excelGenerator.Generate();
                        _wordGenerator.Generate1();
                        messager.SendMail();
                        MessageBox.Show("Отправлено!");
                    },
                    () =>
                    {
                        return EmailList.Count != 0 && _excelModel.CellsCount != 0
                        && _excelModel.RandomMax != 0 && _wordModel.Text.GetHashCode() !=0;
                    });
            }
        }
        public MailnVM()
        {
            EmailList = new ObservableCollection<string>();
            Email = "";
            ExcelModelTemplate = new ExcelModel();
            _excelGenerator = new ExcelGenerator(_excelModel);
            WordModelTemplate = new WordModel();
            _wordGenerator = new WordGenerator(_wordModel);
            messager = new MailMessager(EmailList);
        }
        void Notify(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
