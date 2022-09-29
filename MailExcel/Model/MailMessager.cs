using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;
using System.Collections.ObjectModel;

namespace MailExcel.Model
{
    class MailMessager
    {
        ObservableCollection<string> _emailList;
        string _path;
        string _path1;
        public MailMessager(ObservableCollection<string> emailList,string path = "D:\\1.xlsx",string path1 = "D:\\1.docx")
        {
            _emailList = emailList;
            _path = path;
            _path1 = path1;
        }
        public void SendMail()
        {
            SmtpClient smtpClient = new SmtpClient("smtp.yandex.ru", 25);
            smtpClient.Credentials = new NetworkCredential("totoshkavichys@yandex.ru",
                "vyacheslavov.yrii!");
            smtpClient.EnableSsl = true;
            foreach(string adress in _emailList)
            {
                MailMessage mailMessage = new MailMessage("totoshkavichys@yandex.ru", adress);
                mailMessage.Subject = "Отчет";
                mailMessage.Body = "Добрый день, отправляю вам отчет";
                mailMessage.Attachments.Add(new Attachment(_path));
                mailMessage.Attachments.Add(new Attachment(_path1));
                smtpClient.Send(mailMessage);
            }
        }
    }
}
