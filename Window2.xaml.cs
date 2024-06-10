using Microsoft.Win32;
using System;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace WordLekcia
{
    public partial class Window2 : Window
    {
        private string selectedFilePath;

        public Window2()
        {
            InitializeComponent();
        }

        private void SelectFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                selectedFilePath = openFileDialog.FileName;
                MessageBox.Show("Выбран файл: " + selectedFilePath);
            }
        }

        private void SendButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFilePath))
            {
                MessageBox.Show("Пожалуйста, выберите файл для отправки.");
                return;
            }

            TextRange range = new TextRange(MessegeRtb.Document.ContentStart, MessegeRtb.Document.ContentEnd);

            MailMessage messagev = new MailMessage(From.Text, To.Text, Subject.Text, range.Text)
            {
                IsBodyHtml = true
            };
            messagev.Attachments.Add(new Attachment(selectedFilePath));
            string server = "smtp.mail.ru";
            string servergm = "smtp.mail.ru";
            SmtpClient smtpclient;

            if (server == "smtp.yandex.ru")
            {
                smtpclient = new SmtpClient("smtp.yandex.ru", 587);
            }
            else if (servergm == "smtp.gmail.com")
            {
                smtpclient = new SmtpClient("smtp.gmail.com", 587);
            }
            else if (server == "smtp.rambler.ru")
            {
                smtpclient = new SmtpClient("smtp.rambler.ru", 465);
            }
            else
            {
                smtpclient = new SmtpClient("smtp.mail.ru", 587);
            }

            try
            {
                smtpclient.Send(messagev);
                MessageBox.Show("Сообщение отправлено.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка отправки сообщения: " + ex.Message);
            }
            smtpclient.Credentials = new NetworkCredential(From.Text, Pass.Password);
            smtpclient.EnableSsl = true;
            smtpclient.Send(messagev);
        }
    }
}
